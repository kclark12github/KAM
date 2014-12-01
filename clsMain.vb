Public Class clsMain
    Private mAborting As Boolean = False
    Private mActiveTXLevel As Integer = 0
    Private mCommandTimeout As Integer = 300
    Private mConnection As SqlClient.SqlConnection
    Private mDA As SqlClient.SqlDataAdapter
    Private mReader As StreamReader
    Private mTransaction As SqlClient.SqlTransaction
    Public Sub New()
        MyBase.New()
        mConnection = New SqlClient.SqlConnection("Application Name=KAM;Workstation ID=GGGSCP1;Data Source=GGGSCP1;Initial Catalog=KAM;Packet Size=8192;Connect Timeout=120;Pooling=False;User ID=KCLARK;Password=b2spirit;")
        mConnection.Open()

        mDA = New SqlClient.SqlDataAdapter()
        With mDA
            .DeleteCommand = New SqlClient.SqlCommand : .DeleteCommand.Connection = mConnection
            .InsertCommand = New SqlClient.SqlCommand : .InsertCommand.Connection = mConnection
            .SelectCommand = New SqlClient.SqlCommand : .SelectCommand.Connection = mConnection
            .UpdateCommand = New SqlClient.SqlCommand : .UpdateCommand.Connection = mConnection
        End With
    End Sub
#Region "Database Stuff"
    Public Sub AbortTrans()
        Try
            If mAborting Then Exit Try
            If IsNothing(mTransaction) Then Exit Try

            If Not mConnection Is mTransaction.Connection Then Throw New InvalidCastException("Attempting to Rollback a transaction associated with a different connection.")

            mAborting = True
            mTransaction.Rollback()
            mDA.SelectCommand.Transaction = Nothing
            mDA.DeleteCommand.Transaction = Nothing
            mDA.InsertCommand.Transaction = Nothing
            mDA.UpdateCommand.Transaction = Nothing
            mAborting = False
            mActiveTXLevel = 0
        Finally
            If Not IsNothing(mTransaction) Then mTransaction.Dispose() : mTransaction = Nothing
        End Try
    End Sub
    Public Function BeginTrans(Optional ByVal Level As System.Data.IsolationLevel = IsolationLevel.RepeatableRead) As SqlClient.SqlTransaction
        BeginTrans = Nothing
        Try
            If mActiveTXLevel > 0 Then Throw New Exception("Transaction is already active. Multiple or nested transactions are not permitted.")

            mTransaction = mConnection.BeginTransaction(Level)
            mDA.SelectCommand.Transaction = mTransaction
            mDA.DeleteCommand.Transaction = mTransaction
            mDA.InsertCommand.Transaction = mTransaction
            mDA.UpdateCommand.Transaction = mTransaction
            BeginTrans = mTransaction
            mActiveTXLevel += 1
        Catch ex As Exception
            Try : AbortTrans() : Finally : End Try
        End Try
    End Function
    Public Sub CommitTrans()
        Try
            If Not mConnection Is mTransaction.Connection Then Throw New Exception("Attempting to Commit a transaction associated with a different connection.")

            mTransaction.Commit()
            mDA.SelectCommand.Transaction = Nothing
            mDA.DeleteCommand.Transaction = Nothing
            mDA.InsertCommand.Transaction = Nothing
            mDA.UpdateCommand.Transaction = Nothing
            If mActiveTXLevel = 0 Then Throw New Exception("Commit operation requires an active transaction.")
            mActiveTXLevel -= 1
        Catch ex As Exception
            Try : AbortTrans() : Finally : End Try
        Finally
            mTransaction.Dispose() : mTransaction = Nothing
        End Try
    End Sub
    Public Sub ExecuteCommand(ByVal SQLsource As String, Optional ByRef RecordsAffected As Integer = 0)
        Dim cmd As New SqlClient.SqlCommand
        Try
            With cmd
                .CommandText = SQLsource
                .CommandType = CommandType.Text
                .Connection = mConnection
                .CommandTimeout = mCommandTimeout
                .Transaction = mTransaction
                RecordsAffected = .ExecuteNonQuery
            End With
        Catch ex As Exception
            AbortTrans()
            Throw New Exception(ex.Message, ex)
        End Try
    End Sub
    Public Function ExecuteScalarCommand(ByVal SQLsource As String) As Object
        Dim cmd As New SqlClient.SqlCommand
        ExecuteScalarCommand = Nothing
        Try
            With cmd
                .CommandText = SQLsource
                .CommandTimeout = mCommandTimeout
                .CommandType = CommandType.Text
                .Connection = mConnection
                If Not IsNothing(mTransaction) AndAlso mTransaction.Connection Is .Connection Then .Transaction = mTransaction
                ExecuteScalarCommand = .ExecuteScalar
            End With
        Catch ex As Exception
            AbortTrans()
            Throw New Exception(ex.Message, ex)
        Finally
            cmd.Dispose() : cmd = Nothing
        End Try
    End Function
    Public Function OpenDataSet(SQLSource As String, Optional ByRef RecordCount As Integer = -1) As DataSet
        OpenDataSet = New DataSet
        With mDA
            .SelectCommand.CommandText = SQLSource
            .SelectCommand.CommandType = CommandType.Text
            .SelectCommand.Connection = mConnection
            If Not IsNothing(mTransaction) AndAlso mTransaction.Connection Is .SelectCommand.Connection Then .SelectCommand.Transaction = mTransaction
            .SelectCommand.CommandTimeout = mCommandTimeout
            .FillSchema(OpenDataSet, SchemaType.Source)
            'FillSchema doesn't seem to want to use the default constraint from the database, so roll our own default here...
            'initColumnDefaults(.FillSchema(OpenDataSet, SchemaType.Source)(0).Columns)
            RecordCount = .Fill(OpenDataSet)
        End With
    End Function
#End Region

    Private Function AddKey(root As String, name As String) As Integer
        AddKey = Me.GetKey(name) : If AddKey > 0 Then Exit Function
        Dim SQL As String = String.Format("Insert Into [Key]([root],[name]) Values('{0}','{1}');", root, name.Replace("'", "''"))
        BeginTrans()
        ExecuteScalarCommand(SQL) : AddKey = ExecuteScalarCommand("Select @@IDENTITY;")
        CommitTrans()
    End Function
    Private Sub AddValue(key As Integer, name As String, value As String)
        Dim SQL As String = String.Format("Insert Into [Value]([key],[name],[value]) Values({0},'{1}','{2}');", key, name.Replace("'", "''"), value.Replace("'", "''"))
        BeginTrans()
        ExecuteScalarCommand(SQL)
        CommitTrans()
    End Sub
    Private Function DecodeValue(value As String, ByRef raw As String) As String
        DecodeValue = value
        While DecodeValue.EndsWith("\")
            DecodeValue = DecodeValue.Substring(0, DecodeValue.Length - 1)
            Dim buffer As String = mReader.ReadLine().Trim
            raw = String.Format("{0}{1}{2}", raw, vbCrLf, buffer)
            DecodeValue &= buffer
        End While
        If DecodeValue.ToLower.StartsWith("hex(") Then
            'Now: DecodeValue should be the concatenated series of Unicode characters, and...
            '     raw should be the entire block as it appeared in the file...
            Dim tempValue As String = DecodeValue.Split(":")(1)
            Dim chars() As String = tempValue.Split(",") : tempValue = ""
            For i As Integer = 0 To chars.Length - 1
                If chars(i) <> "" AndAlso chars(i) <> "00" Then tempValue &= Chr(CInt(String.Format("&H{0}", chars(i))))
            Next i
            DecodeValue = tempValue
        End If
    End Function
    Private Function GetKey(name As String) As Integer
        Dim SQL As String = String.Format("Select [id] From [Key] Where [name]='{0}';", name.Replace("'", "''"))
        GetKey = ExecuteScalarCommand(SQL)
    End Function
    Public Function ParseString(Source As String, Token As Integer, Delimiter As String, Optional ByVal Encapsulator As String = "") As String
        ParseString = ""
        Try
            If Source.Length = 0 Then Exit Try
            If Token <= 0 Then Throw New ArgumentException("Token must be at least One.")
            Dim sPos As Integer = 0 'Start Position
            Dim strTemp As String = ""
            Dim iToken As Integer = 0
            While True
                strTemp = ""
                If sPos > Source.Length - 1 Then Exit While
                Dim dPos As Integer = Source.IndexOf(Delimiter, sPos) 'Delimiter Position
                If Encapsulator.Length > 0 Then
                    Dim sPtr As Integer = Source.IndexOf(Encapsulator, sPos) 'Start Pointer
                    Dim ePtr As Integer = Source.IndexOf(Encapsulator, sPtr + 1) 'End Pointer
                    If dPos > sPtr And dPos < ePtr Then dPos = Source.IndexOf(Delimiter, ePtr)
                End If
                If dPos < sPos Then dPos = Source.Length + Delimiter.Length
                If dPos = 0 Then Exit While
                strTemp = Source.Substring(sPos, dPos - sPos)
                sPos = dPos + Delimiter.Length
                iToken += 1 : If iToken = Token Then Exit While
            End While
            If Encapsulator.Length > 0 AndAlso strTemp.StartsWith(Encapsulator) AndAlso strTemp.EndsWith(Encapsulator) Then strTemp = strTemp.Substring(1, strTemp.Length - 2)
            ParseString = strTemp
        Finally
        End Try
    End Function
    Private Sub UpdateKey(id As Integer, raw As String)
        Dim SQL As String = String.Format("Update [Key] Set [raw]='{1}' Where [id]={0};", id, raw.Replace("'", "''"))
        BeginTrans()
        ExecuteScalarCommand(SQL)
        CommitTrans()
    End Sub
    Private Sub UpdateValue(id As Integer, key As Integer, value As String)
        Dim SQL As String = String.Format("Update [Value] Set [key]={0},[value]='{1}' Where [id]={3};", New Object() {key, value.Replace("'", "''"), id})
        BeginTrans()
        ExecuteScalarCommand(SQL)
        CommitTrans()
    End Sub
    Public Sub ProcessFile(FileName As String)
        Dim key As String = ""
        Dim keyID As Integer = 0
        Dim root As String = ""
        Dim raw As String = ""

        Try
            mReader = My.Computer.FileSystem.OpenTextFileReader(FileName, System.Text.Encoding.Unicode)
            While Not mReader.EndOfStream
                Dim buffer As String = mReader.ReadLine().Trim
                Try
                    If buffer = "" OrElse buffer = "Windows Registry Editor Version 5.00" Then Exit Try
                    Select Case buffer.Substring(0, 1)
                        Case "["
                            'First finish-up previous key before beginning processing for the next guy...
                            If keyID <> 0 Then UpdateKey(keyID, raw) : raw = ""

                            'OK, now move on...
                            key = buffer.Substring(1, buffer.Length - 2)    'Strip brackets
                            Console.WriteLine(key)
                            If root = "" Then root = key : Exit Try
                            Dim keypath() As String = key.Split("\") : If keypath(0) <> root Then Throw New RootMismatchException
                            keyID = Me.AddKey(root, key)
                            raw = buffer
                        Case Else
                            raw = String.Format("{0}{1}{2}", raw, vbCrLf, buffer)
                            'TODO: We're not handling embedded quotes properly by using .Split("=")...
                            Dim valuePath() As String = buffer.Split("=")
                            Dim valueName As String = valuePath(0)
                            Dim valueValue As String = DecodeValue(valuePath(1), raw)    'Note: DecodeValue will update raw...
                            AddValue(keyID, valueName, valueValue)
                    End Select
                Catch ex As IndexOutOfRangeException
                Finally
                End Try
            End While
        Finally
            If Not IsNothing(mReader) Then mReader.Close() : mReader = Nothing
        End Try
    End Sub
End Class
Public Class RootMismatchException
    Inherits Exception
End Class