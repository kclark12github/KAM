Public Class clsMain
    Private mAborting As Boolean = False
    Private mActiveTXLevel As Integer = 0
    Private mCommandTimeout As Integer = 300
    Private mConnection As SqlClient.SqlConnection
    Private mDA As SqlClient.SqlDataAdapter
    Private mReader As StreamReader
    Private mWriter As StreamWriter
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
#Region "New Code"
    ''' <summary>Parses a given string by delimiter</summary>
    ''' <param name="Source">String to work on</param>
    ''' <param name="Delimiter">Token delimiter</param>
    ''' <param name="Encapsulator">Optional: Allows for tokens to return strings encapsulated with "Delimiter" characters</param>
    ''' <returns>Returns string array</returns>
    Public Overloads Function ParseStr(Source As String, Delimiter As String, Optional ByVal Encapsulator As String = "") As String()
        Dim delim As String = Delimiter
        ParseStr = New String() {}
        If IsNothing(Source) OrElse Source.Length = 0 Then Throw New ArgumentException("Work string must be specified.")
        If IsNothing(Delimiter) OrElse Delimiter = "" Then Throw New ArgumentException("Delimiter must be specified.")
        If IsNothing(Encapsulator) Then Encapsulator = ""

        If Delimiter.Length > 1 OrElse (Encapsulator.Length > 0 AndAlso Source.IndexOf(Encapsulator) > -1) Then
            'Strategy: Replace all occurrences of Delimiter (not encapsulated by Encapsulator) with a
            '          substitute delimiter which can be later used in a String.Split operation.
            Dim cntEncap As Integer = 0
            delim = Chr(1)
            Dim sPos As Integer = 0
            While sPos < Source.Length
                If sPos + Encapsulator.Length < Source.Length AndAlso Source.Substring(sPos, Encapsulator.Length) = Encapsulator Then
                    cntEncap += 1 : sPos += Encapsulator.Length
                ElseIf sPos + Delimiter.Length < Source.Length AndAlso Source.Substring(sPos, Delimiter.Length) = Delimiter AndAlso cntEncap Mod 2 = 0 Then
                    Source = String.Format("{0}{1}{2}", Source.Substring(0, sPos), delim, Source.Substring(sPos + Delimiter.Length)) : sPos += delim.Length
                Else
                    sPos += 1
                End If
            End While
        End If
        Return Source.Split(delim)
    End Function
    ''' <summary>Retrieve specified token of string</summary>
    ''' <param name="Source">String to work on</param>
    ''' <param name="TokenNum">Returns specified token in string</param>
    ''' <param name="Delimiter">Token delimiter</param>
    ''' <param name="Encapsulator">Optional: Allows for tokens to return strings encapsulated with "Delimiter" characters</param>
    ''' <param name="Preserve">Optional: Preserves encapsulating characters when token is encapsulated</param>
    ''' <returns>Returns string token.  If none is found, will return ""</returns>
    Public Overloads Function ParseStr(Source As String, TokenNum As Integer, Delimiter As String, Optional ByVal Encapsulator As String = "", Optional Preserve As Boolean = False) As String
        ParseStr = ""
        If TokenNum < 1 Then Throw New ArgumentException("TokenNum must be greater than zero.")
        Dim Tokens() As String = Me.ParseStr(Source, Delimiter, Encapsulator)
        If TokenNum <= Tokens.Length Then ParseStr = Tokens(TokenNum - 1)
        If Not Preserve AndAlso Encapsulator.Length > 0 AndAlso ParseStr.StartsWith(Encapsulator) AndAlso ParseStr.EndsWith(Encapsulator) Then ParseStr = ParseStr.Substring(1, ParseStr.Length - 2)
    End Function
    ''' <summary>Retrieve specified token of string</summary>
    ''' <param name="strWork">String to work on</param>
    ''' <param name="intTokenNum">Returns specified token in string</param>
    ''' <param name="strDelimitChr">Token delimiter</param>
    ''' <param name="strEncapChr">Optional: Allows for tokens to return strings encapsulated with "strDelimitChr" characters</param>
    ''' <returns>Returns string token.  If none is found, will return ""</returns>
    Public Overloads Function OldParseStr(ByVal strWork As String, ByVal intTokenNum As Short, ByVal strDelimitChr As String, Optional ByVal strEncapChr As String = "") As String
        Dim intSPos As Integer = 0                'Start Position
        Dim intDPos As Integer = 0                'Delimiter Position
        Dim intSPtr As Integer = 0                'Start Pointer
        Dim intEPtr As Integer = 0                'End Pointer

        OldParseStr = ""
        If IsNothing(strWork) OrElse strWork = "" Then Throw New ArgumentException("Work string must be specified.")
        If intTokenNum <= 0 Then Throw New ArgumentException("TokenNum must be greater than zero.")
        If IsNothing(strDelimitChr) OrElse strDelimitChr = "" Then Throw New ArgumentException("Delimiter must be specified.")
        If IsNothing(strEncapChr) Then strEncapChr = ""

        Dim intCurrentTokenNum As Short = 0S
        Dim intWorkStrLen As Integer = strWork.Length
        Dim intEncapStatus As Boolean = CBool(strEncapChr.Length > 0)
        Dim intDelimitLen As Integer = strDelimitChr.Length
        If intWorkStrLen = 0 Or intSPos > intWorkStrLen Then Exit Function

        Dim strTemp As String = ""
        While True
            strTemp = ""
            If intSPos > intWorkStrLen Then Exit While
            intDPos = strWork.IndexOf(strDelimitChr, intSPos) 'intDPos = InStr(intSPos, strWork, strDelimitChr)
            If intEncapStatus Then
                intSPtr = strWork.IndexOf(strEncapChr, intSPos) 'intSPtr = InStr(intSPos, strWork, strEncapChr)
                intEPtr = strWork.IndexOf(strEncapChr, intSPtr + 1) 'intEPtr = InStr(intSPtr + 1, strWork, strEncapChr)
                If intDPos > intSPtr And intDPos < intEPtr Then intDPos = strWork.IndexOf(strDelimitChr, intEPtr) 'intDPos = InStr(intEPtr, strWork, strDelimitChr)
            End If

            If intDPos < intSPos Then intDPos = intWorkStrLen + intDelimitLen

            If intDPos <= 0 Then Exit While
            strTemp = strWork.Substring(intSPos, Math.Min(strWork.Length - intSPos, intDPos - intSPos)) 'strTemp = Mid(strWork, intSPos, intDPos - intSPos)
            intSPos = intDPos + intDelimitLen
            intCurrentTokenNum += 1 : If intCurrentTokenNum = intTokenNum Then Exit While
        End While
        If intEncapStatus Then
            'ParseStr = ReplaceCS(strTemp, strEncapChr, "", OpMode.StringBinaryCompare)
            If strTemp.StartsWith(strEncapChr) AndAlso strTemp.EndsWith(strEncapChr) Then strTemp = strTemp.Substring(1, strTemp.Length - 2)
        End If
        OldParseStr = strTemp
    End Function
    ''' <summary>Counts number of tokens in a string</summary>
    ''' <param name="Source">String to work on</param>
    ''' <param name="Delimiter">String Delimiter</param>
    ''' <param name="Encapsulator">Optional: Allows for tokens to return strings encapsulated with "strDelimiter" characters</param>
    ''' <returns>Number of tokens found</returns>
    Public Overloads Function TokenCount(ByVal Source As String, ByVal Delimiter As String, Optional ByVal Encapsulator As String = "") As Integer
        TokenCount = 0
        If IsNothing(Source) OrElse Source.Length = 0 Then Exit Function
        If IsNothing(Delimiter) OrElse Delimiter = "" Then Exit Function
        If IsNothing(Encapsulator) Then Encapsulator = ""

        Dim Tokens() As String = Me.ParseStr(Source, Delimiter, Encapsulator)
        TokenCount = Tokens.Length
        For i As Integer = Tokens.GetUpperBound(0) To 0 Step -1
            If Tokens(i).Length = 0 Then TokenCount -= 1 Else Exit For
        Next i
    End Function
    Public Function OldTokenCount(ByVal strWork As String, ByVal strDelimiter As String, Optional ByVal strEncapChr As String = "") As Integer
        Dim intDPos As Integer                'Delimiter Position
        Dim intSPtr As Integer                'Start Pointer
        Dim intEPtr As Integer                'End Pointer
        Dim intWorkStrLen As Integer = Len(strWork)
        Dim intEncapStatus As Integer
        Dim intSPos As Integer = 1            'Start Position
        Dim strTemp As String = ""
        Dim intDelimitLen As Integer = Len(strDelimiter)
        OldTokenCount = 0
        If Len(strEncapChr) Then intEncapStatus = Len(strEncapChr)
        If intWorkStrLen = 0 Or intSPos > intWorkStrLen Then Exit Function 'Try

        While True
            strTemp = ""
            If intSPos > intWorkStrLen Then Exit While
            intDPos = InStr(intSPos, strWork, strDelimiter)
            If intEncapStatus Then
                intSPtr = InStr(intSPos, strWork, strEncapChr)
                intEPtr = InStr(intSPtr + 1, strWork, strEncapChr)
                If intDPos > intSPtr And intDPos < intEPtr Then intDPos = InStr(intEPtr, strWork, strDelimiter)
            End If

            If intDPos < intSPos Then intDPos = intWorkStrLen + intDelimitLen

            If intDPos = 0 Then Exit While
            strTemp = Mid(strWork, intSPos, intDPos - intSPos)
            intSPos = intDPos + intDelimitLen
            OldTokenCount += 1
        End While
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
        Dim Tokens() As String
        Dim delim As String = Delimiter
        ParseString = ""
        If Source.Length = 0 OrElse Token < 1 Then Exit Function
        If Source.IndexOf(Encapsulator) > -1 Then
            'Before getting started, check for escaped encapsulating characters...
            Source = Source.Replace(String.Format("\{0}", Encapsulator), String.Format("{0}", Encapsulator)) 'Should work for embedded double-quotes anyway...
            Dim cntEncap As Integer = 0
            delim = Chr(1)
            For i As Integer = 0 To Source.Length - 1
                Select Case Source.Substring(i, 1)
                    Case Encapsulator : cntEncap += 1
                    Case Delimiter : If cntEncap Mod 2 = 0 Then Mid(Source, i + 1, 1) = delim 'Cannot use .Substring to assign value
                End Select
            Next i
        End If
        Tokens = Source.Split(delim) : If Token <= Tokens.Length Then ParseString = Tokens(Token - 1)
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
    Public Sub OutputFile(FileName As String)
        Try
            Dim SQL As String = _
                "Select [Key].[name],[Key].[raw],[Value].[name],[Value].[value] " & _
                "From [Key] " & _
                    "Inner Join [Value] On [Value].[key]=[Key].[id] " & _
                "Where Lower([Key].[name]) Like '%adobe%' " & _
                    "Or Lower([Value].[name]) Like '%adobe%' " & _
                    "Or Lower([Value].[value]) Like '%adobe%' " & _
                    "Or Lower([Key].[name]) Like '%{6d4b8b88-427d-4363-a533-020c624e4ad1}%' " & _
                    "Or Lower([Key].[name]) Like '%{ac76ba86-7ad7-1033-7b44-aa1000000001}%' " & _
                    "Or Lower([Key].[name]) Like '%{dc6efb56-9cfa-464d-8880-44885d7dc193}%' " & _
                    "Or Lower([Key].[name]) Like '%68ab67ca7da73301b744aa0100000010%' " & _
                "Order By [Key].[name]; "
            Dim ds As DataSet = OpenDataSet(SQL)
            mWriter = My.Computer.FileSystem.OpenTextFileWriter(FileName, False, System.Text.Encoding.Unicode)
            mWriter.WriteLine("Windows Registry Editor Version 5.00" & vbCrLf)
            For iRow As Integer = 0 To ds.Tables(0).DefaultView.Count - 1
                Dim drv As DataRowView = ds.Tables(0).DefaultView(iRow)
                Dim KeyName As String = drv(0)
                Dim KeyTokenCount = TokenCount(KeyName, "\")
                If KeyTokenCount > 2 Then
                    Dim tmpKeyName As String = String.Format("{0}\", ParseStr(KeyName, 1, "\"))
                    For i As Integer = 2 To KeyTokenCount - 1
                        tmpKeyName &= String.Format("{0}\", ParseStr(KeyName, i, "\"))
                        mWriter.WriteLine(String.Format("[{0}]{1}", tmpKeyName, vbCrLf))
                    Next i
                End If
                mWriter.WriteLine(drv(1) & vbCrLf)
            Next iRow
        Finally
            If Not IsNothing(mWriter) Then mWriter.Close() : mWriter = Nothing
        End Try
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
                            Dim valueName As String = ParseString(buffer, 1, "=", """")
                            Dim valueValue As String = DecodeValue(ParseString(buffer, 2, "=", """"), raw)    'Note: DecodeValue will update raw...
                            AddValue(keyID, valueName, valueValue)
                    End Select
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