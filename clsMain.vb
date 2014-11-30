Public Class clsMain
    Public Sub Main()
        Dim fileUnit As Integer = FreeFile()
        Dim fileName As String = ""
        Dim key As String = ""
        Dim keyID As Integer = 0
        Dim parentID As Integer = 0
        Dim root As String = ""
        Dim raw As String = ""
        FileOpen(fileUnit, fileName, OpenMode.Input, OpenAccess.Read, OpenShare.Default)
        While Not EOF(fileUnit)
            Dim buffer As String = LineInput(fileUnit).Trim
            Try
                If buffer = "" OrElse buffer = "Windows Registry Editor Version 5.00" Then Throw New KeyNotFoundException()
                Select Case buffer.Substring(0, 1)
                    Case "["
                        If root = "" Then
                            root = buffer.Substring(1, buffer.Length - 2)

                        Else
                        End If
                    Case """"
                        raw &= buffer

                    Case "@"
                        raw &= buffer
                    Case Else
                End Select

            Catch ex As KeyNotFoundException
            End Try
        End While
    End Sub
End Class
