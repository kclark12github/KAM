Module modMain
    Dim objMain As clsMain = New clsMain
    Sub Main()
        Dim test As String = """Assembly""=""Microsoft.Office.Interop.Outlook, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71E9BCE111E9429C"""
        Console.WriteLine(objMain.ParseString(test, 1, "=", """"))
        Console.WriteLine(objMain.ParseString(test, 2, "=", """"))
        Console.WriteLine(objMain.ParseString(test, 3, "=", """"))
        objMain.ProcessFile("L:\HKCU.reg")
    End Sub
End Module
