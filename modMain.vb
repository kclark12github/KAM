Module modMain
    Dim objMain As clsMain = New clsMain
    Sub Main()
        'objMain.ProcessFile("L:\HKCU.reg")
        'objMain.ProcessFile("L:\HKCR.reg")
        'objMain.ProcessFile("L:\HKLM.reg")
        objMain.OutputFile("L:\Adobe.reg")
    End Sub
End Module
