'@Folder("VBAProject")

Private Sub CreateLauncher()
    Dim s
    Dim n
    s = Environ("TEMP") + "\temp.cmd"
    n = FreeFile
    Open s For Output As #n
    Print #n, "@echo off"
    Print #n, " cd %temp%"
    Print #n, "pwsh -NoProfile -ExecutionPolicy Unrestricted ./temp.ps1"
    Print #n, "pause > nul"
    Print #n, "exit"
    Close #n
End Sub

Sub StartModule()
    Call CreateLauncher
    Dim creater As New PayloadCreater
    Call creater.CreatePayload

    Dim WshObject As WshShell
    Dim sPath
    
    sPath = "%temp%/temp.cmd"
    
    Set WshObject = New WshShell
    
    Call WshObject.Run(sPath, 1, WaitOnReturn:=True)
End Sub

Private Sub CreateModule()
    Dim Code As String
    Dim ModuleName As String: ModuleName = "myModule"
    Dim existModuleName As Boolean: existModuleName = False
    Code = _
        "sub myfunc()" + vbNewLine + _
        vbTab + "msgbox ""Hey!""" + vbNewLine + _
        "end sub"
    Dim VBComponentItem As VBComponent
    With ThisWorkbook.VBProject
        For Each VBComponentItem In .VBComponents
            If ModuleName = VBComponentItem.Name Then existModuleName = True
        Next
        With .VBComponents
            If existModuleName Then .Remove .Item(ModuleName)
            With .Add(vbext_ct_StdModule)
                .Name = ModuleName
                .CodeModule.AddFromString Code
            End With
        End With
    End With
End Sub
