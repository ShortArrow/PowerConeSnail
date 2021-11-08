Sub CreateModule()
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