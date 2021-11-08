$codeList = @(
    "Sub CreatePayload()"
    "Dim s"
    "Dim n"
    "s = Environ(`"TEMP`") + `"\temp.ps1`""
    "n = FreeFile"
    "Open s For Output As #n"
    "Print #n, `"Write-Host (Get-Location).Path `""
    "Print #n, `"Write-Host this`""
    "Print #n, `"`""
    "Print #n, `"`""
    "Print #n, `"`""
    "Close #n"
    "End Sub"
)
echo $codeList[0]