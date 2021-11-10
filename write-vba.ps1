using namespace System.Management.Automation
using namespace System.Collections.Generic
using namespace System.Runtime.InteropServices
using namespace Microsoft.Office.Interop.Excel

# VBAのEnumを使うためにアセンブリをロード
# Powershell 7.1ではアセンブリ名で見つけることが出来なかったのでPathで指定
foreach ($item in @(
        "C:\Windows\assembly\GAC_MSIL\Microsoft.VisualBasic\*\Microsoft.VisualBasic.dll"
        "C:\Windows\assembly\GAC_MSIL\office\*\OFFICE.DLL"
        "C:\Windows\assembly\GAC_MSIL\Microsoft.Vbe.Interop\*\Microsoft.Vbe.Interop.dll"
        "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\*\Microsoft.Office.Interop.Excel.dll"
    )) {
    Add-Type -Path $item
}

# COMObjectに拡張メソッドを追加
Update-TypeData -TypeName System.__ComObject -MemberType ScriptMethod -MemberName Tee -Value {
    param(
        [Stack[WeakReference]]
        $stack
    )
    $stack.Push([WeakReference]::new($this))

    Write-Output $this
}

# Comを解放するために使う弱参照のスタック
$refs = [Stack[WeakReference]]::new()

function CreateModule {
    param(
        [string]$fileName
    )
    $filePath = Join-Path $pwd $fileName  
    $excel = New-Object -ComObject Excel.Application  
    [string]$ModuleName = "PayloadCreater"
    [bool]$existModuleName = $false
    [string[]]$codeList = Get-Content "src/Classes/PayloadCreater.vb"
    [string]$Code = ""
    foreach ($item in $codelist) {
        $Code += $item + "`n"
    }

    [ExcelSecurityRegistry]$excelRegistry = [ExcelSecurityRegistry]::new()
    $excelRegistry.SetWritable()

    $workbook = $excel.Workbooks.Open($filePath)
    foreach ($item in $workbook.VBProject.VBComponents) {
        if ($item.Name -eq $ModuleName) {
            $existModuleName = $true
        }
    }
    if ($existModuleName) {
        $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item($ModuleName))
    }
    $VBComponent = $workbook.VBProject.VBComponents.Add([Microsoft.Vbe.Interop.vbext_ComponentType]::vbext_ct_ClassModule)
    $VBComponent.Name = $ModuleName
    $VBComponent.CodeModule.AddFromString($Code)
    $workbook.Save()
    $excel.Quit()
    
    $excelRegistry.SetToBefore()
}

class ExcelSecurityRegistry {
    [int]$defaultAccessVBOM
    [int]$defaultVBAWarnings
    [string]$excelRegistryPath = "HKCU:\Software\Microsoft\Office\15.0\excel\Security"
    [void]SetWritable() {
        New-ItemProperty -Path $this.excelRegistryPath -Name `
            AccessVBOM -Value 1 -Force | Out-Null
        New-ItemProperty -Path $this.excelRegistryPath -Name `
            VBAWarnings -Value 1 -Force | Out-Null
    }
    [void]SetToBefore() {
        New-ItemProperty -Path $this.excelRegistryPath -Name `
            AccessVBOM -Value $this.defaultAccessVBOM -Force | Out-Null
        New-ItemProperty -Path $this.excelRegistryPath -Name `
            VBAWarnings -Value $this.defaultVBAWarnings -Force | Out-Null
    }
    ExcelSecurityRegistry() {
        $this.defaultAccessVBOM = (Get-ItemProperty -Path $this.excelRegistryPath).AccessVBOM
        $this.defaultVBAWarnings = (Get-ItemProperty -Path $this.excelRegistryPath).VBAWarnings
    }
}


CreateModule -fileName build/ExecutePwsh.xlsm

While ($refs.Count) {
    # スタックから弱参照を取得
    $comRef = $refs.Pop()

    # 解放するCOMを参照してる変数を全て取得
    $comVar = (Get-Variable).where{ [object]::ReferenceEquals($comRef.Target, $_.Value) }

    # Applicationオブジェクトであるかの判定
    $isApp = $comRef.Target -is [Microsoft.Office.Interop.Excel.Application]

    # アプリケーションの終了前にガベージ コレクトを強制
    if ($isApp) {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
        $comRef.Target.Quit()
    }

    # COMObjectの解放
    # ※正しく動作していればApplicationオブジェクトが解放された瞬間にExcelが終了する
    while ([Marshal]::ReleaseComObject($comRef.Target)) { }
    $comRef.Target = $null

    # 変数を削除
    $comVar | Remove-Variable
    Remove-Variable comRef

    # Application オブジェクトのガベージ コレクトを強制
    if ($isApp) {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }
}
