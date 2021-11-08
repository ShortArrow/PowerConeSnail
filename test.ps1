using namespace System.Management.Automation
using namespace System.Collections.Generic
using namespace System.Runtime.InteropServices
using namespace Microsoft.Office.Interop.Excel

param(
    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "Path to one locations.")]
    [ValidateScript( { Test-Path $_  -Include "*.xlsx" })]
    [string]
    $ExcelPath
)

# VBAのEnumを使うためにアセンブリをロード
# Powershell 7.1ではアセンブリ名で見つけることが出来なかったのでPathで指定
@(
    "C:\Windows\assembly\GAC_MSIL\Microsoft.VisualBasic\*\Microsoft.VisualBasic.dll"
    "C:\Windows\assembly\GAC_MSIL\office\*\OFFICE.DLL"
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Vbe.Interop\*\Microsoft.Vbe.Interop.dll"
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\*\Microsoft.Office.Interop.Excel.dll"
).ForEach{
    Add-Type -path $_
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

# Excel操作本体
$app = (New-Object -ComObject Excel.Application).Tee($refs)
$book = $app.Workbooks.Tee($refs).Open($ExcelPath).Tee($refs)
$book.Worksheets.Tee($refs).Item('Sheet1').Tee($refs).Range('A1:B2').Tee($refs)._NewEnum.Tee($refs).ForEach{
    $_.Text
    [void][Marshal]::ReleaseComObject($_)
}

# (Get-CimInstance -Class Win32_Process -Filter "ProcessId = $((ps excel).id)" -Property "CommandLine").CommandLine

# ファイルを閉じる
$book.Close()

# COMの開放
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