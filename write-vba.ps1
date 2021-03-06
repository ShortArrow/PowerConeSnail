using namespace System.Management.Automation
using namespace System.Collections.Generic
using namespace System.Runtime.InteropServices
using namespace Microsoft.Office.Interop.Excel

function Write-Modules {
    param(
        [string]$fileName = "build/ExecutePwsh.xlsm",
        [string]$ClassDir = "src/Classes",
        [string]$ModuleDir = "src/Modules"
    )
    $filePath = Join-Path $pwd $fileName  
    $excel = New-Object -ComObject Excel.Application

    [ExcelSecurityRegistry]$excelRegistry = [ExcelSecurityRegistry]::new()
    $excelRegistry.SetWritable()
    
    $workbook = $excel.Workbooks.Open($filePath)
    
    Write-Module -workbook $workbook -srcPath "src/Classes/PayloadCreater.vb" -codetype vbext_ct_ClassModule
    Write-Module -workbook $workbook -srcPath "src/Modules/Module1.vb" -codetype vbext_ct_StdModule
    
    $workbook.Save()
    $excel.Quit()

    # Clear variavles referencing __ComObject(as same $val=$null)
    # In VBA as same as set `Nothing`
    Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable

    # force GC
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()

    # Cleanup auto-variable
    1 | ForEach-Object { $_ } > $null
    [gc]::Collect()
    
    $excelRegistry.SetToBefore()
}

function Write-Module {
    param (
        [Microsoft.Vbe.Interop.vbext_ComponentType]
        $codetype,
        [string]
        $srcPath,
        [System.__ComObject]
        $workbook
    )
    
    [string[]]$codeList = Get-Content $srcPath
    [string]$Code = ""
    foreach ($item in $codelist) {
        $Code += $item + "`n"
    }
    
    [string]$ModuleName = Split-Path $srcPath -LeafBase
    [bool]$existModuleName = $false
    foreach ($item in $workbook.VBProject.VBComponents) {
        if ($item.Name -eq $ModuleName) {
            $existModuleName = $true
        }
    }
    if ($existModuleName) {
        $workbook.VBProject.VBComponents.Remove($workbook.VBProject.VBComponents.Item($ModuleName))
    }
    $VBComponent = $workbook.VBProject.VBComponents.Add($codetype)
    $VBComponent.Name = $ModuleName
    $VBComponent.CodeModule.AddFromString($Code)
}

class ExcelSecurityRegistry {
    [int]$defaultAccessVBOM
    [int]$defaultVBAWarnings
    [string]$excelRegistryPath
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
        $this.excelRegistryPath = Get-ExcelRegistryRoot
        $this.defaultAccessVBOM = (Get-ItemProperty -Path $this.excelRegistryPath).AccessVBOM
        $this.defaultVBAWarnings = (Get-ItemProperty -Path $this.excelRegistryPath).VBAWarnings
    }
}

function Get-ExcelRegistryRoot {
    [OutputType([string])]
    param ()
    $officeRoot = "HKCU:\Software\Microsoft\Office\"
    $securityPath = "excel\Security"
    $list = Get-ChildItem $officeRoot
    [double]$highest = 0
    foreach ($item in $list) {
        [double]$returnedInt = 0
        if ([double]::TryParse((Split-Path $item -Leaf), [ref]$returnedInt)) {
            if (($highest -lt $returnedInt) -and (Test-Path (Join-Path $item.PSPath $securityPath))) {
                $highest = $returnedInt
            }
        }
    }

    return $officeRoot + ("{0:0.0}\" -f $highest) + $securityPath
}

Write-Modules -fileName build/ExecutePwsh.xlsm
