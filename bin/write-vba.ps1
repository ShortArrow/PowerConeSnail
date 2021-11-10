using namespace System.Management.Automation
using namespace System.Collections.Generic
using namespace System.Runtime.InteropServices
using namespace Microsoft.Office.Interop.Excel
using namespace Microsoft.Vbe.Interop

param (
    [string]$SrcFile = "src/main.ps1",
    [string]$DistFile = "build/ExecutePwsh.xlsm"
)
# [Reflection.Assembly]::LoadWithPartialName("Microsoft.Vbe.Interop") | Out-Null
# Add-Type -AssemblyName Microsoft.Vbe.Interop

@(
    "C:\Windows\assembly\GAC_MSIL\Microsoft.VisualBasic\*\Microsoft.VisualBasic.dll"
    "C:\Windows\assembly\GAC_MSIL\office\*\OFFICE.DLL"
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Vbe.Interop\*\Microsoft.Vbe.Interop.dll"
    "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\*\Microsoft.Office.Interop.Excel.dll"
).ForEach{
    Add-Type -path $_
}


# enum CodeTypes {
#     ClassModule = [Microsoft.Vbe.Interop.vbext_ComponentType]::vbext_ct_ClassModule
#     StdModule = [Microsoft.Vbe.Interop.vbext_ComponentType]::vbext_ct_StdModule
# }

function Write-Modules {
    param(
        [string]$Dist = "build/ExecutePwsh.xlsm",
        [string]$ClassDir = "bin/Classes",
        [string]$ModuleDir = "bin/Modules",
        [string]$Src = "src/main.ps1"
    )
    $filePath = Join-Path $pwd $Dist  
    $excel = New-Object -ComObject Excel.Application

    [ExcelSecurityRegistry]$excelRegistry = [ExcelSecurityRegistry]::new()
    $excelRegistry.SetWritable()
    
    $workbook = $excel.Workbooks.Open($filePath)
    
    . bin\build-module.ps1
    Build-Vba -sourceFilename "src/main.ps1" -temporaryFilename "\temp.ps1"
    Write-Module -workbook $workbook -srcPath "bin/Classes/PayloadCreater.vb" -codetype vbext_ct_ClassModule
    Write-Module -workbook $workbook -srcPath "bin/Modules/Module1.vb" -codetype vbext_ct_StdModule
    
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

Write-Modules -Dist $DistFile -Src $SrcFile
