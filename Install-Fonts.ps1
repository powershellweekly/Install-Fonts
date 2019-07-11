<#
.Synopsis
   Install-Fonts
   Author: Michael J. Thomas
   Created: 07/10/2019
   Updated: 07/10/2019
.DESCRIPTION
   Install Fonts on a computer.
.EXAMPLE
   Install-Fonts -Files "C:\Fonts"
.EXAMPLE
   Install-Fonts -File "C:\Fonts\LeanStatus.ttf"
#>
function Install-Fonts
{
    [CmdletBinding()]
    Param
    (
        [string[]]$Files,
        [string]$File
    )

    $objShell = New-Object -ComObject Shell.Application
    $Fonts = $objShell.NameSpace(20)
    If (!($Files -eq $null)){  Get-ChildItem "$Files\*.ttf" | ForEach-Object {$Fonts.CopyHere($_.FullName)} }
    ElseIf (!($File -eq $null)){ $Fonts.CopyHere($File) }
  
}
