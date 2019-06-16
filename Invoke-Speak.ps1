<#
.Synopsis
   Invoke-Speak
   Created by Michael J. Thomas
   Date:    06/15/2019
   Updated: 06/15/2019

.DESCRIPTION
    Invoke-Speak allows you to speak to a remote computer with Microsoft Speach
    Invoke-Speak -Words $Words -ComputerName $Computers -Speed (-10 to 10) Slow to Fast
    Default Speed is 2.
    
.EXAMPLE
   Invoke-Speak "Hello World!"
   Invoke-Speak "Hello World! -ComputerName "Computer01", "Computer02" -Speed 2
#>
function Invoke-Speak{
[cmdLetBinding()]
Param(
[string]$Words,
[string[]]$ComputerName =$env:COMPUTERNAME,
[int]$Speed = 2
)

Foreach ($Computer in $ComputerName){
Invoke-Command -ComputerName $Computer -ScriptBlock{
Add-Type -AssemblyName System.Speech 
$Speak = New-Object System.Speech.Synthesis.SpeechSynthesizer 
$Speak.Rate = $Using:Speed
$Speak.Speak($Using:Words)
$Speak.Dispose()
}
}
}

Invoke-Speak "Hello World!"