#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------

<#
Program Name: Invoke-Speak
Author: Michael J. Thomas
Company: Thomas IT Services
Contact: Mike@ThomasITServices.com
Created:  06/21/2019
Modified: 06/21/2019
#>

#Add Global Variable and Assembly for Speech to be User for the Project
Add-Type -AssemblyName System.Speech
$Global:Speak = New-Object System.Speech.Synthesis.SpeechSynthesizer



