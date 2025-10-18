REM This sample code shows how to start Set-OutlookSignatures in a hidden, non-blocking window
REM
REM A PowerShell windows will pop up for a second.
REM For a completely hidden method, read FAQ 'Start Set-OutlookSignatures in hidden/invisible mode' in 'README' file
REM 
REM You have to adapt it to fit your environment.
REM The sample code is written in a generic way, which allows for easy adaption.
REM 
REM Would you like support? ExplicIT Consulting (https://explicitconsulting.at) offers professional support for this and other open source code.


start powershell.exe -WindowStyle hidden -Command "Start-Transcript; & '\\server.example.com\share\folder\Set-OutlookSignatures\set-outlooksignatures.ps1' -SignatureTemplatePath '\\server.example.com\share\folder\templates\Signatures DOCX' -SignatureIniFile '\\server.example.com\share\folder\Signatures DOCX\_.ini' -OOFTemplatePath '\\server.example.com\share\folder\templates\Out-of-Office DOCX' -OOFIniFile '\\server.example.com\share\folder\templates\Out-of-Office DOCX\_.ini'; Stop-Transcript"
