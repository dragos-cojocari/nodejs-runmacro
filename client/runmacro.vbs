'*******************************************************************************
' *
' * Licensed Materials - Property of IBM
' *
' * (c) Copyright IBM Corp. 2008, 2013 All Rights Reserved
' *
' * US Government Users Restricted Rights - Use, duplication or
' * disclosure restricted by GSA ADP Schedule Contract with
' * IBM Corp.
' *
'*******************************************************************************

Sub checkError( message)
	If Err.Number <> 0 Then
		wscript.echo "Error " & message & ". Error: " & Err.Number & " - " &  Err.Description
		Err.Clear
	end if
End Sub

' ignore errors
' On Error Resume Next

' start only if 2 or more arguments are provided 
if WScript.Arguments.Count <> 2 then
	wscript.echo "Syntax: runmacro.vbs document macro"
	wscript.quit
end if

Dim cmd 
cmd = "d:\tools\curl\runRemoteMacro.bat " & Wscript.Arguments(0) & " " & Wscript.Arguments(1)

wscript.echo "Running command: " & cmd

' run the CURL command and wait for it
Set oShell = CreateObject ("WScript.Shell")
oShell.run cmd, 1, true

wscript.echo "All done"