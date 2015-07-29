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
On Error Resume Next

' start only if 2 or more arguments are provided 
if WScript.Arguments.Count <> 2 then
	wscript.quit
end if

wscript.echo "Starting Word ..."
Set word = CreateObject("Word.Application")
word.Visible = FALSE
checkError "creating Word instance"

wscript.echo "Loading document ..."
Set doc = word.Documents.Open( Wscript.Arguments(0))
checkError "loading document"

if doc <> nul then

	wscript.echo "Running macro ..."
	word.Run Wscript.Arguments(1)
	
	checkError "running macro"

	if doc.Saved = false then
		wscript.echo "Saving document ..."
		doc.Save
	end if

	checkError "saving document"
end if

wscript.echo " "

' avoid MS Word prompting to save normal.dot
word.NormalTemplate.Saved = true
word.Quit 0
checkError "closing Word instance"

wscript.echo "Macro execution finished"