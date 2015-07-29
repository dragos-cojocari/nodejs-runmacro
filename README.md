# nodejs-runmacro

Run Word macros in a node.js server

API: 
- /macro 
	- method: POST
	- form parameters:
		- macro: the name of the macro to run. Only alphanumerical characters are permitted
		- document: the Word document

Response:
- 400 if the parameters are not provided-
- the Word document after the macro was executed on it
   
Dependencies:
- node.js modules: formidable  
- Windows cscript engine 
- Microsoft Word must be installed on the same machine

Installation: 
- copy the js and vbs file in a location visible to node.js
- install formidable if required: npm install formidable

Run the service:
-command: node macroService.js <path to temp folder>

Access the service programatically:
- client/runRemoteMacro.bat

Test artifacts:
- example/testMacroService.html 
- example/testMacro.doc with the insertDate date
