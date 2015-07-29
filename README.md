# nodejs-runmacro

Run Word macros in a node.js server

Usage: POST request on /macro with two form parameters
- macro: the name of the macro to run
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

Running the service:
-command: node macroService.js <path to temp folder>

You can use the example HTML form and the test Word document, macro is insertDate.  



