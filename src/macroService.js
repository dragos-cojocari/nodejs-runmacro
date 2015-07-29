//----------------------------------------------------------------------------------
// Author: Dragos Cojocari
// Last modified date: 2015.07.29
// DISCLAIMER: the example is provided as-is without warranty of any kind, express or implied.
// LICENSE: you can freely use, modify and distribute this file as long as you keep this note
//----------------------------------------------------------------------------------

// Load the http module to create an http server.
var http = require('http');
var url = require('url');
var qs = require('querystring');
var formidable = require('formidable');

var util = require('util');
var exec = require('child_process').exec;
var child;

var crypto = require('crypto');
var fs = require('fs'); 

// read the temp folder
var tmpFolder = process.argv[2];
if ( tmpFolder == undefined)
{
	console.log("No temp folder specified.\nSyntax is: node chartServer.js <full path to temp folder>");
	return;
}
console.log("Using temp folder: " + tmpFolder);

// the server function
var server = http.createServer(serverFunc);

//Listen on port 8000, IP defaults to 127.0.0.1
server.listen(8000);

// Put a friendly message on the terminal
console.log("Server running at http://127.0.0.1:8000/");

	
function serverFunc(request, response) {
	// only allow POST 
	if (request.method !== 'POST') { 
		response.writeHead(405); 
		response.end('Request method not supported. The chart server supports only POST', 'utf8'); 
		return;
	}
  
	var url_parts = url.parse(request.url, true);
	if( url_parts.pathname !== "/macro")
	{
		response.writeHead(400);
		response.end( 'Invalid request. \nSyntax is: http://server:port/macro"');
		return;
	}
  
	var form = new formidable.IncomingForm({ uploadDir: tmpFolder});
	
	// form.parse analyzes the incoming stream data, picking apart the different fields and files for you.
	form.parse(request, function(err, fields, files) {
		if (err) {

		  response.writeHead(500, {'content-type': 'text/plain'});
		  response.write( "Error: " );
		  response.end( err.message);
		  
		  console.error(err.message);
		  return;
		}
		
		console.log("Request received " + new Date());
		//console.log(util.inspect({fields: fields, files: files}));
		
		if ( files.document == undefined || files.document.size == 0 || fields.macro == undefined || fields.macro == "")
		{
			response.writeHead(500, {'content-type': 'text/plain'});
			response.end( 'Invalid request. A file and a macro must be specified.');
			return;
		}
		
		// for security reasons only allow alphanumerical characters
		var pattern = /^[a-z0-9]+$/i;
		if ( !pattern.test( fields.macro))
		{
			response.writeHead(500, {'content-type': 'text/plain'});
			response.end( 'Invalid request. Macro name can only contain alphanumerical characters.');
			return;
		}
		
		//console.log( "Received: " + files.file.path);
		runMacro( files.document.path, fields.macro, response);
	});
}

function runMacro( path, macro, response)
{
	var cmd = 'cmd /c cscript runMacro.vbs ' + path + ' ' +  macro;
	console.log('running cmd: ' + cmd);
		
	// run it
	child = exec( cmd, // command line argument directly in string
	function (error, stdout, stderr) {      // one easy function to capture data/errors
		console.log(stdout.trim());
		console.log(stderr.trim());
		if (error !== null) {
		  console.log('exec error: ' + error);
		}
		
		// read the stream
		var stream = fs.createReadStream(path);
		
		// error handler
		stream.on('error', function(error) {
			response.writeHead(500);
			response.end( "Could not read the result file.");
			return;
		});
		
		// error handler
		stream.on('end', function() {
			console.log('Document served. Deleting temp file: ' + path);
			fs.unlinkSync(path);
		});
		
		// stream it back
		response.writeHead(200, {"Content-Type": "application/msword", "Content-disposition" : "attachment; filename=document.doc"});
		stream.pipe( response);
	});
}


