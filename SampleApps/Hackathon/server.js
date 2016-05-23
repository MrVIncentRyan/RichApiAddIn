var port = process.env.port || 8080;
var fs = require('fs');
var options = {key: fs.readFileSync('./key.pem'), cert: fs.readFileSync('./cert.crt')};

var app = require('express')();
var server = require('https').Server(options, app);
var io = require('socket.io')(server);

server.listen(port, function() {
	console.log("Listening:");
});

io.on('connection', function(socket) {
	var notebookId;
	socket.on('notebookId', function(data) {
		notebookId = data;
		socket.join(notebookId);
		console.log(notebookId);
	});
	
	socket.on('presentContent', function(data) {
		if (notebookId) {
			socket.broadcast.to(notebookId).emit('presentContent', data);
		}
	});
	
	socket.on('quiz', function(data) {
		if (notebookId) {
			socket.broadcast.to(notebookId).emit('quiz', data);
		}
	});
	
	socket.on('quizResponse', function(data) {
		if (notebookId) {
			socket.broadcast.to(notebookId).emit('quizResponse', data);
		}
	});
	
	socket.on('quizDone', function() {
		if (notebookId) {
			socket.broadcast.to(notebookId).emit('quizDone');
		}
	});
});