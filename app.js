const express = require('express');
const path = require('path');
const http = require('http');
const app = express();

app.use(express.static(path.join(__dirname, 'public')));

const port = process.env.PORT || '3002';
app.set('port', port);

app.get('/', function(req, res){
	console.log('m here');
	res.sendFile(path.join(__dirname+'/public/taskpane.html'));
})

const server = http.createServer(app);
server.listen(port, () => {
	console.log("Listening on port " + port);
});