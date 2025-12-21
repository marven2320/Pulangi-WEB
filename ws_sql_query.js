#!/usr/bin/env node
var WebSocketServer = require('websocket').server;
var http = require('https');
const fs = require('fs');
const path = require('path')
const express = require('express');
const serve = require('express-static');
var clients = [];
var clientsIP = [];

const options = {
    key: fs.readFileSync('key.pem'),
    cert: fs.readFileSync('cert.pem')
};

var app = express();

app.use(serve(__dirname + '/'));

var server = http.createServer(options, app);
server.listen(3000, function () {
    console.log((new Date()) + ' Server is listening on port 3000');
});


//MYSQL
const mysql = require('mysql2');
const pool = mysql.createPool({
    host: 'localhost',
    user: 'root',
    password: 'KaRMSys2025!',
    database: 'pulangi_data',
    connectionLimit: 10
});

//------start-------- FOR WEBSOCKET

wsServer = new WebSocketServer({
    httpServer: server,
    // You should not use autoAcceptConnections for production
    // applications, as it defeats all standard cross-origin protection
    // facilities built into the protocol and the browser.  You should
    // *always* verify the connection's origin and decide whether or not
    // to accept it.
    autoAcceptConnections: false
});

function originIsAllowed(origin) {
    // put logic here to detect whether the specified origin is allowed.
    return true;
}

wsServer.on('request', function (request) {
    if (!originIsAllowed(request.origin)) {
        // Make sure we only accept requests from an allowed origin
        request.reject();
        console.log((new Date()) + ' Connection from origin ' + request.origin + ' rejected.');
        return;
    }
    var connection = request.accept('echo-protocol', request.origin);
    //console.log(connection);
    clients.push(connection);
    clientsIP.push(connection.remoteAddress);
    console.log(clientsIP);
    console.log((new Date()) + ' Connection accepted.');

    connection.on('message', function (message) {
        if (message.type === 'utf8') {
            console.log('Received Message: ' + message.utf8Data);
            connection.sendUTF(message.utf8Data);

            try {
                var data = JSON.parse(message.utf8Data);
                // Ensure data has the expected structure before building query
                if (data && data[0] && data[0].startdate && data[0].enddate) {
                    const sql = 'SELECT * from pulangi WHERE date BETWEEN\'' + data[0].startdate + '\'' + ' AND \'' + data[0].enddate + '\'' + ' ORDER by time ASC';
                    console.log(sql);

                    pool.query({ sql }, (err, result, fields) => {
                        if (err instanceof Error) {
                            console.log(err);
                            return;
                        }
                        // console.log(JSON.stringify(result)); // results contains rows returned by server
                        connection.send(JSON.stringify(result));
                    });
                }
            } catch (err) {
                console.log(err);
            }
        }
        else if (message.type === 'binary') {
            //console.log('Received Binary Message of ' + message.binaryData.length + ' bytes');
            connection.sendBytes(message.binaryData);
        }
    });

    connection.on('close', function (reasonCode, description) {
        //console.log((new Date()) + ' Peer ' + connection.remoteAddress + ' disconnected.');
        const index = clientsIP.indexOf(connection.remoteAddress);
        if (index > -1) { // only splice array when item is found
            clients.splice(index, 1); // 2nd parameter means remove one item only
            clientsIP.splice(index, 1); // 2nd parameter means remove one item only
        }
    });
});
//------end-------- FOR WEBSOCKET
