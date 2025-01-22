#!/usr/bin/env node
var WebSocketServer = require('websocket').server;
var http = require('http');

var server = http.createServer(function(request, response) {
    console.log((new Date()) + ' Received request for ' + request.url);
    response.writeHead(404);
    response.end();
});
server.listen(3000, function() {
    console.log((new Date()) + ' Server is listening on port 3000');
});

const mysql = require('mysql2');
const pool = mysql.createPool({
    host: 'localhost',
    user: 'root',
    //password: 'KaRMSys2025!',
    database: 'pulangi_data',
    connectionLimit: 10
  });
var table = [
    ['"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"']
];

var d = new Date();
var wsdata = [
    {
        id: "pulangi",
        unitnum: 1,
        mw: 0,
        mvar: 0,
        freq: 0,
        vab: 0,
        vbc: 0,
        vca: 0,
        ia: 0,
        ib: 0,
        ic: 0,
        pfa: 0,
        pfb: 0,
        pfc: 0,
        time: d.toTimeString().split(' ')[0],
        date: d.toISOString().split('T')[0]
    },
    {
        id: "pulangi",
        unitnum: 2,
        mw: 0,
        mvar: 0,
        freq: 0,
        vab: 0,
        vbc: 0,
        vca: 0,
        ia: 0,
        ib: 0,
        ic: 0,
        pfa: 0,
        pfb: 0,
        pfc: 0,
        time: d.toTimeString().split(' ')[0],
        date: d.toISOString().split('T')[0]
    },
    {
        id: "pulangi",
        unitnum: 3,
        mw: 0,
        mvar: 0,
        freq: 0,
        vab: 0,
        vbc: 0,
        vca: 0,
        ia: 0,
        ib: 0,
        ic: 0,
        pfa: 0,
        pfb: 0,
        pfc: 0,
        time: d.toTimeString().split(' ')[0],
        date: d.toISOString().split('T')[0]
    }
];

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

wsServer.on('request', function(request) {
    if (!originIsAllowed(request.origin)) {
      // Make sure we only accept requests from an allowed origin
      request.reject();
      console.log((new Date()) + ' Connection from origin ' + request.origin + ' rejected.');
      return;
    }
    
    var connection = request.accept('echo-protocol', request.origin);
    console.log((new Date()) + ' Connection accepted.');
    
    setInterval(function(){
        
        wsdata = [
            {
                id: "pulangi",
                unitnum: 1,
                mw: Math.random()*100,
                mvar: Math.random()*100,
                freq: Math.random()*10,
                vab: Math.random()*100,
                vbc: Math.random()*100,
                vca: Math.random()*100,
                ia: Math.random()*100,
                ib: Math.random()*100,
                ic: Math.random(),
                pfa: Math.random(),
                pfb: Math.random(),
                pfc: Math.random(),
                time: d.toTimeString().split(' ')[0],
                date: d.toISOString().split('T')[0]
            },
            {
                id: "pulangi",
                unitnum: 2,
                mw: Math.random()*100,
                mvar: Math.random()*100,
                freq: Math.random()*100,
                vab: Math.random()*100,
                vbc: Math.random()*100,
                vca: Math.random()*100,
                ia: Math.random()*100,
                ib: Math.random()*100,
                ic: Math.random()*100,
                pfa: Math.random(),
                pfb: Math.random(),
                pfc: Math.random(),
                time: d.toTimeString().split(' ')[0],
                date: d.toISOString().split('T')[0]
            },
            {
                id: "pulangi",
                unitnum: 3,
                mw: Math.random()*100,
                mvar: Math.random()*100,
                freq: Math.random()*100,
                vab: Math.random()*100,
                vbc: Math.random()*100,
                vca: Math.random()*100,
                ia: Math.random()*100,
                ib: Math.random()*100,
                ic: Math.random()*100,
                pfa: Math.random(),
                pfb: Math.random(),
                pfc: Math.random(),
                time: d.toTimeString().split(' ')[0],
                date: d.toISOString().split('T')[0]
            }
                ];


        connection.send(JSON.stringify(wsdata));
        
    },1000);
    

    connection.on('message', function(message) {
           var data = JSON.parse(message.utf8Data); 
        
            
            console.log('Received Message: ' + data[0].startdate);
            console.log('Received Message: ' + data[0].enddate);


            const  sql = 'SELECT * from pulangi WHERE date BETWEEN \''+ data[0].startdate+ '\'' + ' AND \''+data[0].enddate+ '\''+' ORDER by date ASC';
            //const  sql = 'SELECT * from pulangi WHERE mw= 100';
            //console.log(sql);

            try {
                  pool.query({sql},(err, result, fields) => {
                  if (err instanceof Error) {
                    console.log(err);
                    return;
                  }
                  //console.log(result); // results contains rows returned by server
                  connection.send(JSON.stringify(result));
                  });
            }catch (err) {
                  console.log(err);
            }

           
        
    });


    connection.on('close', function(reasonCode, description) {
        console.log((new Date()) + ' Peer ' + connection.remoteAddress + ' disconnected.');
    });
});
