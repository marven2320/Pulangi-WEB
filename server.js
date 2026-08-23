#!/usr/bin/env node
var WebSocketServer = require('websocket').server;
var https = require('https');
const fs = require('fs');
const path = require('path')
const express = require('express');
const serve   = require('express-static');


const options = {
  key: fs.readFileSync('key.pem'),
  cert: fs.readFileSync('cert.pem')
};

var app = express();

const mysql = require('mysql2');
const pool = mysql.createPool({
    host: 'localhost',
    user: 'root',
    password: 'KaRMSys2025!',
    database: 'pulangi_data',
    connectionLimit: 10
  });

// Shared with fetch-generation-data.js (the standalone CLI script) - same
// validation + SQL, single source of truth for what viewdata.html's charts
// and the CLI tool both see.
const { getGenerationData, isValidParams } = require('./generation-data-query');

// GET /api/generation-data?date=YYYY-MM-DD&start=HH:MM:SS&end=HH:MM:SS
// Returns per-timestamp power (mw1/mw2/mw3) and frequency (freq1/freq2/freq3)
// readings for all 3 units, straight from the `pulangi` table - used by
// viewdata.html's Power/Frequency dual-axis charts so it no longer needs a
// manual CSV/Excel upload.
// NOTE: this must be registered before app.use(serve(...)) below - the
// express-static middleware does not call next() on a non-matching path,
// it errors out instead, which would otherwise shadow this route entirely.
app.get('/api/generation-data', (req, res) => {
    const { date, start, end } = req.query;

    if (!isValidParams(date, start, end)) {
        res.status(400).json({ error: 'date must be YYYY-MM-DD and start/end must be HH:MM:SS' });
        return;
    }

    getGenerationData(pool, { date, start, end }, (err, rows) => {
        if (err) {
            console.log((new Date()) + ' /api/generation-data query error: ' + err);
            res.status(500).json({ error: 'database query failed' });
            return;
        }
        res.json({ rows });
    });
});

// GET /api/generation-data-file
// Serves the current year's raw Generation Data workbook as-is for download -
// this is the file dgr_merger.js maintains at
// rawdata/Pulangi IV HEP - Generation Data - RAW_<year>.xlsx. Used by the
// Reports page's "Generation Data" option (Open Report), which just opens
// this file rather than querying/filtering anything itself.
app.get('/api/generation-data-file', (req, res) => {
    const year = new Date().getFullYear();
    const fileName = `Pulangi IV HEP - Generation Data - RAW_${year}.xlsx`;
    const filePath = path.join(__dirname, 'rawdata', fileName);

    if (!fs.existsSync(filePath)) {
        res.status(404).json({ error: `${fileName} not found` });
        return;
    }

    res.download(filePath, fileName, (err) => {
        if (err) {
            console.log((new Date()) + ' /api/generation-data-file download error: ' + err);
        }
    });
});

app.use(serve(__dirname + '/'));

 var server = https.createServer(options, app);
 server.listen(8000, function() {
    console.log((new Date()) + ' Server is listening on port 8000');
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
    
    function getRandomArbitrary(min, max) {
          return Math.random() * (max - min) + min;
    }

    setInterval(function(){
        
        wsdata = [
            {
                id: "pulangi",
                unitnum: 1,
                mw: Math.random()*100,
                mvar: Math.random()*100,
                freq: getRandomArbitrary(59,61),
                vab: getRandomArbitrary(-1,13.8),
                vbc: getRandomArbitrary(-1,13.8),
                vca: getRandomArbitrary(-1,13.8),
                ia: Math.random()*5000,
                ib: Math.random()*5000,
                ic: Math.random(),
                pfa: getRandomArbitrary(0,1),
                pfb: getRandomArbitrary(0,1),
                pfc: getRandomArbitrary(0,1),
                time: d.toTimeString().split(' ')[0],
                date: d.toISOString().split('T')[0]
            },
            {
                id: "pulangi",
                unitnum: 2,
                mw: Math.random()*100,
                mvar: Math.random()*100,
                freq: getRandomArbitrary(59,61),
                vab: getRandomArbitrary(-1,13.8),
                vbc: getRandomArbitrary(-1,13.8),
                vca: getRandomArbitrary(-1,13.8),
                ia: Math.random()*5000,
                ib: Math.random()*5000,
                ic: Math.random()*5000,
                pfa: getRandomArbitrary(0,1),
                pfb: getRandomArbitrary(0,1),
                pfc: getRandomArbitrary(0,1),
                time: d.toTimeString().split(' ')[0],
                date: d.toISOString().split('T')[0]
            },
            {
                id: "pulangi",
                unitnum: 3,
                mw: Math.random()*100,
                mvar: Math.random()*100,
                freq: getRandomArbitrary(59,61),
                vab: getRandomArbitrary(-1,13.8),
                vbc: getRandomArbitrary(-1,13.8),
                vca: getRandomArbitrary(-1,13.8),
                ia: Math.random()*5000,
                ib: Math.random()*5000,
                ic: Math.random()*5000,
                pfa: getRandomArbitrary(0,1),
                pfb: getRandomArbitrary(0,1),
                pfc: getRandomArbitrary(0,1),
                time: d.toTimeString().split(' ')[0],
                date: d.toISOString().split('T')[0]
            }
                ];


        connection.send(JSON.stringify(wsdata));
        
    },1000);
    

    connection.on('message', function(message) {
           var data = JSON.parse(message.utf8Data); 
        
            
            //console.log('Received Message: ' + data[0].startdate);
            //console.log('Received Message: ' + data[0].enddate);


            const  sql = 'SELECT * from pulangi WHERE date BETWEEN \''+ data[0].startdate+ '\'' + ' AND \''+data[0].enddate+ '\''+' ORDER by date ASC';
            
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
