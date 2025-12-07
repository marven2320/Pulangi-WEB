#!/usr/bin/env node
var WebSocketServer = require('websocket').server;
var http = require('https');
const fs = require('fs');
const path = require('path')
const express = require('express');
const serve   = require('express-static');
var clients = [];
var clientsIP = [];


const options = {
  key: fs.readFileSync('key.pem'),
  cert: fs.readFileSync('cert.pem')
};

var app = express();

app.use(serve(__dirname + '/'));

 var server = http.createServer(options, app);
 server.listen(8000, function() {
    console.log((new Date()) + ' Server is listening on port 8000');
});

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
]

//MYSQL
const mysql = require('mysql2');
const pool = mysql.createPool({
    host: 'localhost',
    user: 'root',
    password: 'KaRMSys2025!',
    database: 'pulangi_data',
    connectionLimit: 10
  });

//MODBUSRTU
const ModbusRTU = require("modbus-serial");
// create an empty modbus client
const client =  new ModbusRTU();
client.connectRTUBuffered("/dev/serial0", {baudRate: 9600});
client.setTimeout(500);



console.log("Starting client.js!");
const ipcon = ["5.0.0.85","5.0.0.101","5.0.0.117"]; // ip address of SEL
const clientN = [0, 1, 2];

//SEL DATA
var val;
var P = ["","",""];
var S = ["","",""];
var Q = ["","",""];
var Ptotal = "";
var Qtotal = "";
var VAB = ["","",""];
var VBC = ["","",""];
var VCA = ["","",""];
var Vave = ["","",""];
var IA = ["","",""];
var IB = ["","",""];
var IC = ["","",""];
var Iave = ["","",""];
var F = ["","",""];
var PFA = ["","",""];
var PFB = ["","",""];
var PFC = ["","",""];
var PFave = ["","",""];
var E = ["","",""];
var sec = "", min = "", hour = "", day = "", month = "", year = "";
var servdate = "";
var servtime = "";
var date = "";
var time = "";
var datacomplete = [0,0,0];
//FOR DATABASE TABLE, 42 columns
var table = [
    ['"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"']
];

for (let i = 0;i < ipcon.length; i++)
{
    try{
    client[i].connectTCP(ipcon[i], {port: 502});
    client[i].setTimeout(5000);
    }catch{
        
    }
}
//------start-------- FOR MODBUS COMMUNICATION 
const getMetersValue = async (meters) => {
    try{
        // get value of all meters
        for(let meter of meters) {
            await getMeterValue(meter)
            // wait before get another device
            await sleep(100);
	}
    } catch(e){
        // if error, handle them here (it should not)
        //console.log(e)
    } finally {
        // after get all data from slave, repeat it again
        setImmediate(() => {
            getMetersValue(clientN);
        })
    }
}

const getMeterValue = async (clientn) => {
    try {
        // PSQ = 4040 starting address, 6 memloc, 2 memloc per value
        // energy = 600 starting addres, 2 memloc, 2 memloc per value
        // FREQ = 900 starting addres, 1 memloc, 1 memloc per value
        // PF = 912 starting address, 3 memloc, 1 memloc per value
        // I = 4000 starting addres, 6 memloc, 2 memloc per value
        // V = 4014 starting address, 6 memloc, 2 memloc per value

        //READ ENERGY
        await client.setID(clientn+1);
        val =  await client.readInputRegisters(600, 2);
        E[clientn] = ((val.data[0]<<16)|(val.data[1])|0/100000).toString();
        await sleep(100);

        //READ FREQ
        val =  await client.readInputRegisters(900, 1);
        F[clientn] = ((val.data|0)/100).toFixed(4).toString();
        await sleep(100);

        //READ PF
        val =  await client.readInputRegisters(912, 3);
        PFA[clientn] = ((val.data[0]|0)/100).toFixed(4).toString();
        PFB[clientn] = ((val.data[1]|0)/100).toFixed(4).toString();
        PFC[clientn] = ((val.data[2]|0)/100).toFixed(4).toString();
        PFave[clientn] = (((val.data[0]|0)/100+(val.data[1]|0)/100+(val.data[2]|0)/100)/3).toFixed(4).toString();
        await sleep(100);
        
        //READ I
        val =  await client.readInputRegisters(4000, 6);
        IA[clientn] = ((((val.data[0]<<16)|(val.data[1]))|0)/100000).toFixed(4).toString();
        IB[clientn] = ((((val.data[2]<<16)|(val.data[3]))|0)/100000).toFixed(4).toString();
        IC[clientn] = ((((val.data[4]<<16)|(val.data[5]))|0)/100000).toFixed(4).toString();
        Iave[clientn] = (((((val.data[0]<<16)|(val.data[1]))|0)/100000+(((val.data[2]<<16)|(val.data[3]))|0)/100000+(((val.data[4]<<16)|(val.data[5]))|0)/100000)/3).toFixed(4).toString();
        await sleep(100);

        //READ V
        val =  await client.readInputRegisters(4014, 6);
        VAB[clientn] = ((((val.data[0]<<16)|(val.data[1]))|0)/100).toFixed(4).toString();
        VBC[clientn] = ((((val.data[2]<<16)|(val.data[3]))|0)/100).toFixed(4).toString();
        VCA[clientn] = ((((val.data[4]<<16)|(val.data[5]))|0)/100).toFixed(4).toString();
        Vave[clientn] = (((((val.data[0]<<16)|(val.data[1]))|0)/100+(((val.data[2]<<16)|(val.data[3]))|0)/100+(((val.data[4]<<16)|(val.data[5]))|0)/100)/3).toFixed(4).toString();
        await sleep(100);

        //READ PSQ
        val =  await client.readInputRegisters(4040, 6);
        P[clientn] = ((((val.data[0]<<16)|(val.data[1]))|0)/100000).toFixed(4).toString();
        S[clientn] = ((((val.data[2]<<16)|(val.data[3]))|0)/100000).toFixed(4).toString();
        Q[clientn] = ((((val.data[4]<<16)|(val.data[5]))|0)/100000).toFixed(4).toString();
        if(datacomplete[0] && datacomplete[1] && datacomplete[2]){
            Ptotal = (parseFloat(P[0])+parseFloat(P[1])+parseFloat(P[2])).toFixed(4).toString();
            Qtotal = (parseFloat(Q[0])+parseFloat(Q[1])+parseFloat(Q[2])).toFixed(4).toString();
        }
        await sleep(100);

        //READ DATETIME
        val =  await client.readInputRegisters(200, 6);
        sec = val.data[0].toString().padStart(2, '0');
        min = val.data[1].toString().padStart(2, '0');
        hour = val.data[2].toString().padStart(2, '0');
        day = val.data[3].toString().padStart(2, '0');
        month = val.data[4].toString().padStart(2, '0');
        year = val.data[5].toString().padStart(2, '0');
        // servdate = year + "-" + month + "-" + day;
        // servtime = hour + ":" + min + ":" + sec;
        await sleep(100);
        datacomplete[clientn] = 1;
        var d = new Date();
        wsdata[clientn] = {
            id: "pulangi",
            unitnum: clientn+1,
            mw: P[clientn],
            mvar: Q[clientn],
            freq: F[clientn],
            vab: VAB[clientn],
            vbc: VBC[clientn],
            vca: VCA[clientn],
            ia: IA[clientn],
            ib: IB[clientn],
            ic: IC[clientn],
            pfa: PFA[clientn],
            pfb: PFB[clientn],
            pfc: PFC[clientn],
            time: d.toTimeString().split(' ')[0],
            date: d.toISOString().split('T')[0]
        }
        // console.log(clientn.toString());
        // console.log(E[clientn]);
        // console.log(F[clientn]);
        // console.log(PFA[clientn]);
        // console.log(PFB[clientn]);
        // console.log(PFC[clientn]);
        // console.log(IA[clientn]);
        // console.log(IB[clientn]);
        // console.log(IC[clientn]);
        // console.log(VAB[clientn]);
        // console.log(VBC[clientn]);
        // console.log(VCA[clientn]);
        // console.log(P[clientn]);
        // console.log(S[clientn]);
        // console.log(Q[clientn]);
        // console.log(Ptotal);
        // console.log(Qtotal);
        // console.log(servtime);
        // console.log(time);
        // console.log(servdate);
        // console.log("-----------------------");
        // return the value
        //return clientdata[clientn];
    } catch(e){
        // if error return -1
        client[clientn].connectTCP(ipcon[clientn], {port: 502});
        client[clientn].setTimeout(5000);
        datacomplete[clientn] = 0;
        //console.log(e);
        return -1
    }
}
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// start get value
getMetersValue(clientN);
//------end-------- FOR MODBUS COMMUNICATION 

//------start-------- FOR DATASTORAGE
setInterval(function(){

    //if (datacomplete[0] && datacomplete[1] && datacomplete[2]){
        
        var d = new Date();
        servtime = d.toTimeString().split(' ')[0];
        var year = d.toLocaleString("default", { year: "numeric" });
		var month = d.toLocaleString("default", { month: "2-digit" });
		var day = d.toLocaleString("default", { day: "2-digit" });
		servdate = year + "-" + month + "-" + day;
        

        table.push(['"'+P[0]+'"','"'+P[1]+'"','"'+P[2]+'"','"'+Ptotal+'"','"'+Q[0]+'"','"'+Q[1]+'"','"'+Q[2]+'"','"'+Qtotal+'"','"'+E[0]+'"','"'+E[1]+'"','"'+E[2]+'"',
            '"'+VAB[0]+'"','"'+VAB[1]+'"','"'+VAB[2]+'"','"'+VBC[0]+'"','"'+VBC[1]+'"','"'+VBC[2]+'"','"'+VCA[0]+'"','"'+VCA[1]+'"','"'+VCA[2]+'"',
            '"'+IA[0]+'"','"'+IA[1]+'"','"'+IA[2]+'"','"'+IB[0]+'"','"'+IB[1]+'"','"'+IB[2]+'"','"'+IC[0]+'"','"'+IC[1]+'"','"'+IC[2]+'"',
            '"'+PFA[0]+'"','"'+PFA[1]+'"','"'+PFA[2]+'"','"'+PFB[0]+'"','"'+PFB[1]+'"','"'+PFB[2]+'"','"'+PFC[0]+'"','"'+PFC[1]+'"','"'+PFC[2]+'"','"'+F[0]+'"','"'+F[1]+'"','"'+F[2]+'"','"'+servtime+'"','"'+servdate+'"']);
        if (table[0][table.length-1] == '"0"'){
            table.shift();
        }
        //console.log(table);
    //}
    
},1000);

setInterval(function(){

    savetoDatabase();
    table = [
        ['"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"','"0"']
    ];
    // console.log("Saved to Database!");
},60000);

function savetoDatabase(){

    for(let i = 0;i<table.length;i++){
        // For pool initialization, see above
        const sql = "INSERT INTO `pulangi` (`mw1`,`mw2`,`mw3`,`mw`,`mvar1`,`mvar2`,`mvar3`,`mvar`,`energy1`,`energy2`,`energy3`," +
        "`vab1`,`vab2`,`vab3`,`vbc1`,`vbc2`,`vbc3`,`vca1`,`vca2`,`vca3`,`ia1`,`ia2`,`ia3`,`ib1`,`ib2`,`ib3`,`ic1`,`ic2`,`ic3`,"+
        "`pfa1`,`pfa2`,`pfa3`,`pfb1`,`pfb2`,`pfb3`,`pfc1`,`pfc2`,`pfc3`,`freq1`,`freq2`,`freq3`,`time`,`date`) value ("+table[i].toString()+")";
        pool.query(
            {
              sql,
              // ... other options
            },
            (err, result, fields) => {
              if (err instanceof Error) {
                //console.log(err);
                return;
              }
          
            //   console.log(result);
            //   console.log(fields);
            }
          );
    }   
}
//------end-------- FOR DATASTORAGE

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

wsServer.on('request', function(request) {
    if (!originIsAllowed(request.origin)) {
      // Make sure we only accept requests from an allowed origin
      request.reject();
      console.log((new Date()) + ' Connection from origin ' + request.origin + ' rejected.');
      return;
    }
    var connection = request.accept('echo-protocol',request.origin);
    clients.push(connection);
    clientsIP.push(connection.remoteAddress);
    //console.log(clientsIP);
    //console.log((new Date()) + ' Connection accepted.');
    connection.on('message', function(message) {
        if (message.type === 'utf8') {
            //console.log('Received Message: ' + message.utf8Data);
            connection.sendUTF(message.utf8Data);
	    var data = JSON.parse(message.utf8Data);
	    const sql = 'SELECT * from pulangi WHERE date BETWEEN\''+ data[0].startdate+ '\'' + ' AND \''+data[0].enddate+ '\''+' ORDER by time ASC';
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
	 }
        else if (message.type === 'binary') {
            //console.log('Received Binary Message of ' + message.binaryData.length + ' bytes');
            connection.sendBytes(message.binaryData);
        }
    });
    connection.on('close', function(reasonCode, description) {
        //console.log((new Date()) + ' Peer ' + connection.remoteAddress + ' disconnected.');
	const index = clientsIP.indexOf(connection.remoteAddress);
        if (index > -1) { // only splice array when item is found
            clients.splice(index, 1); // 2nd parameter means remove one item only
            clientsIP.splice(index, 1); // 2nd parameter means remove one item only
        }
    });

    setInterval(function(){
        try{
            for(let i = 0;i < clients.length;i++){
                clients[i].sendUTF(JSON.stringify(wsdata));
                //console.log(JSON.stringify(wsdata[i])); 
            }
        }catch(e){
            //console.log(e);
        }
    }, 1000);
});
//------end-------- FOR WEBSOCKET
