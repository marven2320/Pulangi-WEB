
/** Global Variables **/

var tag = 0;
var wlevel = 0;
var temp = 0;
var actdtime = 0;
var flag = 1;
var recon;
var errflag1 = 0;
var errflag2 = 0;
var errflag3 = 0;
var serverflag = 0;
var notconnect_counter = 0;
var reboot_counter = 180; //30mins
var query_counter = 0;
var send_loop;
var rawdata = [0, 0, 0];
var opening = [0, 0, 0];
  
/**Startup I2C **/
const i2c = require('i2c-bus');
const { Buffer } = require('node:buffer');
const DS1115_ADDR = 0x48;
const CONVERSION_REG = 0x00;
const CONFIG_REG = 0x01;


setTimeout(function()
{
	const i2c1 = i2c.openSync(1);
	console.log(i2c1.scanSync());
	i2c1.closeSync();
},500);

setTimeout(function()
{
	const i2c1 = i2c.openSync(1);
	i2c1.i2cWriteSync(DS1115_ADDR,1,Buffer.from([CONFIG_REG])); //set read from config register
	i2c1.closeSync();
},500);
setTimeout(function()
{
	const i2c1 = i2c.openSync(1);
	let rawData = i2c1.readWordSync(DS1115_ADDR,CONFIG_REG); //read config register
	console.log(rawData);
	rawData = (rawData >> 8) + ((rawData & 0xff) << 8); //Switch MSB and LSB
	console.log(rawData);
	i2c1.closeSync();
},500);
setTimeout(function()
{
	const i2c1 = i2c.openSync(1);
	i2c1.i2cWriteSync(DS1115_ADDR,1,Buffer.from([CONVERSION_REG])); //set read from conversion register
	i2c1.closeSync();
},500);

const toVoltage = rawData => //Convert rawdata to voltage
{
  rawData = (rawData >> 8) + ((rawData & 0xff) << 8); //Switch MSB and LSB
  //console.log(rawData);
  let voltage = (rawData & 0xffff) / 65535 * 4.096 * 2;
  
  return voltage;
};

const toMeter = voltage => //Convert voltage to m
{
  
  let meter = -0.5033*voltage + 1.2395;
  return meter*10;
};

const toLevel = meter => //Convert voltage to m
{
  
  let level = 287.663 - meter;
  return level;
};
/**end Startup I2C **/

/** Setup Websocket client*/
var WebSocketClient = require('websocket').client;
var sha1 = require('js-sha1');
var shell = require('shelljs');
var clientWebSocket = new WebSocketClient();

clientWebSocket.on('connectFailed', function(error) 
{
    errflag1 = 1;
    serverflag = 0;
    //console.log('Connect Error: ' + error.toString());
});

clientWebSocket.on('connect', function(connection) 
{
    console.log('WebSocket Client Connected');
	clearInterval(send_loop);
    notconnect_counter = 0;
    serverflag = 1;

    connection.on('error', function(error) 
    {
		errflag2 = 1;
		serverflag = 0;
       // console.log("Connection Error: " + error.toString());
    });
    connection.on('close', function() 
    {
		errflag3 = 1;
		serverflag = 0;
       // console.log('echo-protocol Connection Closed');
    });
    connection.on('message', function(message) 
    {
      //  if (message.type === 'utf8') {
         //console.log("Received: " + message.utf8Data);
      // }
    });
    
	send_loop = setInterval(function()
	{
		if (serverflag==1)
		{
			if (flag == 1)
			{
				if (query_counter == 0){
					var d = new Date();
					actdtime = Date.now().toString();
					var client_data = {
						id: "pulangi",
						unitnum: 1,
						tag: 1,
						mw: null,
						mvar: null,
						energy: null,
						freq: null,
						opening: opening[query_counter],
						temp: 0,
						actTime: d.toTimeString().split(' ')[0],
						mydate: d.toISOString().split('T')[0]
						};
					connection.send(JSON.stringify(client_data));
					console.log(JSON.stringify(client_data));
					flag = 0;
				}else if (query_counter == 1){
					var d = new Date();
					actdtime = Date.now().toString();
					var client_data = {
						id: "pulangi",
						unitnum: 2,
						tag: 1,
						mw: null,
						mvar: null,
						energy: null,
						freq: null,
						opening: opening[query_counter],
						temp: 0,
						actTime: d.toTimeString().split(' ')[0],
						mydate: d.toISOString().split('T')[0]
						};
					connection.send(JSON.stringify(client_data));
					console.log(JSON.stringify(client_data));
					flag = 0;
				}else if (query_counter == 2){
					var d = new Date();
					actdtime = Date.now().toString();
					var client_data = {
						id: "pulangi",
						unitnum: 3,
						tag: 1,
						mw: null,
						mvar: null,
						energy: null,
						freq: null,
						opening: opening[query_counter],
						temp: 0,
						actTime: d.toTimeString().split(' ')[0],
						mydate: d.toISOString().split('T')[0]
						};
					connection.send(JSON.stringify(client_data));
					console.log(JSON.stringify(client_data));
					flag = 0;
					query_counter = -1;
				}
				query_counter++;
			}
		}
	},5000);
});
/** end Setup Websocket client*/

/** reconnect fcn */
setInterval(function(){
	try{
		if ((errflag1==1)||(errflag2==1)||(errflag3==1))
		{
			clearInterval(send_loop);
			notconnect_counter++;
			if (notconnect_counter==1)
			{
				console.log("Connecting......");
			}
			console.log("Failed Attempts Counter: "+ notconnect_counter);
			if (notconnect_counter>reboot_counter)
			{
				shell.exec("sudo reboot");
			}
			errflag1 = 0;
			errflag2 = 0;
			errflag3 = 0;
			clientWebSocket.connect('ws://5.0.0.121:8000/','echo-protocol');
		}
	}catch{}
},10000);
	
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

//Read Data ADC data
setInterval(function()
{
  const mux = [0xC2, 0xD2, 0xE2];
  let voltage = [0, 0, 0];
  for(int i=0;i<3;i++)
  {
    setTimeout(function()	
    {
      const i2c1 = i2c.openSync(1);
      /*Set to continuous conversion mode[bit8=0b0],
    	 * FSR +-4.096V[bit11:9=0b001], 
    	 * AINx&GND[bit14:12=0b1xx], x=0b00,0b01,0b10,0b11;
    	 * 250SPS[bit7:5=0b101]
  	  */
  	  i2c1.i2cWriteSync(DS1115_ADDR,3,Buffer.from([CONFIG_REG, mux[i], 0xA3])); 
      rawdata[i] = i2c1.readWordSync(DS1115_ADDR, CONVERSION_REG);
      voltage[i] = toVoltage(rawdata[i]);
	  opening[i] = voltage[i]*10; //percent opening
      console.log(voltage[i] + 'V');
      i2c1.closeSync();
    },500);
  }
	flag = 1;
}, 2000);
clientWebSocket.connect('ws://5.0.0.121:8000/','echo-protocol');
