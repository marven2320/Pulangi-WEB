
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
const ADS1115 = require('ads1115');
const i2c = require('i2c-bus');

// Both ADS1115 devices live on the SAME physical I2C bus and are
// distinguished by their address pin wiring (0x48 = ADDR->GND, 0x4B = ADDR->SCL).
const I2C_BUS_NUM = 1;
const ADS1115_ADDR_1 = 0x48; // gates 1: ch 0+1 (differential)
const ADS1115_ADDR_2 = 0x4A; // gate 2: ch 0+1 (differential)
const ADS1115_ADDR_3 = 0x4B; // gate 3: ch 0+1 (differential)
const adc_pin = ['0+1'];

let ads1115_48 = null;
let ads1115_4B = null;
let i2cBus = null;

// Open the bus once and create both device handles on it. Reused for every
// read cycle instead of reopening/leaking a new file descriptor each tick.
i2c.openPromisified(I2C_BUS_NUM).then(async (bus) => {
	i2cBus = bus;

	ads1115_48 = await ADS1115(bus, ADS1115_ADDR_1);
	ads1115_48.gain = 2;

	ads1115_4A = await ADS1115(bus, ADS1115_ADDR_2);
	ads1115_4A.gain = 2;

	ads1115_4B = await ADS1115(bus, ADS1115_ADDR_3);
	ads1115_4B.gain = 2;

	console.log('I2C bus ' + I2C_BUS_NUM + ' ready (0x48, 0x4A, 0x4B)');
}).catch((err) => {
	console.log('Failed to open I2C bus ' + I2C_BUS_NUM + ': ' + err.message);
});

function int16ToDecimal(val) {
    return (val << 16) >> 16;
}

function toVoltage(val) {
    return int16ToDecimal(val)*2*2.048/65535 ;
}

function toGateOpening(val) {
        return toVoltage(val)*5*10*1.25;
}
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
			clientWebSocket.connect('wss://5.0.0.121:8000/', 'echo-protocol', null, null, { rejectUnauthorized: false });
		}
	}catch{}
},10000);
	
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

//Read Data ADC data
setInterval(async function()
{
	if (!ads1115_48 || !ads1115_4B) {
		return; // I2C bus/devices not ready yet
	}

	try {
		// 0x48, ch 0+1 (differential) -> gate 1
		rawdata[0] = await ads1115_48.measure(adc_pin[0]);
		opening[0] = toGateOpening(rawdata[0]);

		// 0x4A, ch 0+1 (differential) -> gate 2
		rawdata[1] = await ads1115_4A.measure(adc_pin[0]);
		opening[1] = toGateOpening(rawdata[1]);

		// 0x4B, ch 0+1 (differential) -> gate 3
		rawdata[2] = await ads1115_4B.measure(adc_pin[0]);
		opening[2] = toGateOpening(rawdata[2]);

		console.log(opening);
		flag = 1;
	} catch (err) {
		console.log('I2C read error: ' + err.message);
	}
}, 2000);
clientWebSocket.connect('wss://5.0.0.121:8000/', 'echo-protocol', null, null, { rejectUnauthorized: false });
