const express = require('express')
var fs = require('fs');
const path = require('path')
const PORT = process.env.PORT || 5000
const request= require('request')
var bodyParser = require('body-parser');
var multer = require('multer'); // v1.0.5
const storage = multer.memoryStorage();
var upload = multer({ storage }).single('image'); // for parsing multipart/form-data
var empty = require('is-empty');

require('request-debug')(request)

var app=express()
var router = express.Router()

state={
	applicationConfig : {
			clientID: process.env.APPID,
			apppwd: process.env.APPPWD,
			graphEndpoint: "https://graph.microsoft.com/beta",
			redirectUrl:process.env.INSTANCE,
			scope:"user.read Files.ReadWrite.All offline_access"
	},
	auth: {
		token: null,
		clientcode: null,
	},
	filename: ""
}

app.use(bodyParser.json()); // for parsing application/json
app.use(bodyParser.urlencoded({ extended: true })); // for parsing application/x-www-form-urlencoded
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');
app.listen(PORT, () => console.log(`Listening on ${ PORT }`));

app.get('/', (req, res) => {res.redirect('getuser')});
app.get('/getuser', (req,res) => getUserInfo(req,res));
app.get('/token', (req,res) => {
	state.auth.clientcode = req.query.code;
	getToken(state.auth.clientcode, res);
});
app.get('/download', (req,res) => {res.sendFile(path.join(__dirname, 'public', "document.pdf"))})
app.get('/upload', (req,res) => res.render('upload'));
app.post('/upload', upload, (req,res) => uploadToDrive(req,res));
app.get('/logout', (req, res)=>{state.auth.token=null;res.redirect('/')})

function getAcceptance(res){
	res.redirect('https://login.microsoftonline.com/common/oauth2/v2.0/authorize'
	+ '?client_id='+ state.applicationConfig.clientID
	+ '&redirect_uri='+ state.applicationConfig.redirectUrl
	+ '&scope='+ state.applicationConfig.scope
	+ "&response_type=code");
}

function buildTokenRequest (code) {
return	{
	'port': 443,
	'uri':'https://login.microsoftonline.com/common/oauth2/v2.0/token',
	'method': 'POST',
	'headers':{
		'User-Agent':'request',
		'Content-Type': 'application/x-www-form-urlencoded'
	},
	'form':{
		'client_id':state.applicationConfig.clientID,
		'client_secret':state.applicationConfig.apppwd,
		'code':code,
		'redirect_uri':state.applicationConfig.redirectUrl,
		'grant_type':'authorization_code',
		'scope':state.applicationConfig.scope,
	}
}}

function getToken(clientcode, res){
	// check to see if token already exists and clientcode
	// different from the one we're getting
	if (clientcode != state.auth.clientcode) getAcceptance(res);
	if (!state.auth.token)
	{
		request(buildTokenRequest(clientcode),
			(err, httpResponse, body) => {
			if (err) {
			return console.error('upload failed:', err);
			}
			let response = JSON.parse(body);
			if ((httpResponse == 400)&&(response.error=='invalid_grant'))
				{
					// handle error from server when granting token
				}
			newtoken=response.access_token;
			state.auth.token=newtoken;
			if(newtoken) extendTokenLifetime();
			res.redirect('/');
		})
		.on('error', function(err) {
			console.log(err)
		})
		request.end;
	}
	else getAcceptance(res) // needs res to redirect
}

function getUserOptions(tkn) {
	uri=state.applicationConfig.graphEndpoint+"/me";
	return {'uri':uri,
		'headers':{
		'Authorization':'Bearer '+tkn,
		'User-Agent':'request',
		'Content-Type': 'application/json'
	}}
}

function getUserInfo(req,res){
	if (state.auth.token)
	{
		request(getUserOptions(state.auth.token), (err,httpResponse,body) => {
		if (err){return console.error('err:', err)}
		let response = JSON.parse(body);
		response.token=true; // passing to view that we have a token
		res.render('response', {resp: response});
		});
	}
	else
	{
		if (!(empty(state.auth.clientcode)))
		{
			getToken(state.auth.clientcode,res);
		}
		else {
			getAcceptance(res);
		}
	}}

function extendTokenLifetime(){
	requestOptions = {
		'uri': state.applicationConfig.graphEndpoint + "/policies",
		'headers' : {
			'Authorization': "Bearer " + state.auth.token,
			'Content-Type': 'application/json'
		},
		json: true,
		'body':	{
			"displayName":"CustomTokenLifetimePolicy",
			"definition":["{\"TokenLifetimePolicy\":{\"Version\":1,\"MaxAgeSingleFactor	\":true}}"],
			"type":"TokenLifetimePolicy"
		  }
	};
	request(requestOptions, (err,httpResponse,body) => {
		if (err){return console.error('Policy error when trying to extend lifetime is :', err)}
	})
}

function uploadToDrive(req,res){
	tkn=state.auth.token;
	buildReq = {
		'port':443,
		'uri':state.applicationConfig.graphEndpoint
			+ '/me/drive/root:/'+ req.file.originalname +':/content',
		'headers':{
		'Authorization':'Bearer '+tkn,
		'Content-Type':req.file.mimetype
		},
		'body':req.file.buffer
	}
	state.filename = req.file.originalname;
	request.put(buildReq, (err,httpResponse,body) => {
		//PUT /me/drive/root:/FolderA/FileB.txt:/content
		// Content-Type: text/plain
		if (err){return console.error('PUT error is :', err)}
		info = JSON.parse(body)
		uploaded = info.id;
		buildUrl = state.applicationConfig.graphEndpoint + "/drive/items/" + uploaded + "/content?format=pdf";
		res.render('upload', {resp:info})
		request.get(
		{	'uri': buildUrl,
			'encoding' : null,
			'headers':{
				'Authorization':'Bearer '+tkn
			}
		}, (err,httpResponse,body) => {
			const buffer = Buffer.from(body, 'utf8');
			fs.writeFileSync('public/document.pdf', buffer);
		})
	});
}
