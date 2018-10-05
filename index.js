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
			apppwd:process.env.APPPWD,
			graphEndpoint: "https://graph.microsoft.com/beta",
			redirectUrl:process.env.INSTANCE
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
app.get('/gettoken', (req,res) => {
	getToken(req,res)
	});
app.get('/token', (req,res) => {
	console.log("got clientcode: ", req.query.code);
	state.auth.clientcode = req.query.code;
	res.redirect('/gettoken?code='+ req.query.code);
});
app.get('/download', (req,res) => {res.sendFile(path.join(__dirname, 'public', "document.pdf"))})
app.get('/upload', (req,res) => res.render('upload'));
app.post('/upload', upload, (req,res) => uploadToDrive(req,res));
app.get('/logout', (req, res)=>{state.auth.token=null;res.redirect('/')})

function getAcceptance(res){
	console.log("getting acceptance: ", state.applicationConfig);
	res.redirect('https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id='
	+ state.applicationConfig.clientID + '&scope=user.read Files.ReadWrite.All offline_access&redirect_uri='
	+ state.applicationConfig.redirectUrl + "&response_type=code");
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
		'client_id':this.state.applicationConfig.clientID,
		'client_secret':this.state.applicationConfig.apppwd,
		'code':code,
		'redirect_uri':this.state.applicationConfig.redirectUrl,
		'grant_type':'authorization_code',
		'scope':'user.read Files.ReadWrite Files.ReadWrite.All',
	}
}}

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
		response.token = !(empty(state.auth.token));
		res.render('response', {resp: response});
		});
	}
	else getAcceptance(res);}


function getToken(req,res){
	// check to see if token already exists and clientcode
	// different from the one we're getting
	if (((req.query.code)||state.auth.clientcode)
	    || (!state.auth.token))
	{
		self=this;
		console.log("now having to get token with code: ", self.state.applicationConfig.clientID);
		request(buildTokenRequest(self.state.auth.clientcode),
			(err, httpResponse, body) => {
			if (err) {
			return console.error('upload failed:', err);
			}
			let response = JSON.parse(body);
			if ((httpResponse == 400)&&(response.error=='invalid_grant'))
				console.log("what ?! : ", this.state.auth.clientcode);
				this.state.auth.clientcode = null;
			state.auth.token=response.access_token;
			res.redirect('/');
		})
		.on('error', function(err) {
			console.log(err)
		})
		request.end;
	}
	else getAcceptance(res) // needs res to redirect
}

function uploadToDrive(req,res){
	// make sure we still have a token
	if ((!state.auth.token)||(state.auth.clientcode))
		res.redirect('/gettoken?code='+state.applicationConfig.clientID);
	else if (!state.auth.token) res.redirect('/gettoken');
	tkn=state.auth.token;
	// load the uploaded image as binary data
	// build our request to OneDrive
	buildReq = {
		'port':443,
		'uri':this.state.applicationConfig.graphEndpoint
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
