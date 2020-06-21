// - customnewtab app ID: 5024a142-154d-45c4-9ca4-2013bca8919a
const redirectUri = "https://stargateprovider.github.io/odt/odt.html"
var credentials = {client_id:"", user_id:"", access_token:"", code:"", client_secret:""}

var queryOptions = {
	clientId: credentials.client_id,

	//client_secret: onedrive_client_secret,
	//refresh_token: onedrive_refresh_token,
	//grant_type: 'refresh_token'

	action: "query",
	multiSelect: true,
	openInNewWindow: false,
	advanced: {
		redirectUri: redirectUri,
		//redirectUri: "https://localhost",
		queryParameters: "select=id,name,size,file,folder,@microsoft.graph.downloadUrl",
		filter: "folder,.json",
		navigation: {
		  entryLocation: {
			sharePoint: {
			  itemPath: "22D9B7E9A1387531!21975"
			},
			disable: true
		  }
		}
	},

	success: function(files) {
		ODDownload();
		//console.log(files);
		//var file = new Blob([files.value[0]["@microsoft.graph.downloadUrl"]], {type: 'application/json'});

		let a = document.createElement('a');
		//a.href = window.URL.createObjectURL(file);
		a.href = files.value[0]["@microsoft.graph.downloadUrl"];
		a.download = 'config.json';
		a.click();

		alert(a.href);
		//window.webkitURL.createObjectURL(file);
		// var odOptions = {
		// 	clientId: clientId,
		// 	action: "download",
		// 	multiSelect: true,
		// 	openInNewWindow: false,
		// 	advanced: {
		// 		redirectUri: "https://stargateprovider.github.io/customnewtab/odt.html",
		// 		queryParameters: "select=id,name,size,file,folder,@microsoft.graph.downloadUrl",
		// 		filter: "folder,.json",
		// 		accessToken: files.accessToken
		// 	},
		// 	success: 'oneDriveFilePickerSuccess2',
		// 	cancel: 'oneDriveFilePickerCancel',
		// 	error: 'oneDriveFilePickerError'
		// }
		// OneDrive.open(odOptions);
	},
	cancel: function(e) {
		console.log('OneDrive Launch Cancelled!');
	},
	error: function() {
		console.log('OneDrive Launch Failed!');
	}
}

var saveOptions = {
	clientId: credentials.client_id,

	//client_secret: onedrive_client_secret,
	//refresh_token: onedrive_refresh_token,
	//grant_type: 'refresh_token'

	action: "save",
	//sourceInputElementId: "fileUploadControl",
	sourceUri: "file:///C:/Users/kasutaja/prog/odt/odt.html",
	filename: "file.html",
	advanced: {
		redirectUri: redirectUri
	},

	success: function(files) {
		console.log(files);
		// var odOptions = {
		// 	clientId: credentials.client_id,
		// 	action: "download",
		// 	multiSelect: true,
		// 	openInNewWindow: false,
		// 	advanced: {
		// 		redirectUri: "https://stargateprovider.github.io/customnewtab/odt.html",
		// 		queryParameters: "select=id,name,size,file,folder,@microsoft.graph.downloadUrl",
		// 		filter: "folder,.json",
		// 		accessToken: files.accessToken
		// 	},
		// 	success: 'oneDriveFilePickerSuccess2',
		// 	cancel: 'oneDriveFilePickerCancel',
		// 	error: 'oneDriveFilePickerError'
		// }
		// OneDrive.open(odOptions);
	},
	cancel: function(e) {
		console.log('OneDrive Launch Cancelled!');
	},
	error: function() {
		console.log('OneDrive Launch Failed!');
	}
}



function launchOneDrivePicker(){
	let res = OneDrive.open(queryOptions);

	//const client = OneDrive.init(odOptions);
	//let res = client.api('/me/drive/root/children').get();
	console.log("Result: ")
	console.log(res);
}

function saveFiles(){
	let res = OneDrive.save(saveOptions);

	//const client = OneDrive.init(odOptions);
	//let res = client.api('/me/drive/root/children').get();
	console.log(res);
}


async function fetchAsync (url, bodyOptions) {

	defaults = {
		method: 'POST', // *GET, POST, PUT, DELETE, etc.
		//mode: 'no-cors', // no-cors, *cors, same-origin
		//cache: 'no-cache', // *default, no-cache, reload, force-cache, only-if-cached
		//credentials: 'same-origin', // include, *same-origin, omit
		headers: {
		  //'Content-Type': 'application/json'
		  'Content-Type': 'application/x-www-form-urlencoded'
		},
		//redirect: 'follow', // manual, *follow, error
		//referrerPolicy: 'unsafe-url',
		//no-referrer, *no-referrer-when-downgrade, origin, origin-when-cross-origin, same-origin, strict-origin, strict-origin-when-cross-origin, unsafe-url
		body: JSON.stringify(bodyOptions) // body data type must match "Content-Type" header
	};
	/*var compiled = new URL(url);
	Object.keys(options).forEach((key,index) => {
		defaults[key] = options[key];
		compiled.searchParams.append(key, defaults[key]);
	});
	console.log(defaults.client_id);*/

	let response = await fetch(url, defaults);
	let data = await response.json();
	return data;
}

async function requestAsync(url, params) {
	var http = new XMLHttpRequest();
	http.withCredentials = false;
	http.open('POST', url, true);
	http.setRequestHeader('Content-type', 'application/x-www-form-urlencoded');

	http.onreadystatechange = function() {
	    if(http.readyState == 4 && http.status == 200) {
	        console.log(http.responseText);
	    }
	}
	http.send(params);
}



function ODDownload() {
	//const url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
	//const url2 = 'https://graph.microsoft.com/v1.0/drive/root:/prog/python/uudised/config.json';
	/*let headers = {
		'Authorization': "Bearer " + JSON.parse(body).access_token,
		'Content-Type': mime.getType(file)
	};*/

	var options = {
		clientId: credentials.client_id,

		//client_secret: onedrive_client_secret,
		//refresh_token: onedrive_refresh_token,
		//grant_type: 'refresh_token'

		action: "save",
		//sourceInputElementId: "fileUploadControl",
		sourceUri: "file:///C:/Users/kasutaja/prog/odt/odt.html",
		filename: "file.txt",
		advanced: {
			redirectUri: redirectUri,

			navigation: {
			  entryLocation: {
				sharePoint: {
				  itemPath: "22D9B7E9A1387531!21975"
				},
				disable: true
			  }
			}
		}
	}

	//document.getElementById("clientIdDiv").textContent = credentials.client_id;
	//document.getElementById("tokenDiv").textContent = token;
	//document.getElementById("userIdDiv").textContent = userId;
	//console.log("=====");

	var params = "client_id="+credentials.client_id+"&redirect_uri="+redirectUri+"&client_secret="+credentials.client_secret+"&code="+credentials.code+"&grant_type=authorization_code";
	var url = encodeURI("http://login.live.com/oauth20_token.srf/"+params);
	var body = {client_id:credentials.client_id, code:credentials.code, grant_type:"authorization_code", redirectUri:redirectUri};

	fetchAsync(url, body).then(r=>console.log(r));
	//console.log(response);
}


function ODUpload(){
	var file = 'odt.html'; // Filename you want to upload on your local PC
	var onedrive_folder = 'prog'; // Folder name on OneDrive
	var onedrive_filename = 'upload.html'; // Filename on OneDrive
	const url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
	const url2 = 'https://graph.microsoft.com/v1.0/drive/root:/prog/odt.html:/content';
	/*let headers = {
		'Authorization': "Bearer " + JSON.parse(body).access_token,
		'Content-Type': mime.getType(file)
	};*/

	fetch(url2, {method: "POST", headers: saveOptions})
		.then(data=>{console.log(data)})
		.then(res=>{console.log(res)})
		.catch(error=>console.log(error));
}


function setValues() {
	credentials.client_id = document.getElementById("clientIdInput").value;
	credentials.user_id = document.getElementById("userIdInput").value;
	credentials.access_token = document.getElementById("tokenInput").value;
	credentials.code = document.getElementById("authcodeInput").value;
	credentials.client_secret = document.getElementById("secretInput").value;
}

/*document.addEventListener("DOMContentLoaded", function(e) {
	var rawFile = new XMLHttpRequest();
	rawFile.overrideMimeType('application/json');
	rawFile.open("GET", "credentials.json", true);
	rawFile.setRequestHeader('Content-type', 'application/json')
	rawFile.onload = function() {
		if (rawFile.readyState === 4 && rawFile.status == "200") {
			credentials = rawFile.json;
			console.log(rawFile.json);
		}
	}
	rawFile.send(null);
});*/