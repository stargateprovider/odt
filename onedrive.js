// - customnewtab app ID: 5024a142-154d-45c4-9ca4-2013bca8919a

var odOptions = {
	clientId: "5024a142-154d-45c4-9ca4-2013bca8919a",

    //client_secret: onedrive_client_secret,
    //refresh_token: onedrive_refresh_token,
    //grant_type: 'refresh_token'

	action: "query",
	multiSelect: true,
	openInNewWindow: false,
	advanced: {
		redirectUri: "https://stargateprovider.github.io/odt/odt.html",
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
	}

	success: 'oneDriveFilePickerSuccess',
	cancel: 'oneDriveFilePickerCancel',
	error: 'oneDriveFilePickerError'
}


var oneDriveFilePickerError = function() {
	console.log('OneDrive Launch Failed!');
}
var oneDriveFilePickerSuccess2 = function(files) {
	console.log(files);
}
var oneDriveFilePickerSuccess = function(files) {
	console.log(files);
	alert(files);
	var file = new Blob([files], {type: 'application/json'});

	let a = document.createElement('a');
	a.href = window.URL.createObjectURL(file);
	a.download = 'abc.txt';
	a.click();
	//window.webkitURL.createObjectURL(file);
	// var odOptions = {
	// 	clientId: "5024a142-154d-45c4-9ca4-2013bca8919a",
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
}
var oneDriveFilePickerCancel = function(e) {
	console.log('OneDrive Launch Cancelled!');
}

function launchOneDrivePicker(){

	//OneDrive.open(odOptions);

	const client = OneDrive.init(odOptions);
	let res = client.api('/me/drive/root/children').get();
	console.log(res);
}


function ODUpload(){
	var fs = require('fs');
	var mime = require('mime');
	var request = require('request');

	var file = 'odt.html'; // Filename you want to upload on your local PC
	var onedrive_folder = 'prog'; // Folder name on OneDrive
	var onedrive_filename = 'upload.html'; // Filename on OneDrive

	request.post({
	    url: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
	    form: odOptions,
	}, function(error, response, body) {
	    fs.readFile(file, function read(e, f) {
	        request.put({
	            url: 'https://graph.microsoft.com/v1.0/drive/root:/' + onedrive_folder + '/' + onedrive_filename + ':/content',
	            headers: {
	                'Authorization': "Bearer " + JSON.parse(body).access_token,
	                'Content-Type': mime.getType(file), // When you use old version, please modify this to "mime.lookup(file)",
	            },
	            body: f,
	        }, function(er, re, bo) {
	            console.log(bo);
	        });
	    });
	});
}