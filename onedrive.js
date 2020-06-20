// - customnewtab app ID: 5024a142-154d-45c4-9ca4-2013bca8919a

var queryOptions = {
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
	},
	cancel: function(e) {
		console.log('OneDrive Launch Cancelled!');
	},
	error: function() {
		console.log('OneDrive Launch Failed!');
	}
}

var saveOptions = {
	clientId: "5024a142-154d-45c4-9ca4-2013bca8919a",

    //client_secret: onedrive_client_secret,
    //refresh_token: onedrive_refresh_token,
    //grant_type: 'refresh_token'

	action: "save",
  	//sourceInputElementId: "fileUploadControl",
  	sourceUri: "file:///C:/Users/kasutaja/prog/odt/odt.html",
  	filename: "file.txt",
  	openInNewWindow: false,
	advanced: {
		redirectUri: "https://stargateprovider.github.io/odt/odt.html",
		
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
		console.log(files);
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
	let res = OneDrive.open(saveOptions);

	//const client = OneDrive.init(odOptions);
	//let res = client.api('/me/drive/root/children').get();
	console.log(res);
}

function ODDownload() {
	const url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
	const url2 = 'https://graph.microsoft.com/v1.0/drive/root:/prog/python/uudised/config.json';
    /*let headers = {
        'Authorization': "Bearer " + JSON.parse(body).access_token,
        'Content-Type': mime.getType(file)
    };*/

	fetch(url2, {method: "GET", headers: odOptions})
		.then(data=>{console.log(data)})
		.then(res=>{console.log(res)})
		.catch(error=>{console.log(error)});
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

	fetch(url2, {method: "POST", headers: odOptions})
		.then(data=>{console.log(data)})
		.then(res=>{console.log(res)})
		.catch(error=>console.log(error));

}