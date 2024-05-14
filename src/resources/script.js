/*

*/
let contextService;
let APIService;


async function init() {
    try {

        console.log('init');

        contextService = await AssetsClientSdk.AssetsPluginContext.get();
        APIService = await AssetsClientSdk.AssetsApiClient.fromPluginContext//(contextService);

        const context = contextService.context;
        console.log(context);

        document.getElementById("info").style.display = "none";
        document.getElementById("main").style.display = "block";

        // Add Event
        let dropArea = document.getElementById('droparea');

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, preventDefaults, false)
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            dropArea.addEventListener(eventName, highlight, false)
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropArea.addEventListener(eventName, unhighlight, false)
        });

        dropArea.addEventListener('drop', handleDrop, false);


        document.getElementById("btnUpdate").addEventListener("click", btnUpdate);
        document.getElementById("btnClose").addEventListener("click", btnClose);

        document.getElementById("btnUpdate").style.display = "block";
        //document.getElementById("btnUpdate").style.filter = "blur(1px)";
        //document.getElementById("btnUpdate").disabled = true;
        document.getElementById("btnClose").style.display = "block";

    } catch (error) {
        console.log(error);
        document.getElementById('error').innerHTML = error;

        document.getElementById("droparea").style.display = "none";
        document.getElementById("areaInfo").style.display = "none";
        document.getElementById("btnUpdate").style.display = "none";
        document.getElementById("btnClose").style.display = "none";

    }
}


function handleDrop(e) {
    let dt = e.dataTransfer
    let files = dt.files

    let supportedFormat = ['csv', 'ods', 'xls', 'xlsx'];
    let supportedMIME = ['text/csv',
        'application/vnd.oasis.opendocument.spreadsheet',
        'application/vnd.ms-excel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];

    let reader = new FileReader();

    let fileExt = (files[0].name).substr(((files[0].name).lastIndexOf('.') + 1));
    //console.log(fileExt);

    if (supportedFormat.includes(fileExt)) {
        reader.onload = function (e) {
            var data = "";
            var bytes = new Uint8Array(e.target.result);
            for (var i = 0; i < bytes.byteLength; i++) {
                data += String.fromCharCode(bytes[i]);
            }
            let objJson = tableToJSON(data);
            //console.log(objJson);
        }
    } else {
        document.getElementById('areaDisplay').innerText = "Unsupported file format.";
    };

    reader.readAsArrayBuffer(files[0]);

    document.getElementById("btnUpdate").style.display = "block";
    document.getElementById("btnClose").style.display = "block";
};

async function tableToJSON(data) {
    let workbook = XLSX.read(data, {
        type: 'binary'
    });

    //get the name of First Sheet.
    let Sheet = workbook.SheetNames[0];

    let jsonSheet = XLSX.utils.sheet_to_json(workbook.Sheets[Sheet]);

    document.getElementById('areaDisplay').innerText = JSON.stringify(jsonSheet, null, 2);
    return jsonSheet;
};


async function btnUpdate() {
    try {

        // As of Assets 6.96, Metadata Report column headers uses Field Name instead of Internal Name,
        // "Assets ID" vs. "id".
        // Read from app/locale/messages_en_US.json to get the mapping
        let fieldMapping = await contextService.fetch('/app/locale/messages_en_US.json', {
            method: 'GET',
        });
        //console.log(fieldMapping);
        // TODO: Error checking

        let jsonDataNew = document.getElementById('areaDisplay').innerText;
        document.getElementById('areaDisplay').innerText = "";
        jsonDataNew = JSON.parse(jsonDataNew);
        //console.log(jsonDataNew);

        // Change new "field label" back to "technical label"
        // Get all the new "field label" from the array
        let arrKey = [];
        for (let idx = 0; idx < jsonDataNew.length; idx++) {
            arrKey.push(...(Object.keys(jsonDataNew[idx])));
        };

        // using Set to remove all duplicates
        let setKey = new Set(arrKey);
        arrKey = Array.from(setKey);

        let newMap = {}
        for (let idx = 0; idx < arrKey.length; idx++) {
            let label = Object.keys(fieldMapping).find(key => fieldMapping[key] === arrKey[idx]);
            label = (label.split('.')).slice(-1);
            newMap[arrKey[idx]] = label.toString();
        };

        // Rename the keys
        let jsonData = [];
        for (let idx = 0; idx < jsonDataNew.length; idx++) {
            let objTemp = {};
            for (let item in jsonDataNew[idx]) {
                objTemp[newMap[item]] = jsonDataNew[idx][item];
            }
            jsonData.push(objTemp)
        };
        //console.log(jsonData);

        for (let idx = 0; idx < jsonData.length; idx++) {
            //console.log(Object.keys(jsonData[i]));
            if (jsonData[idx].id === null || typeof jsonData[idx].id === 'undefined') {
                //console.log('Invalid Entry');
                document.getElementById('areaDisplay').innerText += 'Invalid Entry: ' + idx + '\n';
            }
            else {

                let assetId = jsonData[idx].id;
                delete jsonData[idx].id;

                let service = '/services/update';
                let request = await contextService.fetch(service, {
                    method: 'POST',
                    body: {
                        id: assetId,
                        metadata: JSON.stringify(jsonData[idx]),
                    }
                });

                if (request.errorname) {
                    document.getElementById('areaDisplay').innerText += 'updating: ' + assetId + ' ' + request.errorname + '\n';
                } else {
                    document.getElementById('areaDisplay').innerText += 'updating: ' + assetId + ' Ok\n';
                };
            };
        };
        document.getElementById("btnUpdate").style.display = "none";
        document.getElementById("btnClose").style.display = "block";
    } catch (error) {
        console.log('*** CATCH ***');
        console.log(error);
        document.getElementById('error').innerHTML = error;

        document.getElementById("droparea").style.display = "none";
        document.getElementById("areaInfo").style.display = "none";
        document.getElementById("btnUpdate").style.display = "none";
        document.getElementById("btnClose").style.display = "none";
    }
};


function btnClose() {
    contextService.close();
};


function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function highlight(e) {
    let dropArea = document.getElementById('droparea');
    dropArea.classList.add('highlight');
};

function unhighlight(e) {
    let dropArea = document.getElementById('droparea');
    dropArea.classList.remove('highlight');
};


