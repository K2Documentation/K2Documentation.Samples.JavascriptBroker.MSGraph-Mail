import "@k2oss/k2-broker-core";

metadata = {
    "systemName": "MSGraphMail",
    "displayName": "Microsoft Graph - Mail Broker",
    "description": "Sample JS Broker for Mail using MSGraph"
    // https://docs.microsoft.com/en-us/graph/outlook-share-messages-folders
};

ondescribe = function() {
    postSchema({ objects: {
                "message": {
                    displayName: "Message",
                    description: "An Outlook message",
                    properties: {
                        "id": {
                            displayName: "ID",
                            type: "string" 
                        },
                        "subject": {
                            displayName: "Subject",
                            type: "string"
                        },
                        "from": {
                            displayName: "From",
                            type: "string"
                        },
                        "body": {
                            displayName: "Body",
                            type: "string"
                        }
                    },
                    methods: {
                        "get": {
                            displayName: "Get Message",
                            type: "read",
                            inputs: ["accessToken","userPrincipalName","id"],
                            requiredInputs: [ "id" , "userPrincipalName"],
                            outputs: [ "id", "subject", "from", "body" ]
                        },
                        "list": {
                            displayName: "Get List",
                            type: "list",
                            inputs: ["accessToken","userPrincipalName"],
                            outputs: [ "id", "subject", "from", "body" ]
                        }
                    }
                }
            }
        }
)};

onexecute = function(objectName, methodName, parameters, properties) {
    switch (objectName)
    {
        case "message": onexecuteMessage(methodName, parameters, properties); break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

 function onexecuteMessage(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName)
    {
        case "get": onexecuteMessageGet(parameters, properties); break;
        case "list": onexecuteMessageList(parameters, properties); break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}

function onexecuteMessageGet(parameters: SingleRecord, properties: SingleRecord) {
    var xhr = new XMLHttpRequest();

    xhr.onreadystatechange = function() {
            if (xhr.readyState !== 4) return;
            if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);

            //console.log(xhr.responseText);
            var obj = JSON.parse(xhr.responseText);
            postResult({
                "id": obj.id,
                "subject": obj.subject,
                "from": obj.from.emailAddress.address,
                "body": obj.body.content
            });            
    };

    var url = "https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(parameters["userPrincipalName"]) + "/mailfolders%28%27Inbox%27%29/messages/" + encodeURIComponent(parameters["id"]);
    //console.log(url);
    xhr.open("GET", url);

    // Use Service Instance OAuth configuration
    //xhr.withCredentials = true;

    // Use Access Token
    xhr.setRequestHeader("Authorization", "Bearer " + encodeURIComponent(parameters["accessToken"]));
    xhr.send();   
}

function onexecuteMessageList(parameters: SingleRecord, properties: SingleRecord) {
    var xhr = new XMLHttpRequest();

    xhr.onreadystatechange = function() {
            if (xhr.readyState !== 4) return;
            if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);

            //console.log(xhr.responseText);
            var obj = JSON.parse(xhr.responseText);
            postResult(obj.map(x => {
                return {
                    "id": x.id,
                    "subject": x.subject,
                    "from": x.from.emailAddress.address,
                    "body": x.body.content
                }
            }));
    };

    // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/map
    var url = "https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(parameters["userPrincipalName"]) + "/mailfolders%28%27Inbox%27%29/messages";
    //console.log(url);
    xhr.open("GET", url);

    // Use Access Tokens
    xhr.setRequestHeader("Authorization", "Bearer " + encodeURIComponent(parameters["accessToken"]));
    xhr.send(); 
}