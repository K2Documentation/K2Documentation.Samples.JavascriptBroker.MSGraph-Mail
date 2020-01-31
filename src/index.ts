import "@k2oss/k2-broker-core";

metadata = {
    "systemName": "MSGraphMail",
    "displayName": "Microsoft Graph - Mail Broker",
    "description": "Sample JS Broker for Mail using MSGraph"
    // https://docs.microsoft.com/en-us/graph/outlook-share-messages-folders
};

ondescribe = function() {
    postSchema({ objects: {
                "com.k2.sample.msgraph.mail.message": {
                    displayName: "Message",
                    description: "An Outlook message",
                    properties: {
                        "com.k2.sample.msgraph.mail.message.id": {
                            displayName: "ID",
                            type: "string" 
                        },
                        "com.k2.sample.msgraph.mail.message.subject": {
                            displayName: "Subject",
                            type: "string"
                        },
                        "com.k2.sample.msgraph.mail.message.from": {
                            displayName: "From",
                            type: "string"
                        },
                        "com.k2.sample.msgraph.mail.message.body": {
                            displayName: "Body",
                            type: "string"
                        }
                    },
                    methods: {
                        "com.k2.sample.msgraph.mail.message.get": {
                            displayName: "Get Message",
                            type: "read",
                            inputs: ["com.k2.sample.msgraph.accessToken","com.k2.sample.msgraph.mail.message.userPrincipalName","com.k2.sample.msgraph.mail.message.id"],
                            requiredInputs: [ "com.k2.sample.msgraph.mail.message.id" , "com.k2.sample.msgraph.mail.message.userPrincipalName"],
                            outputs: [ "com.k2.sample.msgraph.mail.message.id", "com.k2.sample.msgraph.mail.message.subject", "com.k2.sample.msgraph.mail.message.from", "com.k2.sample.msgraph.mail.message.body" ]
                        },
                        "com.k2.sample.msgraph.mail.message.list": {
                            displayName: "Get List",
                            type: "list",
                            inputs: ["com.k2.sample.msgraph.accessToken","com.k2.sample.msgraph.mail.message.userPrincipalName"],
                            outputs: [ "com.k2.sample.msgraph.mail.message.id", "com.k2.sample.msgraph.mail.message.subject", "com.k2.sample.msgraph.mail.message.from", "com.k2.sample.msgraph.mail.message.body" ]
                        }
                    }
                }
            }
        }
)};

onexecute = function(objectName, methodName, parameters, properties) {
    switch (objectName)
    {
        case "com.k2.sample.msgraph.mail.message": onexecuteMessage(methodName, parameters, properties); break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

function onexecuteMessage(methodName: string, parameters: SingleRecord, properties: SingleRecord) {
    switch (methodName)
    {
        case "com.k2.sample.msgraph.mail.message.get": onexecuteMessageGet(parameters, properties); break;
        case "com.k2.sample.msgraph.mail.message.list": onexecuteMessageList(parameters, properties); break;
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
            "com.k2.sample.msgraph.mail.message.id": obj.id,
            "com.k2.sample.msgraph.mail.message.subject": obj.subject,
            "com.k2.sample.msgraph.mail.message.from": obj.from.emailAddress.address,
            "com.k2.sample.msgraph.mail.message.body": obj.body.content
        }); 
    };

    var url = "https://graph.microsoft.com/v1.0/users/" + parameters["com.k2.sample.msgraph.mail.message.userPrincipalName"] + "/mailfolders%28%27Inbox%27%29/messages/" + parameters["com.k2.sample.msgraph.mail.message.id"];
    //console.log(url);
    xhr.open("GET", url);

    // Use Service Instance OAuth configuration
    //xhr.withCredentials = true;

    // Use Access Token
    xhr.setRequestHeader("Authorization", "Bearer " + parameters["com.k2.sample.msgraph.accessToken"]);
    xhr.send();
}

function onexecuteMessageList(parameters: SingleRecord, properties: SingleRecord) {
    var xhr = new XMLHttpRequest();
    
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);

        //console.log(xhr.responseText);
        var obj = JSON.parse(xhr.responseText);
        for (var key in obj) {
            postResult({
            "com.k2.sample.msgraph.mail.message.id": obj[key].id,
            "com.k2.sample.msgraph.mail.message.subject": obj[key].subject,
            "com.k2.sample.msgraph.mail.message.from": obj[key].from.emailAddress.address,
            "com.k2.sample.msgraph.mail.message.body": obj[key].body.content
        }})); 
        // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/map
    };

    var url = "https://graph.microsoft.com/v1.0/users/" + parameters["com.k2.sample.msgraph.mail.message.userPrincipalName"] + "/mailfolders%28%27Inbox%27%29/messages";
    //console.log(url);
    xhr.open("GET", url);

    // Use Access Tokens
    xhr.setRequestHeader("Authorization", "Bearer " + parameters["com.k2.sample.msgraph.accessToken"]);
    xhr.send();
}