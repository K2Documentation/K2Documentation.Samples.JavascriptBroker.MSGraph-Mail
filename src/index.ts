import "@k2oss/k2-broker-core";

metadata = {
    "systemName": "MSGraphMail",
    "displayName": "Microsoft Graph - Mail Broker",
    "description": "Sample JS Broker for Mail using MSGraph"
    // https://docs.microsoft.com/en-us/graph/outlook-share-messages-folders
};

ondescribe = async function (): Promise<void> {
    postSchema({
        objects: {
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
                        inputs: ["accessToken", "userPrincipalName", "id"],
                        requiredInputs: ["id", "userPrincipalName"],
                        outputs: ["id", "subject", "from", "body"]
                    },
                    "list": {
                        displayName: "Get List",
                        type: "list",
                        inputs: ["accessToken", "userPrincipalName"],
                        outputs: ["id", "subject", "from", "body"]
                    }
                }
            }
        }
    })
};

onexecute = async function ({objectName, methodName, parameters, properties}): Promise<void> {
    switch (objectName) {
        case "message": await onexecuteMessage(methodName, parameters, properties); break;
        default: throw new Error("The object " + objectName + " is not supported.");
    }
}

async function onexecuteMessage(methodName: string, parameters: SingleRecord, properties: SingleRecord): Promise<void> {
    switch (methodName) {
        case "get": await onexecuteMessageGet(parameters, properties); break;
        case "list": await onexecuteMessageList(parameters, properties); break;
        default: throw new Error("The method " + methodName + " is not supported.");
    }
}

function onexecuteMessageGet(parameters: SingleRecord, properties: SingleRecord): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        var xhr = new XMLHttpRequest();

        xhr.onreadystatechange = function () {
            try {
                if (xhr.readyState !== 4) return;
                if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);
                var obj = JSON.parse(xhr.responseText);
                postResult({
                    "id": obj.id,
                    "subject": obj.subject,
                    "from": obj.from.emailAddress.address,
                    "body": obj.body.content
                });
                resolve();
            } catch (error) {
                reject(error);
            }
        };

        var url = "https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(parameters["userPrincipalName"]) + "/mailfolders%28%27Inbox%27%29/messages/" + encodeURIComponent(parameters["id"]);
        xhr.open("GET", url);

        // Use Service Instance OAuth configuration
        //xhr.withCredentials = true;

        // Use Access Token
        xhr.setRequestHeader("Authorization", "Bearer " + encodeURIComponent(parameters["accessToken"]));
        xhr.send();
    });
}

function onexecuteMessageList(parameters: SingleRecord, properties: SingleRecord): Promise<void> {
    return new Promise<void>((resolve, reject) => {
        var xhr = new XMLHttpRequest();

        xhr.onreadystatechange = function () {
            try {
                if (xhr.readyState !== 4) return;
                if (xhr.status !== 200) throw new Error("Failed with status " + xhr.status);

                var obj = JSON.parse(xhr.responseText);
                postResult(obj.map(x => {
                    return {
                        "id": x.id,
                        "subject": x.subject,
                        "from": x.from.emailAddress.address,
                        "body": x.body.content
                    }
                }));
                resolve();
            } catch (error) {
                reject(error);
            }
        };

        // https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/map
        var url = "https://graph.microsoft.com/v1.0/users/" + encodeURIComponent(parameters["userPrincipalName"]) + "/mailfolders%28%27Inbox%27%29/messages";
        xhr.open("GET", url);

        // Use Access Tokens
        xhr.setRequestHeader("Authorization", "Bearer " + encodeURIComponent(parameters["accessToken"]));
        xhr.send();
    });
}