//This placeholder test.ts file is a subset from the K2 Broker Template project 
//(https://github.com/K2Documentation/K2Documentation.Samples.JavascriptBroker.Template/blob/master/src/test.ts)
//and is provided here just for reference. It must be modified with the proper objects and methods.
import {test} from 'vitest';
import '@k2oss/k2-broker-core/test-framework';
import './index';

function mock(name: string, value: any) 
{
    global[name] = value;
}

test('describe returns the hardcoded instance', async t => {
    let schema = null;
    mock("postSchema", function (result: any) {
        schema = result;
    });

    await Promise.resolve<void>(
        ondescribe({
            configuration: {},
        })
    );

    t.expect(schema).toStrictEqual({
        objects: {
            todo: {
                displayName: "TODO",
                description: "Manages a TODO list",
                properties: {
                    id: {
                        displayName: "ID",
                        type: "number",
                    },
                    userId: {
                        displayName: "User ID",
                        type: "number",
                    },
                    title: {
                        displayName: "Title",
                        type: "string",
                    },
                    completed: {
                        displayName: "Completed",
                        type: "boolean",
                    },
                },
                methods: {
                    get: {
                        displayName: "Get TODO",
                        type: "read",
                        inputs: ["id"],
                        outputs: ["id", "userId", "title", "completed"],
                    },
                    getParams: {
                        displayName: "Get TODO",
                        type: "read",
                        parameters: {
                            pid: {
                                displayName: "param1",
                                description: "Description Of Param 1",
                                type: "number",
                            },
                        },
                        requiredParameters: ["pid"],
                        outputs: ["id"],
                    },
                },
            },
        },
    });
});

test('execute fails with the wrong parameters', async t => {
    await t.expect(
        onexecute({
            objectName: "test1",
            methodName: "unused",
            parameters: {},
            properties: {},
            configuration: {},
            schema: {},
        })
    ).rejects.toThrow("The object test1 is not supported.");

    await t.expect(
        onexecute({
            objectName: "todo",
            methodName: "test2",
            parameters: {},
            properties: {},
            configuration: {},
            schema: {},
        })
    ).rejects.toThrow("The method test2 is not supported.");
});

test('execute passes with method params', async t => {
    let result: any = null;
    function pr(r: any) {
        result = r;
    }

    mock("postResult", pr);

    await Promise.resolve<void>(
        onexecute({
            objectName: "todo",
            methodName: "getParams",
            parameters: {
                pid: 456,
            },
            properties: {},
            configuration: {},
            schema: {},
        })
    );

    t.expect(result).toStrictEqual({id:456})
});

test('execute passes', async t => {

    let xhr: {[key:string]: any} = null;
    class XHR {
        public onreadystatechange: () => void;
        public readyState: number;
        public status: number;
        public responseText: string;
        private recorder: {[key:string]: any};

        constructor() {
            xhr = this.recorder = {};
            this.recorder.headers = {};
        }

        open(method: string, url: string) {
            this.recorder.opened = {method, url};   
        }

        setRequestHeader(key: string, value: string) {
            this.recorder.headers[key] = value;
        }

        send() {
            queueMicrotask(() =>
            {
                this.readyState = 4;
                this.status = 200;
                this.responseText = JSON.stringify({
                    "id": 123,
                    "userId": 51,
                    "title": "Groceries",
                    "completed": false
                });
                this.onreadystatechange();
                delete this.responseText;
            });
        }
    }

    mock('XMLHttpRequest', XHR);

    let result: any = null;
    function pr(r: any) {
        result = r;
    }

    mock('postResult', pr);

    await Promise.resolve<void>(onexecute({
        objectName: 'todo',
        methodName: 'get',
        parameters: {},
        properties: {
            "id": 123
        },
        configuration: {},
        schema: {}
    }));

    t.expect(xhr).toStrictEqual({
        opened: {
            method: "GET",
            url: "https://jsonplaceholder.typicode.com/todos/123",
        },
        headers: {
            test: "test value",
        },
    });

    t.expect(result).toStrictEqual({
        id: 123,
        userId: 51,
        title: "Groceries",
        completed: false,
    });
});