import "@k2oss/k2-broker-core";

  metadata = {
  "systemName": "com.k2.Get-OAuthToken",
  "displayName": "JSSP - GetOAuth Token",
  "description": "Grabs OAuth Token to pass to call to bypass token caching etc.",
  "configuration": {
    "host": {
        "displayName": "host",
        "type": "string",
        "value": "https://login.microsoftonline.com",
        "required": true
        
    },
    "grant_type": {
        "displayName": "grant_type",
        "type": "string",
        "value": "client_credentials",
        "required": true
    },
    "client_id": {
        "displayName": "client_id",
        "type": "string",
        "value": "Application ID from app",
        "required": true
    },
    "tenant": {
      "displayName": "tenant",
      "type": "string",
      "value": "The directory tenant that you want to request permission from.Application ID from app",
      "required": true
  },
    "client_secret": {
        "displayName": "client_secret",
        "type": "string",
        "value": "The Application Secret that you generated for your app",
        "required": true
    },
    "scope": {
      "displayName": "scope",
      "type": "string",
      "value": "https://graph.microsoft.com/.default",
      "required": true
  }
}
};

ondescribe = async function ({ configuration }): Promise<void> {
  postSchema({
    objects: {
      todo: {
        displayName: "getOauthToken",
        description: "Grabs OauthToken",
        properties: {
          token_type: {
            displayName: "token_type",
            type: "string",
          },
          expires_in: {
            displayName: "expires_in",
            type: "number",
          },
          ext_expires_in: {
            displayName: "ext_expires_in",
            type: "number",
          },
          access_token: {
            displayName: "access_token",
            type: "string",
          },
        },
        methods: {
          getOauthToken: {
            displayName: "getOauthToken",
            type: "read",
            outputs: ["token_type","expires_in","ext_expires_in","access_token"],
          },
        },
      },
    },
  });
};


onexecute = async function ({
  objectName,
  methodName,
  parameters,
  properties,
  configuration,
  schema,
}): Promise<void> {
  switch (objectName) {
    case "todo":
      await onexecuteGetToken(methodName, properties, parameters, configuration);
      break;
    default:
      throw new Error("The object " + objectName + " is not supported.");
  }
};

async function onexecuteGetToken(
  methodName: string,
  properties: SingleRecord,
  parameters: SingleRecord,
  configuration: SingleRecord
): Promise<void> {
  switch (methodName) {
    case "getOauthToken":
      await onexecuteGetTheToken(parameters, configuration);
      break;
    default:
      throw new Error("The method " + methodName + " is not supported.");
  }
}

function onexecuteGetTheToken(parameters: SingleRecord, configuration:SingleRecord): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function () {
      try {
        if (xhr.readyState !== 4) return;
        if (xhr.status !== 200)
          throw new Error("Failed with status " + xhr.status);
        var obj = JSON.parse(xhr.responseText);
        postResult({
          token_type: obj.token_type,
          expires_in: obj.expires_in,
          ext_expires_in: obj.ext_expires_in,
          access_token: obj.access_token,
        });
        resolve();
      } catch (e) {
        reject(e);
      }
    };
    var url = configuration["host"] + '/' + configuration["tenant"] + '/oauth2/v2.0/token';
    var data ='grant_type=' + configuration["grant_type"] + '&client_id=' + configuration["client_id"] + '&client_secret='+ configuration["client_secret"] + '&scope='+configuration["scope"];
    var dataFinal = encodeURI(data);
    xhr.open("POST", url);
    xhr.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    xhr.send(dataFinal);
  });
}