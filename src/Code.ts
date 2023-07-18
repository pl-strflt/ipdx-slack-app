const _ = LodashGS.load();

const SCRIPT_ID = ScriptApp.getScriptId();
const CLIENT_ID = PropertiesService.getScriptProperties().getProperty(
  "CLIENT_ID"
) as string;
const CLIENT_SECRET = PropertiesService.getScriptProperties().getProperty(
  "CLIENT_SECRET"
) as string;
// const SCOPE = 'channels:history,channels:read,groups:history,groups:read,usergroups:read,users:read,users:read.email'
const USER_SCOPE =
  "channels:history,channels:read,groups:history,groups:read,search:read,usergroups:read,users:read,users:read.email";
const REDIRECT_URI = `https://script.google.com/macros/d/${SCRIPT_ID}/usercallback`;

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Slack")
    .addItem("Authorize IPDX", "authorize")
    .addItem("Reset Authorization", "reset")
    .addToUi();
}

function getOAuthService() {
  return OAuth2.createService("slack")
    .setAuthorizationBaseUrl("https://slack.com/oauth/v2/authorize")
    .setTokenUrl("https://slack.com/api/oauth.v2.access")
    .setTokenFormat(
      "application/x-www-form-urlencoded" as GoogleAppsScriptOAuth2.TokenFormat
    )
    .setCallbackFunction("authCallback")
    .setPropertyStore(PropertiesService.getUserProperties())
    .setParam("user_scope", USER_SCOPE)
    .setTokenPayloadHandler((payload) => {
      // Otherwise, we get invalid_code error on the callback
      delete (payload as any).client_id;
      return payload;
    })
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET);
}

function authCallback(callbackRequest) {
  const service = getOAuthService();
  const authorized = service.handleCallback(callbackRequest);
  let template;
  if (authorized) {
    const url = `https://slack.com/api/oauth.v2.access`;
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      method: "post",
      muteHttpExceptions: true,
      contentType: "application/x-www-form-urlencoded",
      payload: `code=${
        callbackRequest.parameter.code
      }&client_id=${CLIENT_ID}&client_secret=${CLIENT_SECRET}&redirect_uri=${encodeURIComponent(
        REDIRECT_URI
      )}`,
    };
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (json.ok) {
      const store = PropertiesService.getUserProperties();
      store.setProperty(
        "SLACK_USER_ACCESS_TOKEN",
        json.authed_user.access_token
      );
      // store.setProperty('SLACK_BOT_ACCESS_TOKEN', json.access_token);
      try {
        const test = _get("auth.test");
        store.setProperty("SLACK_USER_ID", test.user_id);
        template = HtmlService.createTemplateFromFile("Success");
        template.user = test.user;
      } catch (e) {
        template = HtmlService.createTemplateFromFile("Failure");
        template.code = JSON.stringify(e, null, 2);
      }
    } else {
      template = HtmlService.createTemplateFromFile("Failure");
      template.code = JSON.stringify(
        [callbackRequest, url, options, json],
        null,
        2
      );
    }
  } else {
    template = HtmlService.createTemplateFromFile("Failure");
    template.code = JSON.stringify(callbackRequest, null, 2);
  }
  const page = template.evaluate();
  SpreadsheetApp.getUi().showModalDialog(page, "Authorization Result");
  return page;
}

function authorize() {
  const service = getOAuthService();
  const template = HtmlService.createTemplateFromFile("Authorize");
  template.authorizationUrl = service.getAuthorizationUrl();
  const page = template.evaluate();
  SpreadsheetApp.getUi().showModalDialog(page, "Authorize IPDX");
}

function reset() {
  const service = getOAuthService();
  service.reset();
}

function _get(endpoint, ...args) {
  const store = PropertiesService.getUserProperties();
  let url = `https://slack.com/api/${endpoint}`;
  if (args) {
    let separator = "?";
    for (const arg of args) {
      url += `${separator}${encodeURIComponent(arg)}`;
      separator = separator === "=" ? "&" : "=";
    }
  }
  console.log(url);
  const headers = {
    Authorization: `Bearer ${store.getProperty("SLACK_USER_ACCESS_TOKEN")}`,
    "Content-Type": "application/x-www-form-urlencoded",
  };
  console.log(headers);
  const response = UrlFetchApp.fetch(url, { headers });
  const text = response.getContentText();
  const json = JSON.parse(text);
  if (!json.ok) {
    throw new Error(json.error);
  }
  return json;
}

function get(endpoint, key, ...args) {
  let response = _get(endpoint, ...args);
  console.log(response);
  if (!args.includes("limit")) {
    while (response?.response_metadata?.next_cursor) {
      const nextResponse = _get(
        endpoint,
        "cursor",
        response.response_metadata.next_cursor,
        ...args
      );
      response = _.mergeWith(response, nextResponse, (objValue, srcValue) => {
        if (_.isArray(objValue)) {
          return objValue.concat(srcValue);
        }
      });
    }
  }
  if (key) {
    if (!_.isArray(key)) {
      key = [key];
    }
    for (const part of key) {
      response = response[part];
    }
  }
  if (!_.isArray(response)) {
    response = [response];
  }
  console.log(response);
  const headers = response.length ? Object.keys(response[0]) : [];
  const rows = [headers];
  response = response.map((item: any) => {
    const row: any[] = [];
    for (const header of headers) {
      const value = item[header];
      if (_.isObject(value)) {
        row.push(JSON.stringify(value));
      } else {
        row.push(value);
      }
    }
    rows.push(row);
  });
  return rows;
}

function conversationsList(...args) {
  return get("conversations.list", "channels", ...args);
}

function conversationsHistory(...args) {
  return get("conversations.history", "messages", ...args);
}

function filterByName(data, ...args) {
  if (_.isArray(data) && data.length) {
    const indices = args
      .map((arg) => data[0].indexOf(arg))
      .filter((arg) => arg !== -1);
    return data.map((row) => indices.map((index) => row[index]));
  } else {
    return data;
  }
}
