class Slack {
  private static instances: { [key: string]: Slack } = {};

  private name: string;
  private service: GoogleAppsScriptOAuth2.OAuth2Service;
  private isClassic: boolean;

  private constructor(name: string) {
    const store = PropertiesService.getScriptProperties();

    const id = store.getProperty(`${name}.id`) as string;
    const secret = store.getProperty(`${name}.secret`) as string;
    let scope = store.getProperty(`${name}.scope`) as string;
    if (scope === '""') {
      scope = "";
    }
    let userScope = store.getProperty(`${name}.userScope`) as string;
    if (userScope === '""') {
      userScope = "";
    }

    const isClassic = userScope === null;
    const authorizationBaseUrl: string = isClassic
      ? "https://slack.com/oauth/authorize"
      : "https://slack.com/oauth/v2/authorize";
    const tokenUrl: string = isClassic
      ? "https://slack.com/api/oauth.access"
      : "https://slack.com/api/oauth.v2.access";

    this.name = name;
    this.service = OAuth2.createService(name)
      .setClientId(id)
      .setClientSecret(secret)
      .setAuthorizationBaseUrl(authorizationBaseUrl)
      .setTokenUrl(tokenUrl)
      .setCallbackFunction("Slack.handleCallback")
      .setPropertyStore(store)
      .setScope(scope)
      .setParam("user_scope", userScope);
    this.isClassic = isClassic;
    Slack.instances[name] = this;
  }

  public static getOrCreate(name: string): Slack {
    const instance = Slack.instances[name];
    if (instance === undefined) {
      return new Slack(name);
    } else {
      return instance;
    }
  }

  public static handleCallback(
    callbackRequest: GoogleAppsScript.Events.DoGet
  ): GoogleAppsScript.HTML.HtmlOutput {
    const instance = Slack.getOrCreate(callbackRequest.parameter.serviceName);
    if (instance === undefined) {
      return HtmlService.createHtmlOutput(
        `Unknown service: ${callbackRequest.parameter.serviceName}`
      );
    } else {
      return instance.handleCallback(callbackRequest);
    }
  }

  public handleCallback(
    callbackRequest: GoogleAppsScript.Events.DoGet
  ): GoogleAppsScript.HTML.HtmlOutput {
    const authorized = this.service.handleCallback(callbackRequest);
    if (authorized) {
      return HtmlService.createHtmlOutput(
        `Success! ${this.name} is authorized. You can close this tab.`
      );
    } else {
      return HtmlService.createHtmlOutput(`Denied. You can close this tab`);
    }
  }

  public getAuthorizationUrl(): string {
    return this.service.getAuthorizationUrl();
  }

  public reset(): void {
    this.service.reset();
  }

  public hasAccess(): boolean {
    return this.service.hasAccess();
  }

  private getAccessToken(): string {
    return this.service.getAccessToken();
  }

  private getUserAccessToken(): string {
    if (this.isClassic) {
      return this.service.getAccessToken();
    } else {
      const store = PropertiesService.getScriptProperties();
      const value = store.getProperty(`oauth2.${this.name}`);
      const json = JSON.parse(value);
      return json.authed_user.access_token;
    }
  }

  public get(endpoint: string, ...args: string[]): any {
    try {
      let url = `https://slack.com/api/${endpoint}`;
      if (args) {
        let separator = "?";
        for (const arg of args) {
          url += `${separator}${encodeURIComponent(arg)}`;
          separator = separator === "=" ? "&" : "=";
        }
      }
      const headers = {
        Authorization: `Bearer ${this.getAccessToken()}`,
        "Content-Type": "application/x-www-form-urlencoded",
      };
      console.log({
        url,
        headers,
      });
      const response = UrlFetchApp.fetch(url, {
        headers,
        muteHttpExceptions: true,
      });
      const text = response.getContentText();
      const json = JSON.parse(text);
      if (!json.ok) {
        throw new Error(json.error);
      }
      return json;
    } catch (e) {
      if (e.message === "ratelimited") {
        // Sleep for a minute and try again
        console.log("Rate limited, sleeping for a minute");
        Utilities.sleep(60000);
        return this.get(endpoint, ...args);
      } else {
        throw e;
      }
    }
  }

  public getPaginated(endpoint: string, ...args: string[]): any[] {
    const responses = [this.get(endpoint, ...args)];
    while (
      responses[responses.length - 1].response_metadata.next_cursor !== ""
    ) {
      responses.push(
        this.get(
          endpoint,
          "cursor",
          responses[responses.length - 1].response_metadata.next_cursor,
          ...args
        )
      );
    }
    return responses;
  }
}
