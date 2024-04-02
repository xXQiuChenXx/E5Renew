import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { DeviceCodeInfo, DeviceCodeCredential } from "@azure/identity";
import { User, Message } from "@microsoft/microsoft-graph-types";
import { settings } from "types";
import { readFileSync, readdirSync } from "fs";
import { GraphAPI } from "./utils";
import msal from "@azure/msal-node";

class GraphAPIClient {
  #settings: settings;
  #client: Client;
  constructor(settings: settings) {
    this.#settings = settings;
  }

  async login() {
    const settings = this.#settings;
    const testClient = new msal.PublicClientApplication({
      auth: {
        clientId: settings.clientId,
      },
    })
    const res = await testClient.acquireTokenByDeviceCode({
      scopes: settings.graphUserScopes,
      deviceCodeCallback: (response) => {
        console.log(response)
      }
    })
    console.log(res);

    
    
    return;
    const deviceCodeCredential = new DeviceCodeCredential({
      clientId: settings.clientId,
      tenantId: settings.tenantId,
      userPromptCallback: (info: DeviceCodeInfo) => {
        console.log(info.message);
      },
    });

    const authProvider = new TokenCredentialAuthenticationProvider(
      deviceCodeCredential,
      {
        scopes: settings.graphUserScopes,
      }
    );

    this.#client = Client.initWithMiddleware({
      authProvider: authProvider,
    });

    const accessToken = await authProvider.getAccessToken();
    const graphAPI = new GraphAPI(this.#client);
    const response = await graphAPI._getInbox();
    console.log(response);
  }

  _getClient() {
    return this.#client;
  }
}

export default GraphAPIClient;
