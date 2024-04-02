import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { DeviceCodeInfo, DeviceCodeCredential } from "@azure/identity";
import { User, Message } from "@microsoft/microsoft-graph-types";
import { settings } from "types";
import { readFileSync, readdirSync } from "fs";

class GraphAPIClient {
  #settings: settings;
  #client: Client;
  constructor(settings: settings) {
    this.#settings = settings;
  }

  async login() {
    const settings = this.#settings;
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

    await authProvider.getAccessToken();
  }

  _getClient() {
    return this.#client;
  }
}

export default GraphAPIClient;
