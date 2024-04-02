import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import { DeviceCodeInfo, DeviceCodeCredential } from "@azure/identity";
import { User, Message } from "@microsoft/microsoft-graph-types";
import { settings } from "types";
import { existsSync, readFileSync, readdirSync, writeFileSync } from "fs";
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
      cache: {
        cachePlugin: {
          async beforeCacheAccess(tokenCacheContext) {
            if (existsSync("./cache.json")) {
              const cache = await readFileSync("./cache.json", "utf-8");
              await tokenCacheContext.tokenCache.deserialize(cache);
            }
          },
          async afterCacheAccess(tokenCacheContext) {
            const cache = await tokenCacheContext.tokenCache.serialize();
            await writeFileSync("./cache.json", cache, "utf-8");
          },
        },
      },
    });

    const res = await testClient.acquireTokenByDeviceCode({
      scopes: settings.graphUserScopes,
      deviceCodeCallback: (response) => {
        console.log(response);
      },
    });

    if(!res?.accessToken) return;

    this.#client = Client.init({
      authProvider: (done) => {
        done(null, res.accessToken)
      }
    })

    const graphAPI = new GraphAPI(this.#client);
    const response = await graphAPI._getInbox();
    console.log(response);
  }

  _getClient() {
    return this.#client;
  }
}

export default GraphAPIClient;
