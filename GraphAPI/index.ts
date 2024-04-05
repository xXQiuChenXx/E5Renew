import { Client, ClientOptions } from "@microsoft/microsoft-graph-client";
import { User, Message } from "@microsoft/microsoft-graph-types";
import { settings } from "types";
import { existsSync, readFileSync, writeFileSync } from "fs";
import { GraphAPI } from "./utils";
import { PublicClientApplication } from "@azure/msal-node";

class GraphAPIClient {
  private settings: settings;
  private client: Client;

  constructor(settings: settings) {
    this.settings = settings;
  }

  async login() {
    const settings = this.settings;

    const testClient = new PublicClientApplication({
      auth: {
        clientId: settings.clientId,
      },
      cache: {
        cachePlugin: {
          async afterCacheAccess(tokenCacheContext) {
            const cache = await tokenCacheContext.tokenCache.serialize();
            writeFileSync(
              "./cache.json",
              JSON.stringify(JSON.parse(cache), null, 4)
            );
          },
          async beforeCacheAccess(tokenCacheContext) {
            if (existsSync("./cache.json")) {
              const content = readFileSync("./cache.json", "utf-8");
              await tokenCacheContext.tokenCache.deserialize(content);
            }
          },
        },
      },
    });

    let res: any;

    if (existsSync("./cache.json")) {
      const acc = await testClient.getTokenCache().getAllAccounts();
      res = await testClient.acquireTokenSilent({
        scopes: settings.graphUserScopes,
        account: acc[0],
      });
    } else {
      res = await testClient.acquireTokenByDeviceCode({
        scopes: settings.graphUserScopes,
        deviceCodeCallback: (response) => {
          console.log(response);
        },
      });
    }

    if (res.accessToken) {
      this.client = Client.init({
        authProvider: (done) => {
          done(null, res.accessToken);
        },
      });
    } else {
      console.log(
        "Unable to Authenticate Your Account, Please Check Your Configurations then Restart the Application"
      );
    }
  }

  async test() {
    const graphAPI = new GraphAPI(this.client);
    // await graphAPI.sendEmailAsync();
  }

  async start() {
    const graphAPI = new GraphAPI(this.client);
    const {
      downloadFile,
      uploadFile,
      listAndDel,
      logDriveInfo,
      getInboxAsync,
      shuffleArray,
      random,
    } = graphAPI;
    const list = [
      downloadFile,
      uploadFile,
      listAndDel,
      logDriveInfo,
      getInboxAsync,
    ];
    1;
    const fns = shuffleArray(list);
    let i = 1;
    for (const fn of fns) {
      setTimeout(async () => {
        try {
          await fn.bind(graphAPI)();
        } catch (error: any) {
          console.log(error);
        }
      }, random() + i * 60 * 1000);
      i++;
    }
  }

  _getClient() {
    return this.client;
  }
}

export default GraphAPIClient;
