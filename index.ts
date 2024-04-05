// import webserver from "./web/main"
import settings from "config";
import GraphAPIClient from "GraphAPI";

async function main() {
  console.log("Welcome to Graph API");
  const client = new GraphAPIClient(settings);
  setInterval(async () => {
    await client.login.bind(client)();
    await client.start.bind(client)();
  }, 30 * 60 * 1000);
  // console.log(webserver);
}

main();
