// import webserver from "./web/main"
import settings from "config";
import GraphAPIClient from "GraphAPI";

async function main() {
    console.log("Welcome to Graph API");
    const client = new GraphAPIClient(settings);
    await client.login();
    // console.log(webserver);
}

main();