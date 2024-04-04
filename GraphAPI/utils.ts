import Downloader from "nodejs-file-downloader";
const { readFile, unlink } = require("fs").promises;
import { Client } from "@microsoft/microsoft-graph-client";
import { API } from "./api";

export class GraphAPI {
  private lock: boolean;
  private client: Client;
  private api: API;

  constructor(client: Client) {
    this.lock = false;
    if (!client) {
      throw new Error("Graph has not been initialized for user auth");
    }
    this.client = client;
    this.api = new API(client);
  }

  async getInboxAsync() {
    try {
      const messagePage = await this.api.getInbox();
      const messages = messagePage.value;

      // Output each message's details
      for (const message of messages) {
        console.log(`Message: ${message.subject ?? "NO SUBJECT"}`);
        console.log(`  From: ${message.from?.emailAddress?.name ?? "UNKNOWN"}`);
        console.log(`  Status: ${message.isRead ? "Read" : "Unread"}`);
        console.log(`  Received: ${message.receivedDateTime}`);
      }

      // If @odata.nextLink is not undefined, there are more messages
      // available on the server
      const moreAvailable = messagePage["@odata.nextLink"] != undefined;
      console.log(`\nMore messages available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting user's inbox: ${err}`);
    }
  }

  async sendEmailAsync() {
    try {
      await this.api.sendEmail({
        subject: "Hello Microsoft Graph",
        body: "Dear Sir/Madam, Welcome",
        recipient: "superadmin@myitbuilder.net",
      });
      console.log("Mail sent.");
    } catch (err) {
      console.log(`Error sending mail: ${err}`);
    }
  }

  async logDriveInfo() {
    try {
      let drive = await this.api.getDrive();
      console.log(
        "================================================================"
      );
      console.log("Drive Name " + drive?.name);
      console.log("Drive ID: " + drive?.id);
      console.log("Drive type: " + drive?.type);
      console.log("Drive Quota (Used): " + drive.quota?.used);
      console.log("Drive Quota (Total): " + drive.quota?.total);
      console.log(
        "================================================================"
      );
    } catch (error: any) {
      console.log(error.message);
    }
  }

  async createFolder(folderName = "Dev Folder") {
    let folder = await this.api.findItems({ search: folderName });
    if (!folder.value.length) {
      await this.client.api("/me/drive/root/children").post({
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      });
    }
  }

  async makeGraphCall() {
    try {
      await this.logDriveInfo();
      await this.createFolder("Dev Folder");
      await this.listAndDel();
      setTimeout(this.uploadFile, this.random());
      setTimeout(this.downloadFile, this.random());
    } catch (err: any) {
      console.log(err.message);
    }
  }

  random() {
    const minMilliseconds = 30 * 60 * 100; // 10 minutes in milliseconds
    const maxMilliseconds = 60 * 60 * 1000; // 20 minutes in milliseconds

    // Generate a random decimal number between 0 and 1
    const randomDecimal = Math.random();

    // Scale and shift the random decimal to fit the desired range
    const randomNumber =
      minMilliseconds +
      Math.floor(randomDecimal * (maxMilliseconds - minMilliseconds + 1));

    return randomNumber;
  }

  async listAndDel() {
    if (!this.lock) {
      const items = await this._getItems("Dev Folder");
      console.log(
        `\nFoynded Items (${items?.value?.length}): \n` +
          items?.value
            ?.map((item: any, i: number) => {
              return `${i + 1}. ${item.name} (${item.size} bytes) - ${item.id}`;
            })
            .join("\n")
      );
      for (const item of items.value) {
        await this.client
          .api(`/me/drive/items/${item.id}/permanentDelete`)
          .post();
        console.log("Deleted: " + item.name);
      }
    }
  }

  async downloadFile() {
    try {
      const items = await this._getItems();
      const res = await this.client
        .api(`/me/drive/items/${items.value[0].id}`)
        .get();
      const url = res["@microsoft.graph.downloadUrl"];

      if (url) {
        console.log("Downloading " + res.name);

        this.lock = true;

        const downloader = new Downloader({
          url: url, //If the file name already exists, a new file with the name 200MB1.zip is created.
          directory: "./downloads", //This folder will be created, if it doesn't exist.
          fileName: res.name,
          onProgress: function (percentage, chunk, remainingSize) {
            //Gets called with each chunk.
            console.log(percentage, "%");
            console.log("Remaining bytes: ", remainingSize);
          },
        });
        const { filePath } = await downloader.download();
        await unlink(filePath);
        this.lock = false;
        console.log("Download All done");
      }
    } catch (error: any) {
      console.log(error.message);
    }
  }

  async uploadFile() {
    this.lock = true;
    try {
      const {
        FileUpload,
        OneDriveLargeFileUploadTask,
      } = require("@microsoft/microsoft-graph-client");

      const file = await readFile("./file.zip");
      const fileName = "file.zip";

      const options = {
        // Relative path from root folder
        path: "Dev Folder",
        fileName: fileName,
        rangeSize: 1024 * 1024,
        uploadEventHandlers: {
          // Called as each "slice" of the file is uploaded
          progress: (range, _) => {
            console.log(
              `Uploaded bytes ${range?.minValue} to ${range?.maxValue}`
            );
          },
        },
      };

      // Create FileUpload object
      const fileUpload = new FileUpload(file, fileName, file.byteLength);
      // Create a OneDrive upload task
      const uploadTask =
        await OneDriveLargeFileUploadTask.createTaskWithFileObject(
          this.client,
          fileUpload,
          options
        );

      // Do the upload
      const uploadResult = await uploadTask.upload();

      // The response body will be of the corresponding type of the
      // item being uploaded. For OneDrive, this is a DriveItem
      const driveItem = uploadResult.responseBody;
      console.log(`Uploaded file with ID: ${driveItem.id}`);
    } catch (error: any) {
      console.log(error.message);
    }
    this.lock = false;
    await this.makeGraphCall();
  }
}
