const Downloader = require("nodejs-file-downloader");
const { readFile, unlink } = require("fs").promises;

module.exports = class GraphAPI {
  constructor(client) {
    this.lock = false;
    if (!client) {
      throw new Error("Graph has not been initialized for user auth");
    }
    this._userClient = client;
  }

  async _getUser(
    select = ["displayName", "mail", "userPrincipalName", "jobTitle"]
  ) {
    return await this._userClient
      .api("/me")
      // Only request specific properties
      .select(select)
      .get();
  }

  async _getInbox() {
    return await this._userClient
      .api("/me/mailFolders/inbox/messages")
      .select(["from", "isRead", "receivedDateTime", "subject"])
      .top(25)
      .orderby("receivedDateTime DESC")
      .get();
  }

  async getInboxAsync() {
    try {
      const messagePage = await this._getInbox();
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

  async _sendEmail(subject: string, body: string, recipient: string) {
    // Create a new message
    const message = {
      subject: subject,
      body: {
        content: body,
        contentType: "text",
      },
      toRecipients: [
        {
          emailAddress: {
            address: recipient,
          },
        },
      ],
    };

    // Send the message
    return await this._userClient.api("me/sendMail").post({
      message: message,
    });
  }

  async sendEmailAsync() {
    try {
      // Send mail to the signed-in user
      // Get the user for their email address
      const user = await this._getUser();
      const userEmail = user?.mail ?? user?.userPrincipalName;

      if (!userEmail) {
        console.log("Couldn't get your email address, canceling...");
        return;
      }

      await this._sendEmail(
        "Testing Microsoft Graph",
        "Hello world!",
        "qiuchenlau@gmail.com"
      );
      console.log("Mail sent.");
    } catch (err) {
      console.log(`Error sending mail: ${err}`);
    }
  }

  async _getDrive() {
    let res = await this._userClient.api("/me/drive").get();
    return {
      id: res.id,
      name: res.name,
      type: res.driveType,
      quota: res.quota,
    };
  }

  async logDriveInfo() {
    try {
      let drive = await this._getDrive();
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

  async listFolder() {
    let res = await this._userClient.api("/me/drive/root/children").get();
    console.log(res);
  }

  async createFolderIfNotExists(folderName = "Dev Folder") {
    let folder = await this._userClient
      .api(`/me/drive/root/search(q='${folderName}')`)
      .get();
    if (!folder.value.length) {
      await this._userClient.api("/me/drive/root/children").post({
        name: folderName,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      });
    }
  }

  async _getItems(folderName = "Dev Folder") {
    let folder = await this._userClient
      .api(`/me/drive/root/search(q='${folderName}')`)
      .get();
    let id = folder.value[0].id;
    return await this._userClient.api(`/me/drive/items/${id}/children`).get();
  }

  async makeGraphCall() {
    try {
      await this.logDriveInfo();
      await this.createFolderIfNotExists("Dev Folder");
      await this.listAndDel();
      setTimeout(this.uploadFile, random());
      setTimeout(this.downloadFile, random());
    } catch (err) {
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
          ?.map((item, i) => {
            return `${i + 1}. ${item.name} (${item.size} bytes) - ${item.id}`;
          })
          .join("\n")
      );
      for (const item of items.value) {
        await this._userClient
          .api(`/me/drive/items/${item.id}/permanentDelete`)
          .post();
        console.log("Deleted: " + item.name);
      }
    }
  }

  async downloadFile() {
    try {
      const items = await this._getItems();
      const res = await this._userClient
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
    } catch (error) {
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
          this._userClient,
          fileUpload,
          options
        );

      // Do the upload
      const uploadResult = await uploadTask.upload();

      // The response body will be of the corresponding type of the
      // item being uploaded. For OneDrive, this is a DriveItem
      const driveItem = uploadResult.responseBody;
      console.log(`Uploaded file with ID: ${driveItem.id}`);
    } catch (error) {
      console.log(error.message);
    }
    this.lock = false;
    await this.makeGraphCall();
  }
};
