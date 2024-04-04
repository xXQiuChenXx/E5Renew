import { Client } from "@microsoft/microsoft-graph-client";

export class API {
  private client: Client;

  constructor(client: Client) {
    this.client = client;
  }

  async getUser(
    select = ["displayName", "mail", "userPrincipalName", "jobTitle"]
  ) {
    return await this.client.api("/me").select(select).get();
  }

  async getInbox() {
    return await this.client
      .api("/me/mailFolders/inbox/messages")
      .select(["from", "isRead", "receivedDateTime", "subject"])
      .top(25)
      .orderby("receivedDateTime DESC")
      .get();
  }

  async sendEmail({
    subject,
    body,
    recipient,
  }: {
    subject: string;
    body: string;
    recipient: string;
  }) {
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
    return await this.client.api("me/sendMail").post({
      message: message,
    });
  }

  async getDrive() {
    let res = await this.client.api("/me/drive").get();
    return {
      id: res.id,
      name: res.name,
      type: res.driveType,
      quota: res.quota,
    };
  }

  async listFolder() {
    const folders = await this.client.api("/me/drive/root/children").get();
    return folders;
  }

  async findChildrens({ folder }: { folder: string }) {
    let items = await this.findItems({ search: folder });
    let id = items.value[0].id;
    return await this.client.api(`/me/drive/items/${id}/children`).get();
  }

  async findItems({ search }: { search: string }) {
    const items = await this.client
      .api(`/me/drive/root/search(q='${search}')`)
      .get();
    return items;
  }
}
