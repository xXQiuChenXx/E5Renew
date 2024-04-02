// Copyright (c) MyItBuilder. All rights reserved.
// Licensed under the MIT license.

import "dotenv/config"
import { settings as settingsType } from "types";

const settings: settingsType = {
    'clientId': process.env.CLIENT_ID || "",
    "clientSecret": process.env.CLIENT_SECRET || "",
    "port": process.env.PORT || "",
    'tenantId': 'common',
    'graphUserScopes': [
        'user.read',
        'mail.read',
        'mail.send',
        "files.read",
        "files.read.all",
        "Files.ReadWrite",
        "Files.ReadWrite.All",
        "offline_access"
    ],
    'folderName': "Dev Folder"
};

export default settings