// Copyright (c) MyItBuilder. All rights reserved.
// Licensed under the MIT license.

import "dotenv/config"

const settings = {
    'clientId': process.env.CLIENT_ID || "",
    "port": process.env.PORT || "",
    'tenantId': 'common',
    'graphUserScopes': [
        'user.read',
        'mail.read',
        'mail.send',
        "files.read",
        "files.read.all",
        "Files.ReadWrite",
        "Files.ReadWrite.All"
    ],
    'folderName': "Dev Folder"
};

export default settings