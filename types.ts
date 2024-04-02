export type settings = {
    'clientId': string,
    "clientSecret": string,
    "port": string,
    'tenantId': string,
    'graphUserScopes': Array<
        'user.read' |
        'mail.read' |
        'mail.send' |
        'files.read' |
        'files.read.all' |
        'Files.ReadWrite' |
        'Files.ReadWrite.All' | 
        "offline_access"
    >,
    'folderName': string
};
