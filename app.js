const express = require('express');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const fs = require('fs');
const { promisify } = require('util');
const streamPipeline = promisify(require('stream').pipeline);
const config = require('./config.json');
const { Client } = require('@microsoft/microsoft-graph-client');

const app = express();

const msalConfig = {
    auth: {
        clientId: config.clientId,
        authority: config.authority,
        clientSecret: config.clientSecret
    },
    system: {
    }
};

async function listAllFiles() {
    try {
      const response = await client.api('/me/drive/root/children').get();
      const fileNames = response.value.map(file => file.name);
      return fileNames;
    } catch (error) {
      console.error('Error listing files:', error);
      return [];
    }
}

async function listAllUsersWithAccess(file) {
  try {
    const response = await client.api(`/me/drive/root:/${file}:/permissions`).get();
      const usersWithAccess = response.value.map(permission => {
          if (permission.grantedTo) {
              if (permission.grantedTo.user) {
                  return permission.grantedTo.user.id;
              } else {
                  return permission.grantedTo[0].user.id;
              }
          } else if (permission.grantedToIdentities && permission.grantedToIdentities.length > 0) {
              return permission.grantedToIdentities[0].user.id;
          } else {
              return null;
          }
      }).filter(displayName => displayName !== null);
    return usersWithAccess;
  } catch (error) {
    console.error(`Error listing users with access to ${file}:`, error);
    return [];
  }
}

async function downloadFile(file) {
    try {
      // Constructing the API endpoint to retrieve the file metadata
      const response = await client.api(`/me/drive/root:/${file}`).get();

      // Extracting the download URL from the metadata
      const downloadUrl = response['@microsoft.graph.downloadUrl'];
  
      if (!downloadUrl) {
        throw new Error(`Failed to get download URL for file ${file}`);
      }
  
      // Fetching the file content using the download URL
      const { default: fetch } = await import('node-fetch');

      const responseStream = await fetch(downloadUrl);
      if (!responseStream.ok) {
        throw new Error(`Failed to download file ${file}: ${responseStream.statusText}`);
      }
      
      // Creating a write stream to save the file
      const writeStream = fs.createWriteStream(file);
      responseStream.body.pipe(writeStream);
  
      console.log(`File ${file} downloaded successfully.`);
    } catch (error) {
      console.error(`Error downloading file ${file}:`, error);
    }
}

async function pollForNewUsers(file) {
    let existingUsers = await listAllUsersWithAccess(file);
    // console.log('Starting poll for new users.....');
    setInterval(async () => {
      try {
        const newUsers = await listAllUsersWithAccess(file);
        // console.log(`Users with access to ${file}:`, newUsers);
        const addedUsers = newUsers.filter(user => !existingUsers.includes(user));
  
        if (addedUsers.length > 0) {
          console.log('New users added to the file:', addedUsers);

        }

        if (newUsers.length < existingUsers.length) {
            console.log('A user was removed. Now, the list of users with access to the file are: ', newUsers);
        }
  
        existingUsers = newUsers;
      } catch (error) {
        console.error('Error while polling for new users:', error);
      }
    }, 1000); // Poll every 1 second
}

const msalClient = new ConfidentialClientApplication(msalConfig);
let accessToken = '';


app.get('/', (req, res) => {
    const authUrlParameters = {
        scopes: ["files.read", "user.read"],
        redirectUri: "http://localhost:3000/redirect",
    };

    msalClient.getAuthCodeUrl(authUrlParameters).then((url) => {
        res.redirect(url);
    }).catch((error) => {
        res.status(500).send(error);
    });
});

app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["files.read", "user.read"],
        redirectUri: "http://localhost:3000/redirect",
    };

    msalClient.acquireTokenByCode(tokenRequest).then((response) => {
        console.log('Access Token Obtained! Ready for serving APIs');
        accessToken = response.accessToken;
        res.send('Login successful!');
    }).catch((error) => {
        console.log(error);
        res.status(500).send('Error completing authentication.');
    });
});

const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
});


app.get('/files', async (req, res) => {
    try {
        let files = await listAllFiles();
        return res.status(200).json({ status: 200, result: {files: files} });
      } catch (error) {
        res.status(500).json({ error: error.message });
      }
});

app.get('/users', async (req, res) => {
    try {
        // let file = 'testing.txt';
        // console.log(req);
        const file = req.query.file;
        let users = await listAllUsersWithAccess(file);
        return res.status(200).json({ status: 200, result: {users: users} });
      } catch (error) {
        res.status(500).json({ error: error.message });
      }
});

app.post('/download/:file', async (req, res) => {
    try {
        const file = req.params.file;
        await downloadFile(file);
        return res.status(200).json({ status: 200, result: "success" });
      } catch (error) {
        res.status(500).json({ error: error.message });
      }
});

app.post('/subscribe/:file', async (req, res) => {
    try {
        const file = req.params.file;
        pollForNewUsers(file);
        console.log('Subscribed to access list of file ', file);
        return res.status(200).json({ status: 200, result: "success" });
      } catch (error) {
        res.status(500).json({ error: error.message });
      }
});

app.listen(3000, () => console.log('App listening on port 3000!'));
