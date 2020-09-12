const open    = require('open');
const msal    = require('@azure/msal-node');
const express = require('express');
const fs      = require('fs');
const path    = require('path');

const cachePath = path.join(require('os').homedir(),'.onedrive-cli.json');
const SERVER_PORT = process.env.PORT || 3000;

function readFile(filename) {
    return new Promise( (resolve,reject) => {
        fs.readFile(filename, (err,data) => err?reject(err):resolve(data) )
    })
}

function writeFile(filename, data) {
    return new Promise( (resolve,reject) => {
        fs.writeFile(filename, data, {encoding: null}, err => err?reject(err):resolve() )
    })
}

const readFromStorage = () => {
    return readFile(cachePath, "utf-8");
};

const writeToStorage = (getMergedState) => {
    return readFromStorage().then(oldFile => {
        const mergedState = getMergedState(oldFile);
        return writeFile(cachePath, mergedState);
    })
};

const cachePlugin = {
    readFromStorage,
    writeToStorage
};

const publicClientConfig = {
    auth: {
        clientId: "592d78a9-a99b-4188-bfe0-8a0331f7ec2d"
    },
    cache: {
        cachePlugin
    },
};

const pca = new msal.PublicClientApplication(publicClientConfig);
const msalTokenCache = pca.getTokenCache();
const scopes = ["User.Read", "Files.Read"];
const app = express();

app.get('/', (req, res) => {
    res.redirect('/login');
});

// Initiates Auth Code Grant
app.get('/login', (req, res) => {
    const authCodeUrlParameters = {
        scopes: scopes,
        redirectUri: "http://localhost:3000/redirect/",
        prompt: "select_account"
        
    };

    // get url to sign user in and consent to scopes needed for application
    pca.getAuthCodeUrl(authCodeUrlParameters)
        .then((response) => {
            res.redirect(response);
        })
        .catch((error) => console.log(JSON.stringify(error)));
});

// Second leg of Auth Code grant
app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        redirectUri: "http://localhost:3000/redirect/",
        scopes: scopes,
    };

    pca.acquireTokenByCode(tokenRequest).then((response) => {
        //console.log("\nResponse: \n:", response);
        res.status(200).send('Login successful. You can close this window now.');
        msalTokenCache.writeToPersistence().then(onAuthComplete);
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});

function login() {
    server = app.listen(3000);

    // We open the browser and wait for the login to complete
    // (when the interactive auth flow above completes)
    (async () => {
        await open('http://localhost:3000');
    })();
}

function onAuthComplete() {
    console.log("You have logged in. You can now call other commands.");
    server.close();
}

function getAuthToken() {
    return msalTokenCache.readFromPersistence()
    .then(() => {
        // Use silent auth flow to get token from cache
        accounts = msalTokenCache.getAllAccounts();

        // Build silent request
        const silentRequest = {
            account: accounts[0], // Index must match the account that is trying to acquire token silently
            scopes: scopes,
        };

        // Acquire Token Silently to be used in MS Graph call
        return pca.acquireTokenSilent(silentRequest)
            .then((response) => {
                //console.log("\nSuccessful silent token acquisition:\nResponse: \n:", response);
                msalTokenCache.writeToPersistence();
                return response.accessToken;
            })
            .catch((error) => {
                console.log('Not authenticated, please login first.');
                //console.log(error);
            });
    })
    .catch((error) => {
        console.log(error);
    });
}

exports.login = login;
exports.getAuthToken = getAuthToken;