const request = require('request');
const { client_id, client_secret, tenant } = require('./azure.json')

let accessToken = ''

function generateToken() {
    const options = {
        method: 'POST',
        url: `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        form: {
            client_id: client_id,
            client_secret: client_secret,
            grant_type: 'client_credentials',
            scope: 'https://graph.microsoft.com/.default'
        },
        json: true
    };

    request(options, function (error, response, body) {
        if (error) throw new Error(error);

        accessToken = body.access_token
        console.log(`> Generated token: ${accessToken}`)
        console.log(' ')
        console.log(`> Preparing to fetch users..`)
        fetchUsers()
    });
}

function fetchUsers() {
    const options = {
        method: 'GET',
        url: `https://graph.microsoft.com/v1.0/users`,
        headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json;odata.metadata=minimal;odata.streaming=true;IEEE754Compatible=false;charset=utf-8'
        }
    };

    request(options, function (error, response, body) {
        if (error || body.error) {
            throw new Error(error || body.error.message);
        }

        console.log(body);
    });

}

generateToken();
