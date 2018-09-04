'use strict';

const Hapi = require('hapi');
const mysql = require('mysql');

// Create a server with a host and port
const server=Hapi.server({
    host:'localhost',
    port:8000
});

// Add the route
server.route({
    method: 'GET',
    path: '/',
    handler: async (request,h) => {
        const query = new Promise(function(resolve, reject){
            const connection = mysql.createConnection({
                host: 'localhost',
                user: 'root',
                password: '',
                database: 'test',
            });

            connection.connect();
            connection.query('SELECT * FROM clubs_seasons', (error, results) => {
                if (error) reject(err);
                resolve(results);
            });
            connection.end();
        });

        return await query;
    }
});

// Start the server
async function start() {

    try {
        await server.start();
    }
    catch (err) {
        console.log(err);
        process.exit(1);
    }

    console.log('Server running at:', server.info.uri);
};

start();
