const sentiment = require('sentiment');
const Hapi = require('hapi');
const socketio = require('socket.io');
const Twitter = require('twitter');
const config = require('./config');

const { PORT, HOST } = process.env;

// Client twitter
const client = new Twitter(config);

// Run server
const server = Hapi.server({
  host: HOST,
  port: PORT,
});

// Socket io listen
const io = socketio.listen(server.listener);

let score = 0;
let total_score = 0;
let avg_score = 0;

// Stream twitter search
const stream = client.stream('statuses/filter', { track: 'javascript' });
stream.on('data', function (tweet) {
  // Send tweet
  const result = sentiment(tweet.text);
  tweet['setimental_score'] = result.score;
  io.emit('tweets', tweet);

  score += result.score;
  total_score++;
  avg_score = (score * total_score) / 100;
  io.emit('score', avg_score);
});

/**
* @method register
* @description Register plugins 
*/
async function register() {
  // Register plugins
  await server.register([require('vision'), require('./routes')]);

  // Config environment for views
  server.views({
    engines: {
      html: {
        module: require('handlebars'),
        compileMode: 'sync'
      }
    },
    compileMode: 'async',
    relativeTo: __dirname,
    path: 'templates',
  });
}

/**
* @method start
* @description Start server 
*/
async function start() {
  try {
    // Register plugins
    await register();

    // Start server
    await server.start();
  } catch (err) {
    console.log(err);
    process.exit(1);
  }

  console.log('Server running at:', server.info.uri);
}


// Run application
start();
