const path = require('path');
const express = require('express');
const bodyParser = require('body-parser');
const favicon = require('serve-favicon');
const exphbs = require('express-handlebars');
const { formatDate, isDivisor2, capitalize } = require('./helpers');

require('./config');

const app = express();
const { PORT } = process.env;

const hbs = exphbs.create({
  helpers: {
    formatDate,
    isDivisor2,
    capitalize,
  },
  defaultLayout: 'main',
  extname: '.hbs',
  partialsDir: path.join(__dirname, '/views/partials/'),
});

// Configure express to use handlebars templates
app.engine('.hbs', hbs.engine);
app.set('view engine', '.hbs');

// Favicon
app.use(favicon(path.join(`${process.cwd()}/public/favicon.ico`)));

// Set folder static files
app.use('/public', express.static(`${process.cwd()}/public`));

// For the verbs HTTP get params
app.use(bodyParser.json()); // to support JSON-encoded bodies
app.use(bodyParser.urlencoded({ // to support URL-encoded bodies
  extended: true,
}));

require('./models/connection');
require('./routes')(app);

app.listen(PORT, () => console.log(`App listening on port ${PORT}!`));