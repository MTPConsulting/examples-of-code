import express from 'express';
import React from 'react';
import { renderToString } from 'react-dom/server';
import mongoose from 'mongoose';
import App from './client/App';
import Html from './client/Html';

const port = 3000;
const server = express();

mongoose.connect(`mongodb://localhost/ssr`);
const Cat = mongoose.model('Cat', { name: String });
const kitty = new Cat({ name: 'Zildjian' });
kitty.save().then(() => console.log('meow'));

server.get('/', (req, res) => {
  const body = renderToString(<App name={kitty.name} />);
  const title = 'Server side Rendering with Styled Components';

  res.send(
    Html({
      title,      
      body,
    })
  );
});

server.listen(port);
console.log(`Serving at http://localhost:${port}`);
