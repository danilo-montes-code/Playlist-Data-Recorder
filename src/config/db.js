require('dotenv').config(); // env vars
import '../models'; // generate models

import { connect } from 'mongoose';

connect(process.env.MONGODB_URI, 
  {
    useUnifiedTopology: true,
    useNewUrlParser: true,
  },
  (error) => {
    if (error) console.log(error);
    else       console.log('connected to db');
});