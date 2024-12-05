const express = require('express');
const app = express();
const path = require('path');


app.use(express.urlencoded({ extended: false }));
app.use(express.static(path.join(__dirname, 'public')));






app.listen('8000', function () {
    console.log('server up and running....')
})
// JavaScript source code
