const express = require('express')
const cronJob = require('./cronJobs/birthdayWishes.js');

const app = express();
app.use(express.json());

app.use((err, req, res, next) => {
    res.status(err.statusCode || 500).json({
        status: 'error',
        message: err.message
    })
})

module.exports = app;