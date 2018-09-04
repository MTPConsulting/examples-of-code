'use strict';

const fs = require('fs');
const nodemailer = require('nodemailer');

const from = '';
const transporter = nodemailer.createTransport({
	service: 'gmail',
	auth: {
		user: from,
		pass: ''
	}
});

const html = fs.readFile('email.html', function (err, html) {
  const mailOptions = {
		from: from,
		to: '',
		subject: 'Sending Email using Node.js',
		html: html,
	};
		
	transporter.sendMail(mailOptions, function(error, info) {
		if (error) {
			console.log(error);
		} else {
			console.log('Email sent: ' + info.response);
		}
	});
});


