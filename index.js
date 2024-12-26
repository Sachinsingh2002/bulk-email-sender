const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const { exit } = require('process');
// download the above packages


// Load your Excel file
const workbook = xlsx.readFile('./list.xlsx'); // Path for the sheet in your local folder
const sheetName = 'Bytedance'; // Change to the name of your sheet
const worksheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(worksheet);

// Email configuration
const newTransporter = ()=>{
  return nodemailer.createTransport({
    pool: true,
    host: "smtp.gmail.com",
    port: 465,
    secure: true,
    auth: {

      user: 
