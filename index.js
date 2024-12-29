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

      user: "sachinsingh2002.ssss@gmail.com",
      pass: "pass pass pass" // add your app password here
    },
  });
}
const transporter = newTransporter();
const sendEmail = async (row) => {
  const { Name, Company, Email, Role, Link } = row; // Adjust column names accordingly
  const nameParts = Name.split(' ');
  const name = nameParts[0];
  const mailOptions = {
    from: 'Sachin Singh <sachinsingh2002.ssss@gmail.com>',
    to: Email,
    subject: `Request for an Interview Opportunity - ${Role} at ${Company}`,
    html: `

    <p>Greetings ${name},</p>
    <p>
      I'm Sachin Singh, a Web Developer at <a href="https://www.rnpsoft.com/">RnPsoft Pvt. Ltd.</a>. 
      I got to know through your LinkedIn post that <b>${Company}</b> is looking for a <b>${Role}</b>, 
      therefore, I have mailed you to tell you about myself. I have:
    </p>
    <ul>
      <li><b>Experience:</b> Web Developer at <a href="https://www.rnpsoft.com/">RnPsoft Pvt. Ltd.</a>, contributing to 
      responsive, cross-browser-compatible web solutions.</li>
      <li><b>Data Analyst:</b> Worked at <a href="https://www.galleonconsultants.com/">Galleon Consultants</a>, building 
      10+ dashboards using <b>Power BI, Tableau, and Excel</b>, improving decision-making by 30%.</li>
      <li>Optimized the order system for <b>Healermate</b>, improving accuracy by 15% using WordPress.</li>
      <li>Worked extensively with <b>JavaScript, React, Next.js, Firebase, Supabase</b>, and more.</li>
      <li>Developed <b>PoemGPT</b>, an AI-driven web app using Next.js, Tailwind CSS, and Replicate API.</li>
      <li>Created a <b>Spam Detection Web App</b> using Python and Flask with 92% accuracy.</li>
      <li>Bachelor's in Computer Science and Engineering from <b>Galgotias University, 2025 Grad</b>.</li>
      <li>Proficient in tools like <b>Power BI, GitHub, and Tableau</b>.</li>
    </ul>