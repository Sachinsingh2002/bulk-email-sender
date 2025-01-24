const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const { exit } = require('process');
// download the above packages


// Load your Excel file
const workbook = xlsx.readFile('./list.xlsx'); // Path for the sheet in your local folder
const sheetName = 'Google'; // Change to the name of your sheet
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
  I'm Sachin Singh, an undergraduate currently pursuing a Bachelor of Technology in Computer Science from 
  <b>Galgotias University</b>. I am passionate about <b>frontend web development</b> and am actively looking for 
  a role in this domain.
</p>
<p>
  I recently completed an internship with <a href="https://www.rnpsoft.com/">RnPsoft Pvt. Ltd.</a>, an IT services and consulting company, where I gained hands-on experience in building responsive and dynamic web solutions using technologies like <b>React</b> and <b>Tailwind CSS</b>. 
</p>
<p>
  Additionally, I worked on several projects, such as:  
  <ul>
    <li><b>PoemGPT:</b> An AI-driven web app built with <b>Next.js</b>, <b>Tailwind CSS</b>, and the <b>Replicate API</b>, allowing users to generate poems in various styles.</li>
    <li><b>Shared Space:</b> A room-finding platform developed using <b>Next.js</b> and <b>Firebase</b> for CRUD operations, featuring a responsive design and dynamic functionality.</li>
    <li><b>Spam Detection App:</b> A machine learning-based web application using <b>Python</b>, <b>Flask</b>, and <b>Naive Bayes</b> to identify spam messages with 92% accuracy.</li>
  </ul>
</p>
<p>
  I am interested in any vacant <b>Frontend Web Developer</b> positions in your company that match my skills. I am eager to bring my creativity, enthusiasm, and technical expertise to your team.
</p>
<p>
  PS: I have attached my <b><a href="https://drive.google.com/file/d/1YXH8nVQfKrDk-YrGgJurS0DGC4YQSSJy/view?usp=sharing">Resume</a></b>, <b><a href="https://www.linkedin.com/in/sachinsingh2002/">LinkedIn Profile</a></b>, and <b><a href="https://github.com/sachinsingh2002">GitHub</a></b> for your reference. If you find me suitable for any opportunities, I would greatly appreciate the chance to discuss how I can contribute to your team.
</p>
<p>
  I am looking forward to your response.
</p>
<p>
  Regards,<br>
  <b>Sachin Singh</b><br>
  <b>Contact No: +91 123456789</b><br>
</p>`

  };

  try {
    await transporter.sendMail(mailOptions);
    console.log('Email sent to', Email);
  } catch (error) {
    console.error('Error sending email:', Email, error);
  }
};

const sendEmailsSynchronously = async () => {
  for (const row of data) {
    await sendEmail(row);
    await new Promise((resolve) => setTimeout(resolve, Math.random()*90000)); // Pause for 1 minute (adjust the duration as needed)
  }
  console.log("Done Sending mails")
  exit()
};

// Call the function to send emails
sendEmailsSynchronously(); 
