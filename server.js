// server.js
const nodemailer = require("nodemailer");
const ExcelJS = require("exceljs");
const dotenv = require("dotenv");

dotenv.config();

// Configure the transporter
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.GMAIL_USER,
    pass: process.env.GMAIL_PASS,
  },
});

// Function to read Excel file and send emails
const sendEmails = async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("emails.xlsx");
  const worksheet = workbook.worksheets[0]; // Assuming emails are in the first sheet

  // Iterate through rows using the forEach method
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    // Get values from the appropriate columns
    const emailCell = row.getCell(1).value; // Email is in the first column
    const subject = row.getCell(2).value; // Subject is in the second column
    const body = row.getCell(3).value; // Body is in the third column
    const attachment = row.getCell(4).value; // Attachment is in the fourth column

    // Extract email as a string
    let email = "";

    // Check if emailCell is an object (like RichText)
    if (emailCell) {
      email =
        typeof emailCell === "string" ? emailCell : emailCell.text || emailCell;
    }

    // Log the row data for debugging
    console.log(
      `Processing row ${rowNumber}: Email=${email}, Subject=${subject}`
    );

    // Validate the email
    if (email && subject && body) {
      const mailOptions = {
        from: process.env.GMAIL_USER,
        to: email, // Here we pass the email directly
        subject: subject,
        text: body,
        attachments: attachment ? [{ path: attachment }] : [],
      };

      // Check if email is valid before sending
      if (typeof email === "string" && email.trim() !== "") {
        transporter
          .sendMail(mailOptions)
          .then(() => {
            console.log(`Email sent to: ${email}`);
          })
          .catch((error) => {
            console.error(`Error sending email to ${email}: ${error.message}`);
          });
      } else {
        console.error(`Invalid email address: ${email}`);
      }
    } else {
      console.error(`Missing email, subject, or body for row ${rowNumber}`);
    }
  });
};

// Start the email sending process
sendEmails();
