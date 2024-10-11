// const express = require('express');
// const xlsx = require('xlsx');
// const fs = require('fs');
// const path = require('path');
// const ejs = require('ejs');
// const puppeteer = require('puppeteer');
// const bodyParser = require('body-parser');

// const app = express();
// const PORT = 3000;

// // Middleware to serve static files from 'public' and 'uploads'
// app.use(express.static('public'));
// app.use(express.static('uploads'));
// app.use(bodyParser.json()); // Parse JSON bodies

// // Load Excel file
// const workbook = xlsx.readFile('employees.xlsx');
// const sheet_name = workbook.SheetNames[0];
// const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name]);

// app.post('/generate-slips', async (req, res) => {
//     const payDate = req.body.payDate; // Get pay date from request
//     const payPeriod = new Date(payDate).toLocaleString('default', { month: 'long', year: 'numeric' });

//     const browser = await puppeteer.launch();
//     const page = await browser.newPage();

//     for (const employee of data) {
//         const salaryPerMonth = parseFloat(employee[" SALARY  PER MONTH"]) || 0;
//         const reserveSalary7 = parseFloat(employee["RESEVER SALARY 7%"]) || 0;
//         const leaveCutOff = parseFloat(employee[" leave cut off"]) || 0;
//         const lopDays = parseFloat(employee["TAKEN LEAVE"]) || 0;

//         // Calculate total deductions
//         const totalDeductions = (leaveCutOff + reserveSalary7).toFixed(2);
//         // const netSalary = (salaryPerMonth - totalDeductions).toFixed(2);
//         const netSalary = employee["PAYABLE SALARY"].toFixed(2) || 0;
//         const amountInWords = `Indian Rupee ${parseFloat(netSalary).toFixed(2)} Only`;
//         const grossEarnings = salaryPerMonth.toFixed(2); // Assuming gross earnings is the salary per month

//         // Log employee data
//         console.log('Processing Employee:', employee["EMPLOYEE NAME"]);

//         // Render the EJS template
//         const template = fs.readFileSync(path.join(__dirname, 'views', 'salarySlip.ejs'), 'utf-8');
//         const html = ejs.render(template, {
//             name: employee["EMPLOYEE NAME"],
//             employeeId: employee["EMPLOYEE ID"],
//             payPeriod: payPeriod,
//             payDate: payDate,
//             netSalary: netSalary,
//             paidDays: employee["TOTAL WORKING DAY"],
//             lopDays: lopDays,
//             salaryPerMonth: salaryPerMonth.toFixed(2),
//             incomeTax: 'Rs.0.00', // Static value for now
//             reserveSalary7: reserveSalary7.toFixed(2),
//             totalDeductions: totalDeductions,
//             grossEarnings: grossEarnings,
//             totalNetPay: netSalary,
//             amountInWords: amountInWords,
//             lwp: leaveCutOff.toFixed(2)
//         });

//         // Create PDF
//         const pdfPath = path.join(__dirname, 'uploads', `${employee["EMPLOYEE NAME"]}_Salary_Slip.pdf`);
//         await page.setContent(html);
//         await page.pdf({ path: pdfPath, format: 'A4' });

//         console.log('PDF generated for:', employee["EMPLOYEE NAME"]);
//     }

//     await browser.close();
//     res.send('Salary slips generated successfully! You can download them from the uploads directory.');
// });

// app.listen(PORT, () => {
//     console.log(`Server is running on http://localhost:${PORT}`);
// });
const express = require('express');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const ejs = require('ejs');
const puppeteer = require('puppeteer');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const cors = require('cors');

const app = express();
const PORT = 3000;
app.set('view engine', 'ejs');

// Middleware to serve static files from 'public' and 'uploads'
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static('uploads'));
app.use(bodyParser.json()); // Parse JSON bodies
app.use(cors()); // Parse JSON bodies

// Load Excel file
const workbook = xlsx.readFile('employees.xlsx');
const sheet_name = workbook.SheetNames[0];
const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name]);

// Set up nodemailer transporter
const transporter = nodemailer.createTransport({
    service: 'gmail', // Use your email service
    auth: {
        user: 'aazikodhr@gmail.com', // Your email
        pass: 'zzij vdfc vqts pbsi' // Your email password or app password
    }
});

app.post('/generate-slips', async (req, res) => {
    const payDate = req.body.payDate; // Get pay date from request
    const payPeriod = new Date(payDate).toLocaleString('default', { month: 'long', year: 'numeric' });

    const browser = await puppeteer.launch();
    const page = await browser.newPage();

    for (const employee of data) {
        const salaryPerMonth = parseFloat(employee[" SALARY  PER MONTH"]) || 0;
        const reserveSalary7 = parseFloat(employee["RESEVER SALARY 7%"]) || 0;
        const leaveCutOff = parseFloat(employee[" leave cut off"]) || 0;
        const lopDays = parseFloat(employee["TAKEN LEAVE"]) || 0;

        const totalDeductions = (leaveCutOff + reserveSalary7).toFixed(2);
        const netSalary = parseFloat(employee["PAYABLE SALARY"]).toFixed(2) || 0;
        const amountInWords = `Indian Rupee ${parseFloat(netSalary).toFixed(2)} Only`;
        const grossEarnings = salaryPerMonth.toFixed(2);

        const template = fs.readFileSync(path.join(__dirname, 'views', 'salarySlip.ejs'), 'utf-8');
        const html = ejs.render(template, {
            name: employee["EMPLOYEE NAME"],
            employeeId: employee["EMPLOYEE ID"],
            payPeriod: payPeriod,
            payDate: payDate,
            netSalary: netSalary,
            paidDays: employee["TOTAL WORKING DAY"],
            lopDays: lopDays,
            salaryPerMonth: salaryPerMonth.toFixed(2),
            incomeTax: 'Rs.0.00', // Static value for now
            reserveSalary7: reserveSalary7.toFixed(2),
            totalDeductions: totalDeductions,
            grossEarnings: grossEarnings,
            totalNetPay: netSalary,
            amountInWords: amountInWords,
            lwp: leaveCutOff.toFixed(2)
        });

        const pdfPath = path.join(__dirname, 'uploads', `${employee["EMPLOYEE NAME"]}_Salary_Slip.pdf`);
        await page.setContent(html);
        await page.pdf({ path: pdfPath, format: 'A4' });

        console.log('PDF generated for:', employee["EMPLOYEE NAME"]);

        // Send the email with the PDF attached
        const mailOptions = {
            from: 'aazikodhr@gmail.com',
            to: employee["EMAIL"], // Use the email field from Excel
            subject: 'Your Salary Slip',
            text: `Dear ${employee["EMPLOYEE NAME"]},\n\nPlease find attached your salary slip for the period of ${payPeriod}.\n\nBest regards,\nYour Company`,
            attachments: [
                {
                    filename: `${employee["EMPLOYEE NAME"]}_Salary_Slip.pdf`,
                    path: pdfPath
                }
            ]
        };

        // Send the email
        await transporter.sendMail(mailOptions);
        console.log('Email sent to:', employee["EMPLOYEE NAME"]);
    }

    await browser.close();
    res.send('Salary slips generated and emails sent successfully!');
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
