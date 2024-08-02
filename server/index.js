const imaps = require('imap-simple');
const XLSX = require('xlsx');
const express = require('express');
const cors = require('cors');
require('dotenv').config();

// Connect to IMAP and fetch emails
const fetchEmails = async () => {
    try {
        const connection = await imaps.connect({
            imap: {
                user: process.env.EMAIL,
                password: process.env.PASSWORD,
                host: process.env.IMAP_HOST,
                port: 993,
                tls: true,
                authTimeout: 3000,
                tlsOptions: {
                    rejectUnauthorized: false
                }
            }
        });

        await connection.openBox('INBOX'); // Open the INBOX

        const searchCriteria = ['ALL']; // Fetch all emails
        const fetchOptions = {
            bodies: ['HEADER.FIELDS (FROM TO SUBJECT DATE)'],
            struct: true
        };

        // Fetch messages
        const messages = await connection.search(searchCriteria, fetchOptions);
        connection.end();
        return messages;
    } catch (err) {
        console.error('Error fetching emails:', err);
        return [];
    } 
    
};
const extractEmailData = (messages) => {
    if (!messages || messages.length === 0) {
        console.error('No messages found');
        return [['No emails found']];
    }

    const data = [['From', 'To', 'Subject', 'Date','Body']];
    messages.forEach(message => {

        // Extract headers
        const headersPart = message.parts.find(part => part.which === 'HEADER.FIELDS (FROM TO SUBJECT DATE)');
        const headers = headersPart ? headersPart.body : {};
        const from = headers.from ? headers.from[0] : '';
        const to = headers.to ? headers.to[0] : '';
        const subject = headers.subject ? headers.subject[0] : '';
        const date = headers.date ? headers.date[0] : '';

        // Extract body
        const bodyPart = message.parts.find(part => part.which === 'TEXT');
        const body = bodyPart ? bodyPart.body : '';

        // Add row to data array
        data.push([from, to, subject, date, body]);
    });
    // console.log(messages)

    return data;
};

const saveToExcel = (data) => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Emails');
    XLSX.writeFile(wb, 'emails.xlsx');
    console.log('Emails saved to emails.xlsx');
};

const main = async () => {
    
    const app = express();

    app.use(cors());
    app.use(express.json());
    const port = 5000;
    app.listen(port, () => {
        console.log(`Server is running on http://localhost:${port}`);
      });
    const messages = await fetchEmails();
    const emailData = extractEmailData(messages);
    saveToExcel(emailData);
};

main();
