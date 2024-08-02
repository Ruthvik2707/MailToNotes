require('dotenv').config();

module.exports = {
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
};
