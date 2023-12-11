const nodemailer = require('nodemailer');

// Configurar o transporte do nodemailer para usar seu servidor SMTP
let transporter = nodemailer.createTransport({
    host: 'host servidor smtp',
    port: 587, //porta smtp
    secure: false, // true para 465, false para outras portas
    auth: {
        user: 'email',
        pass: 'senha'
    },
    tls: {
        rejectUnauthorized: false
    }
});

function sendLogEmail(logFilePath, date) {
    const mailOptions = {
        from: 'email',
        to: 'email do recebedor',
        subject: 'Erro Detectado - ' + date,
        text: 'Um erro foi encontrado, verifique o arquivo de log em anexo',
        attachments: [
            {
                path: logFilePath
            }
        ]
    };

    transporter.sendMail(mailOptions, function(error, info){
        if (error) {
            console.log('Erro ao enviar e-mail: ', error);
        } else {
            console.log('E-mail enviado: ' + info.response);
        }
    });
}

module.exports.sendLogEmail = sendLogEmail;