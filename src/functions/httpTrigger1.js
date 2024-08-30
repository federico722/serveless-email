const { app } = require('@azure/functions');
const Handlebars = require('handlebars');
const { EmailClient } = require("@azure/communication-email");
const fs = require('fs');
const path = require('path');

const connectionString = process.env.CONNECTIONSTRING;
const client = new EmailClient(connectionString);


const compileTemplate = async (templateName, dataTemplate) => {
    const templatePath = path.join(__dirname, templateName);
    const source = await fs.readFile(templatePath, 'utf-8');
    const template = Handlebars.compile(source);
    return template({ name: dataTemplate.name });
};

const sendEmailWithTimeout = async (emailMessage, timeout = 60000) => {
    return new Promise(async (resolve, reject) => {
        const timer = setTimeout(() => {
            reject(new Error('Email send operation timed out'));
        }, timeout);

        try {
            const poller = await client.beginSend(emailMessage);
            const result = await poller.pollUntilDone();
            clearTimeout(timer);
            resolve(result);
        } catch (error) {
            clearTimeout(timer);
            reject(error);
        }
    });
};

app.http('httpTrigger1', {
    methods: ['POST'],
    handler: async (request, context) => {
        try {
            context.log('HTTP trigger function processing request.');
            
            const requestData = await request.json();
            const { subject, templateName, dataTemplate, to } = requestData;

            context.log('Request data parsed successfully.');

            const html = await compileTemplate(templateName, dataTemplate);
            context.log('Email template compiled successfully.');

            const emailMessage = {
                senderAddress: process.env.SENDERTEXT,
                content: {
                    subject: subject,
                    html: html,
                },
                recipients: {
                    to: [{ address: to }]
                },
            };

            context.log('Initiating email send process...');
            sendEmailWithTimeout(emailMessage, 300000)  // 5 minutes timeout
                .then(result => {
                    context.log('Email sent successfully:', result);
                })
                .catch(error => {
                    context.log.error('Error sending email:', error);
                });

            return { body: `Email send process initiated` };
        } catch (error) {
            context.log.error('Error in httpTrigger1:', error);
            return {
                status: 500,
                body: `An error occurred: ${error.message}`
            };
        }
    }
});
