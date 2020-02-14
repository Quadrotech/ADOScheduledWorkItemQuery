import tl = require('azure-pipelines-task-lib/task');
import nodemailer from "nodemailer";
import * as EmailValidator from 'email-validator';

export function getEmailAddresses(): string[] {
    const emailAddresses: string = tl.getInput('emailAddresses', true)!;
    const splittedAddresses = emailAddresses.split(/[\s,\n]+/); // Split on space, comma and newline
    tl.debug("Splitted Addresses: " + splittedAddresses);
    return splittedAddresses;
}

export function validateEmailAddresses(emailAddresses: string[]): boolean {
    for (let emailAddressIndex in emailAddresses) {
        const emailAddress = emailAddresses[emailAddressIndex];
        tl.debug(`Validating E-Mail address: "${emailAddress}"`);
        if (!EmailValidator.validate(emailAddress)) {
            tl.setResult(tl.TaskResult.Failed, `Invalid Email Address: "${emailAddress}"`);
            return false;
        }
    }
    return true;
}

export function sendQueryResult(htmlContent: string, emailAddresses: string[]) {
    const sendMethod = tl.getInput("sendMethod");
    switch (sendMethod) {
        case "SendGrid": {
            sendMailUsingSendGrid(emailAddresses, htmlContent);
            break;
        }
        case "SMTP": {
            tl.debug("Using SMTP as Transport");
            const transporter = getSMTPTransportOptions();
            const smtpFrom = getSMTPFrom();
            const subject = tl.getInput("subject", true);
            // setup email data with unicode symbols
            const mailOptions = {
                from: smtpFrom,
                to: emailAddresses.join(", "),
                subject: subject,
                html: htmlContent // html body
            };
            tl.debug("From: " + mailOptions.from);
            tl.debug("To: " + mailOptions.to);
            tl.debug("Subject: " + subject);
            tl.debug("HTML: " + htmlContent);
            // send mail with defined transport object
            const info = transporter.sendMail(mailOptions, function (error, info) {
                if (error) {
                    tl.error(error.message);
                    return tl.setResult(tl.TaskResult.Failed, error.message);
                }
                tl.debug(info.response);
            });
            break;
        }
        default: {
            tl.setResult(tl.TaskResult.Failed, 'Sending through ' + sendMethod + ' is not supported.');
            return;
        }
    }
}

export function sendMailUsingSendGrid(emailAddresses: string[], html: string) {
    tl.debug("Using SendGrid as Transport");
    const sendGridEndpoint = tl.getInput("connectedServiceNameSendGrid", true)!;
    const sendGridToken = tl.getEndpointAuthorizationParameter(sendGridEndpoint, "apitoken", false);
    const fromEmail = tl.getEndpointAuthorizationParameter(sendGridEndpoint, "senderEmail", false);
    const fromName = tl.getEndpointAuthorizationParameter(sendGridEndpoint, "senderName", false);
    const fromFull = fromName + " <" + fromEmail + ">";
    const subject = tl.getInput("subject", true);
    const sgMail = require('@sendgrid/mail');
    sgMail.setApiKey(sendGridToken);
    const msg = {
        to: emailAddresses,
        from: fromFull,
        subject: subject,
        html: html,
    };
    tl.debug("To: " + msg.to);
    tl.debug("From: " + msg.from);
    tl.debug("Subject: " + msg.subject);
    tl.debug("HTML: " + html);
    sgMail.send(msg);
}

export function getSMTPTransportOptions(): nodemailer.Transporter {
    const smtpEndpoint = tl.getInput("connectedServiceNameSMTP", true)!;
    const scheme = tl.getEndpointAuthorizationScheme(smtpEndpoint, false);
    let auth: any;
    let smtpServer: string;
    let smtpPort: number;
    let tlsOptions: string;
    switch (scheme) {
        case "None": {
            tl.debug("Using unauthenticated SMTP as transport");
            auth = undefined;
            smtpServer = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpserverNoAuth", true)!;
            smtpPort = Number(tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpportNoAuth", true))!;
            tlsOptions = tl.getEndpointAuthorizationParameter(smtpEndpoint, "tlsOptionsNoAuth", true)!;
            break;
        }
        case "UsernamePassword": {
            tl.debug("Using authenticated SMTP as transport");
            const username = tl.getEndpointAuthorizationParameter(smtpEndpoint, "username", true)!;
            const password = tl.getEndpointAuthorizationParameter(smtpEndpoint, "password", true)!;
            auth = { user: username, pass: password };
            smtpServer = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpserverUserPassword", true)!;
            smtpPort = Number(tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpportUserPassword", true))!;
            tlsOptions = tl.getEndpointAuthorizationParameter(smtpEndpoint, "tlsOptionsUserPassword", true)!;
            break;
        }
        default: {
            tl.setResult(tl.TaskResult.Failed, "Scheme \"" + scheme + "\" is invalid");
            throw new Error("Scheme \"" + scheme + "\" is invalid");
        }
    }
    // create reusable transporter object using the default SMTP transport
    let transporter = nodemailer.createTransport({
        host: smtpServer,
        port: smtpPort,
        secure: tlsOptions == "force",
        auth: auth,
        ignoreTLS: tlsOptions == "ignore"
    });
    tl.debug("Host: " + smtpServer);
    tl.debug("Port: " + smtpPort);
    tl.debug("Secure: " + (tlsOptions == "force"));
    tl.debug("IgnoreTLS: " + auth);
    tl.debug("Auth: " + (tlsOptions == "ignore"));
    return transporter;
}

export function getSMTPFrom(): string {
    const smtpEndpoint = tl.getInput("connectedServiceNameSMTP", true)!;
    const scheme = tl.getEndpointAuthorizationScheme(smtpEndpoint, false)!;
    let smtpFromEmail: string;
    let smtpFromName: string;
    switch (scheme) {
        case "None": {
            tl.debug("Using unauthenticated SMTP as transport");
            smtpFromEmail = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpFromEmailNoAuth", true)!;
            smtpFromName = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpFromNameNoAuth", true)!;
            break;
        }
        case "UsernamePassword": {
            tl.debug("Using authenticated SMTP as transport");
            smtpFromEmail = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpFromEmailUserPassword", true)!;
            smtpFromName = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpFromNameUserPassword", true)!;
            break;
        }
        default: {
            tl.setResult(tl.TaskResult.Failed, "Scheme \"" + scheme + "\" is invalid");
            throw new Error("Scheme \"" + scheme + "\" is invalid");
        }
    }
    return smtpFromName + " <" + smtpFromEmail + ">";
}

