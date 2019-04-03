import tl = require('azure-pipelines-task-lib/task');
import * as EmailValidator from 'email-validator';
import * as azdev from "azure-devops-node-api";
import * as witif from "azure-devops-node-api/interfaces/WorkItemTrackingInterfaces";
import * as witapi from "azure-devops-node-api/WorkItemTrackingApi"
import { WorkItemReference } from 'azure-devops-node-api/interfaces/TestInterfaces';
import nodemailer, { TransportOptions } from "nodemailer";

"use strict";

function getHTMLTable(data: any) {
    let result = ['<table border=1>'];
    for(let row of data) {
        result.push('<tr>');
        for(let cell of row){
            result.push(`<td>${cell}</td>`);
        }
        result.push('</tr>');
    }
    result.push('</table>');
    return result.join('\n');
}

function getEmailAddresses() : string[]
{
    const emailAddresses: string = tl.getInput('emailAddresses', true);
    var splittedAddresses = emailAddresses.split(/[\s,\n]+/); // Split on space, comma and newline
    
    tl.debug("Splitted Addresses: " + splittedAddresses);
    return splittedAddresses;
}

function validateEmailAddresses(emailAddresses: string[]) : boolean
{
    for (var emailAddressIndex in emailAddresses)
    {
        var emailAddress = emailAddresses[emailAddressIndex];

        tl.debug("Validating E-Mail address: \"" + emailAddress + "\"");
        if (!EmailValidator.validate(emailAddress))
        {
            tl.setResult(tl.TaskResult.Failed, 'Invalid Email Address: "' + emailAddress + '"');
            return false;
        }
    }

    return true;
}

function sendMailUsingSendGrid(emailAddresses: string[], html: string)
{
    tl.debug("Using SendGrid as Transport");
    var sendGridEndpoint = tl.getInput("connectedServiceNameSendGrid", true);
    var sendGridToken = tl.getEndpointAuthorizationParameter(sendGridEndpoint, "apitoken", false);

    var fromEmail = tl.getEndpointAuthorizationParameter(sendGridEndpoint, "senderEmail", false);
    var fromName = tl.getEndpointAuthorizationParameter(sendGridEndpoint, "senderName", false);
    var fromFull = fromName + " <" + fromEmail + ">";

    var subject = tl.getInput("subject", true);

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

function getSMTPTransportOptions() : nodemailer.Transporter
{
    var smtpEndpoint = tl.getInput("connectedServiceNameSMTP", true);
    var scheme = tl.getEndpointAuthorizationScheme(smtpEndpoint, false);
    
    let auth: any;
    let smtpServer: string;
    let smtpPort: number;
    let tlsOptions: string;

    switch(scheme)
    {
        case "None": {
            tl.debug("Using unauthenticated SMTP as transport");
            auth = undefined;
            smtpServer = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpserverNoAuth", true);
            smtpPort = Number(tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpportNoAuth", true));
            tlsOptions = tl.getEndpointAuthorizationParameter(smtpEndpoint, "tlsOptionsNoAuth", true);
            break;
        }
        case "UsernamePassword": {
            tl.debug("Using authenticated SMTP as transport");
            var username = tl.getEndpointAuthorizationParameter(smtpEndpoint, "username", true);
            var password = tl.getEndpointAuthorizationParameter(smtpEndpoint, "password", true);
            auth = { user: username, pass: password };

            smtpServer = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpserverUserPassword", true);
            smtpPort = Number(tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpportUserPassword", true));
            tlsOptions = tl.getEndpointAuthorizationParameter(smtpEndpoint, "tlsOptionsUserPassword", true);
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

function getSMTPFrom() : string
{
    var smtpEndpoint = tl.getInput("connectedServiceNameSMTP", true);
    var scheme = tl.getEndpointAuthorizationScheme(smtpEndpoint, false);
    
    let smtpFromEmail: string;
    let smtpFromName: string;

    switch(scheme)
    {
        case "None": {
            tl.debug("Using unauthenticated SMTP as transport");
            smtpFromEmail = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpFromEmailNoAuth", true);
            smtpFromName = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpFromNameNoAuth", true);
            break;
        }
        case "UsernamePassword": {
            tl.debug("Using authenticated SMTP as transport");
            smtpFromEmail = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpFromEmailUserPassword", true);
            smtpFromName = tl.getEndpointAuthorizationParameter(smtpEndpoint, "smtpFromNameUserPassword", true);
            break;
        }
        default: {
            tl.setResult(tl.TaskResult.Failed, "Scheme \"" + scheme + "\" is invalid");
            throw new Error("Scheme \"" + scheme + "\" is invalid");
        }
    }

    return smtpFromName + " <" + smtpFromEmail + ">";
}

async function getQueryResult(connection: azdev.WebApi, projectId: string) : Promise<witif.WorkItemQueryResult>
{
    var queryId = getQueryId();   
    let wit: witapi.IWorkItemTrackingApi = await connection.getWorkItemTrackingApi();
    return wit.queryById(queryId);
}

function getQueryId() : string
{
    var queryType: string = tl.getInput('queryType');
    let queryId: string;

    switch (queryType)
    {
        case "My": {
            queryId = tl.getInput('queryMy', true);
            break;
        }

        case "Shared": {
            queryId = tl.getInput('query', true);
            break;
        }

        default: {
            tl.setResult(tl.TaskResult.Failed, "Query Type \"" + queryType + "\" is invalid");
            throw new Error("Query Type \"" + queryType + "\" is invalid");
        }
    }

    return queryId;
}

function getOrgUrl() : string
{
    var queryType: string = tl.getInput('queryType');

    switch (queryType)
    {
        case "My": {
            var patEndpoint = tl.getInput("connectedServiceNameAzureDevOps", true);
            return tl.getEndpointUrl(patEndpoint, false);
        }

        case "Shared": {
            return tl.getEndpointUrl('SystemVssConnection', false);
        }

        default: {
            tl.setResult(tl.TaskResult.Failed, "Query Type \"" + queryType + "\" is invalid");
            throw new Error("Query Type \"" + queryType + "\" is invalid");
        }
    }
}

function getADOConnection() : azdev.WebApi
{
    var queryType: string = tl.getInput('queryType');

    switch (queryType)
    {
        case "My": {
            var patEndpoint = tl.getInput("connectedServiceNameAzureDevOps", true);
            var scheme = tl.getEndpointAuthorizationScheme(patEndpoint, false);
            var orgUrl = getOrgUrl();

            switch (scheme)
            {
                case "Token":
                {
                    const pat = tl.getEndpointAuthorizationParameter(patEndpoint, "apitoken", false);
                    let authHandler = azdev.getPersonalAccessTokenHandler(pat);
                    return new azdev.WebApi(orgUrl, authHandler);
                }

                case "UsernamePassword":
                {
                    const username = tl.getEndpointAuthorizationParameter(patEndpoint, "username", false);
                    const password = tl.getEndpointAuthorizationParameter(patEndpoint, "password", false);
                    let authHandler = azdev.getBasicHandler(username, password);
                    return new azdev.WebApi(orgUrl, authHandler);
                }

                default: 
                {
                    tl.setResult(tl.TaskResult.Failed, "Authentication Scheme \"" + scheme + "\" is invalid");
                    throw new Error("Authentication Scheme \"" + scheme + "\" is invalid");
                }
            }
        }

        case "Shared": {
            const token = tl.getEndpointAuthorizationParameter('SystemVssConnection', 'AccessToken', false);
            var orgUrl = getOrgUrl();
            
            let authHandler = azdev.getBearerHandler(token);
            return new azdev.WebApi(orgUrl, authHandler);    
        }

        default: {
            tl.setResult(tl.TaskResult.Failed, "Query Type \"" + queryType + "\" is invalid");
            throw new Error("Query Type \"" + queryType + "\" is invalid");
        }
    }
}

async function run() {
    try {
        const projectId: string = tl.getInput('project', true);
        var queryId = getQueryId();
        var orgUrl = getOrgUrl();
        
        var emailAddresses = getEmailAddresses();
        if (!validateEmailAddresses(emailAddresses))
        {
            return;
        }

        var connection = getADOConnection();
        var result = await getQueryResult(connection, projectId);
        
        let wit: witapi.IWorkItemTrackingApi = await connection.getWorkItemTrackingApi();
        var query = await wit.getQuery(projectId, queryId);
        
        tl.debug(JSON.stringify(result));

        let ids: number[] = [];

        if (result == null || result.workItems == undefined)
        {
            return;
        }

        for (let item of result.workItems as Array<WorkItemReference>)
        {
            ids.push(Number(item.id));
        }

        let columnNames: string[] = [];
        if (result.columns != null)
        {
            for (var fieldIndex in result.columns)
            {
                if (result.columns[fieldIndex].referenceName == undefined)
                {
                    continue;
                }

                columnNames.push(result.columns[fieldIndex].referenceName!);
            }
        }

        let workItems : witif.WorkItem[] = [];
        if (ids.length > 0)
        {
            let batchRequest = { fields: columnNames, ids: ids };
            workItems = await wit.getWorkItemsBatch(batchRequest);
        }

        let workItemTableStuff : string[][] = []

        workItemTableStuff.push(columnNames);

        for (var wi in workItems)
        {
            var workItem = workItems[wi];
            if (workItem.fields == undefined)
            {
                continue;
            }

            let fields: string[] = [];
            for (let field in columnNames)
            {
                tl.debug(Object.keys(workItem.fields).toString());

                var fieldName = columnNames[field];
                if (!workItem.fields[fieldName])
                {
                    fields.push("");
                } else {
                    if (fieldName == "System.Id")
                    {
                        fields.push("<a href=\"" + orgUrl + "/" + projectId + "/_workItems/edit/" + workItem.fields[fieldName] + "\">" + workItem.fields[fieldName] + "</a>");
                    } else {
                        var fieldValue = workItem.fields[fieldName];
                        if (fieldValue.hasOwnProperty("displayName"))
                        {
                            fields.push(fieldValue.displayName);
                        } else {
                            fields.push(fieldValue);
                        }
                    }
                }
            }

            workItemTableStuff.push(fields);
        }

        let html: string = "Query: <a href=\"" + orgUrl + "/web/qr.aspx?pguid=" + projectId + "&qid=" + queryId + "\">" + query!.path! + "</a><br /><br />";
        html += getHTMLTable(workItemTableStuff);

        var sendMethod = tl.getInput("sendMethod");
        switch(sendMethod)
        {
            case "SendGrid": {
                sendMailUsingSendGrid(emailAddresses, html);
                break;
            }

            case "SMTP": {
                tl.debug("Using SMTP as Transport");
                var transporter = getSMTPTransportOptions();
                var smtpFrom = getSMTPFrom();

                var subject = tl.getInput("subject", true);
                
                // setup email data with unicode symbols
                let mailOptions = {
                    from:  smtpFrom, // sender address
                    to: emailAddresses.join(", "),
                    subject: subject, // Subject line
                    html: html // html body
                };

                tl.debug("From: " + mailOptions.from);
                tl.debug("To: " + mailOptions.to);
                tl.debug("Subject: " + subject);
                tl.debug("HTML: " + html);
            
                // send mail with defined transport object
                let info = await transporter.sendMail(mailOptions, function(error, info) {
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
    catch (err) {
        tl.setResult(tl.TaskResult.Failed, err.message);
    }
}

run();