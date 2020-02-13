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
    const emailAddresses :string = tl.getInput('emailAddresses', true)!;
    const splittedAddresses = emailAddresses.split(/[\s,\n]+/); // Split on space, comma and newline
    
    tl.debug("Splitted Addresses: " + splittedAddresses);
    return splittedAddresses;
}

function validateEmailAddresses(emailAddresses: string[]) : boolean
{
    for (let emailAddressIndex in emailAddresses)
    {
        const emailAddress = emailAddresses[emailAddressIndex];

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

function getSMTPTransportOptions() : nodemailer.Transporter
{
    const smtpEndpoint = tl.getInput("connectedServiceNameSMTP", true)!;
    const scheme = tl.getEndpointAuthorizationScheme(smtpEndpoint, false);
    
    let auth: any;
    let smtpServer: string;
    let smtpPort: number;
    let tlsOptions: string;

    switch(scheme)
    {
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

function getSMTPFrom() : string
{
    const smtpEndpoint = tl.getInput("connectedServiceNameSMTP", true)!;
    const scheme = tl.getEndpointAuthorizationScheme(smtpEndpoint, false)!;
    
    let smtpFromEmail: string;
    let smtpFromName: string;

    switch(scheme)
    {
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

async function getQueryResult(connection: azdev.WebApi, projectId: string) : Promise<witif.WorkItemQueryResult>
{
    const queryId = getQueryId();   
    const wit: witapi.IWorkItemTrackingApi = await connection.getWorkItemTrackingApi();
    return wit.queryById(queryId);
}

function getQueryId() : string
{
    const queryType: string = tl.getInput('queryType')!;
    let queryId: string;

    switch (queryType)
    {
        case "My": {
            queryId = tl.getInput('queryMy', true)!;
            break;
        }

        case "Shared": {
            queryId = tl.getInput('query', true)!;
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
    const queryType: string = tl.getInput('queryType')!;

    switch (queryType)
    {
        case "My": {
            const patEndpoint = tl.getInput("connectedServiceNameAzureDevOps", true)!;
            return tl.getEndpointUrl(patEndpoint, false)!;
        }

        case "Shared": {
            return tl.getEndpointUrl('SystemVssConnection', false)!;
        }

        default: {
            tl.setResult(tl.TaskResult.Failed, "Query Type \"" + queryType + "\" is invalid");
            throw new Error("Query Type \"" + queryType + "\" is invalid");
        }
    }
}

function getADOConnection() : azdev.WebApi
{
    const queryType: string = tl.getInput('queryType')!;

    switch (queryType)
    {
        case "My": {
            const patEndpoint = tl.getInput("connectedServiceNameAzureDevOps", true)!;
            const scheme = tl.getEndpointAuthorizationScheme(patEndpoint, false)!;
            const orgUrl = getOrgUrl();

            switch (scheme)
            {
                case "Token":
                {
                    const pat = tl.getEndpointAuthorizationParameter(patEndpoint, "apitoken", false)!;
                    const authHandler = azdev.getPersonalAccessTokenHandler(pat)!;
                    return new azdev.WebApi(orgUrl, authHandler);
                }

                case "UsernamePassword":
                {
                    const username = tl.getEndpointAuthorizationParameter(patEndpoint, "username", false)!;
                    const password = tl.getEndpointAuthorizationParameter(patEndpoint, "password", false)!;
                    const authHandler = azdev.getBasicHandler(username, password)!;
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
            const token = tl.getEndpointAuthorizationParameter('SystemVssConnection', 'AccessToken', false)!;
            const orgUrl = getOrgUrl();
            
            const authHandler = azdev.getBearerHandler(token)!;
            return new azdev.WebApi(orgUrl, authHandler);    
        }

        default: {
            tl.setResult(tl.TaskResult.Failed, "Query Type \"" + queryType + "\" is invalid");
            throw new Error("Query Type \"" + queryType + "\" is invalid");
        }
    }
}

function convertToBoolean(input: string): boolean | undefined {
    try {
        return JSON.parse(input.toLowerCase());
    }
    catch (e) {
        return undefined;
    }
}

async function run() {
    try {
        const projectId: string = tl.getInput('project', true)!;
        const sendOnEmpty: string = tl.getInput('sendIfEmpty', true)!;

        const sendOnEmptyBool = convertToBoolean(sendOnEmpty);

        const queryId = getQueryId();
        const orgUrl = getOrgUrl();
        
        const emailAddresses = getEmailAddresses();
        if (!validateEmailAddresses(emailAddresses))
        {
            return;
        }

        const connection = getADOConnection();
        const result = await getQueryResult(connection, projectId);
        
        const wit: witapi.IWorkItemTrackingApi = await connection.getWorkItemTrackingApi();
        const query = await wit.getQuery(projectId, queryId);
        
        tl.debug(JSON.stringify(result));

        const ids: number[] = [];

        if (result == null || result.workItems == undefined)
        {
            return;
        }

        for (let item of result.workItems as Array<WorkItemReference>)
        {
            ids.push(Number(item.id));
        }

        const columnNames: string[] = [];
        if (result.columns != null)
        {
            for (let fieldIndex in result.columns)
            {
                if (result.columns[fieldIndex].referenceName == undefined)
                {
                    continue;
                }

                columnNames.push(result.columns[fieldIndex].referenceName!);
            }
        }

        const workItemTableStuff : string[][] = []
        workItemTableStuff.push(columnNames);

        let workItems : witif.WorkItem[] = [];

        if (ids.length == 0)
        {
            if (!sendOnEmptyBool)
            {
                tl.setResult(tl.TaskResult.Succeeded, 'Empty Query. Not sending E-Mail.');
                return;
            }
        }

        if (ids.length > 0)
        {
            while(ids.length)
            {
                const chunk = ids.splice(0,200);
                const batchRequest = { fields: columnNames, ids: chunk };
                workItems.push(...await wit.getWorkItemsBatch(batchRequest));
            }
        }

        workItems = workItems.filter((a) => a.fields != undefined);

        if (result.sortColumns != null && result.sortColumns.length > 0)
        {
            const sortColumn = result.sortColumns[0]!;
            
            const field = sortColumn.field!;
            const referenceName = field.referenceName!;
            tl.debug("Ordering by " + referenceName + (sortColumn.descending! ? " descending" : " ascending"));

            workItems.sort((a, b) => (a.fields![referenceName].localeCompare(b.fields![referenceName])));
            
            if (sortColumn.descending)
            {
                workItems = workItems.reverse();
            }
        }

        for (let wi in workItems)
        {
            const workItem = workItems[wi];
            if (workItem.fields == undefined)
            {
                continue;
            }

            const fields: string[] = [];
            for (let field in columnNames)
            {
                tl.debug(Object.keys(workItem.fields).toString());

                const fieldName = columnNames[field];
                if (!workItem.fields[fieldName])
                {
                    fields.push("");
                } else {
                    if (fieldName == "System.Id")
                    {
                        fields.push("<a href=\"" + orgUrl + "/" + projectId + "/_workItems/edit/" + workItem.fields[fieldName] + "\">" + workItem.fields[fieldName] + "</a>");
                    } else {
                        const fieldValue = workItem.fields[fieldName];
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

        const sendMethod = tl.getInput("sendMethod");
        switch(sendMethod)
        {
            case "SendGrid": {
                sendMailUsingSendGrid(emailAddresses, html);
                break;
            }

            case "SMTP": {
                tl.debug("Using SMTP as Transport");
                const transporter = getSMTPTransportOptions();
                const smtpFrom = getSMTPFrom();

                const subject = tl.getInput("subject", true);
                
                // setup email data with unicode symbols
                const mailOptions = {
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
                const info = transporter.sendMail(mailOptions, function(error, info) {
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