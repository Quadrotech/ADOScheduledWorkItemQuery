import tl = require('azure-pipelines-task-lib/task');
import * as azdev from "azure-devops-node-api";
import * as witif from "azure-devops-node-api/interfaces/WorkItemTrackingInterfaces";
import * as witapi from "azure-devops-node-api/WorkItemTrackingApi"
import { WorkItemReference } from 'azure-devops-node-api/interfaces/TestInterfaces';
import { sendQueryResult, getEmailAddresses, validateEmailAddresses } from './email';
import { convertToBoolean } from './convertToBoolean';
import { sortWorkItems } from './sortWorkItems';

"use strict";

function getHTMLTable(data: string[][]) : string {
    let result = ['<table border=1>'];
    for(let row of data) {
        result.push('<tr>');
        for(let cell of row){
            result.push(`<td>${cell}</td>`);
        }
        result.push('</tr>');
    }
    result.push('</table>');
    return result.join('\n')!;
}

function getHTMLTableHierarchy(threads: any[][], workItems: Array<{id: number, workItem: witif.WorkItem}>, columns: string[]) : string {
    let result = ['<table border=1>'];

    // Main Header
    result.push('<tr>');
    for (let cell of columns) {
        result.push(`<td>${cell}</td>`);
    }
    result.push('</tr>');

    tl.debug(`Work Items:`);
    tl.debug(JSON.stringify(workItems));

    // Items
    for (let i = 0; i < threads.length; i++) {
        let thread = threads[i][0];
        let threadId = threads[i][1];
        tl.debug(`Finding ${threadId}`);
        let workItem = workItems.find(wi => wi.id == threadId)!.workItem;

        // NORMAL ITEM
        result.push('<tr>');
        for(let cell of columns){
            result.push(`<td>${getFieldValue(workItem, cell)}</td>`);
        }
        result.push('</tr>');

        // SUBITEMS
        for (let y = 0; y < thread.length; y++) {
            result.push('<tr>');
            // SUBTABLE
            tl.debug(`Finding Sub: ${thread[y]}`);
            let subItem = workItems.find(wi => wi.id == thread[y])!.workItem;
            for(let cell of columns){
                result.push(`<td>&nbsp;&nbsp;&nbsp;&nbsp;${getFieldValue(subItem, cell)}</td>`);
            }
            result.push('</tr>');
        }
    }

    result.push('</table>');
    return result.join('\n')!;
}

function getProjectId() : string
{
    return tl.getInput('project', true)!;
}

async function getQueryResult(connection: azdev.WebApi) : Promise<witif.WorkItemQueryResult>
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

async function run() {
    try {
        const orgUrl = getOrgUrl();
        const projectId = getProjectId();
        const queryId = getQueryId();
       
        const emailAddresses = getEmailAddresses();
        if (!validateEmailAddresses(emailAddresses))
        {
            return;
        }

        const connection = getADOConnection();
        const result = await getQueryResult(connection);
        
        const wit: witapi.IWorkItemTrackingApi = await connection.getWorkItemTrackingApi();
        const query = await wit.getQuery(projectId, queryId);

        tl.debug(JSON.stringify(result));

        switch(result.queryResultType)
        {
            case witif.QueryResultType.WorkItem:
                tl.debug("QueryResultType.WorkItem")
                const workItemTableStuff = await getFlatItemWorkItemTable(result, wit);
                if (workItemTableStuff == undefined)
                {
                    return;
                }
                
                let html: string = `Query: <a href="${orgUrl}/web/qr.aspx?pguid=${projectId}&qid=${queryId}">${query!.path!}</a><br /><br />`;
                html += getHTMLTable(workItemTableStuff);

                sendQueryResult(html, emailAddresses);
                break;
            case witif.QueryResultType.WorkItemLink:
                tl.debug("QueryResultType.WorkItemLink");
                const tree = await getTree(result, query, wit);
                const workItems = await getWorkItemsFromTree(result, wit, tree);

                if (workItems == undefined) {
                    return;
                }

                const columnNames = getColumnNames(result);

                let html2: string = `Query: <a href="${orgUrl}/web/qr.aspx?pguid=${projectId}&qid=${queryId}">${query!.path!}</a><br /><br />`;
                html2 += getHTMLTableHierarchy(tree, workItems, columnNames);

                sendQueryResult(html2, emailAddresses);
                break;
            default:
                tl.debug(witif.QueryResultType[result.queryResultType!]);
                return tl.setResult(tl.TaskResult.Failed, "queryResultType " + witif.QueryResultType[result.queryResultType!] + " is not supported. (default)");
                break;
        }
    } catch (err) {
        tl.setResult(tl.TaskResult.Failed, err.message);
    }
}

async function getWorkItemsFromTree(result : witif.WorkItemQueryResult, wit: witapi.IWorkItemTrackingApi, tree : any[][]) : Promise<Array<{id: number, workItem: witif.WorkItem}> | undefined>
{
    let workItemStuff: Array<{id: number, workItem: witif.WorkItem}> = [];

    let ids: number[] = []
    for (let i = 0; i < tree.length; i++) {
        let thread = tree[i][0];
        ids.push(tree[i][1]);
        for (let y = 0; y < thread.length; y++) {
            ids.push(thread[y]);
        }
    }
    if (ids.length == 0) {
        const sendOnEmpty: string = tl.getInput('sendIfEmpty', true)!;
        const sendOnEmptyBool = convertToBoolean(sendOnEmpty);
        if (!sendOnEmptyBool) {
            tl.setResult(tl.TaskResult.Succeeded, 'Empty Query. Not sending E-Mail.');
            return;
        }
    }

    let uniqueIds = [...new Set(ids)];
    const columnNames = getColumnNames(result);

    if (uniqueIds.length > 0)
    {
        while(uniqueIds.length)
        {
            const chunk = uniqueIds.splice(0,200);
            const batchRequest = { fields: columnNames, ids: chunk };
            const items = await wit.getWorkItemsBatch(batchRequest);

            for(let item of items as witif.WorkItem[])
            {
                workItemStuff.push({id: item.id!, workItem: item});
            }
        }
    }

    return workItemStuff;
}


async function getFlatItemWorkItemTable(result: witif.WorkItemQueryResult, wit: witapi.IWorkItemTrackingApi) : Promise<string[][] | undefined>
{
    const ids: number[] = [];

    if (result == null || result.workItems == undefined)
    {
        return;
    }

    for (let item of result.workItems as Array<WorkItemReference>)
    {
        ids.push(Number(item.id));
    }

    const columnNames: string[] = getColumnNames(result);
    const workItemTableStuff : string[][] = []
    workItemTableStuff.push(columnNames);

    let workItems : witif.WorkItem[] = [];

    if (ids.length == 0)
    {
        const sendOnEmpty: string = tl.getInput('sendIfEmpty', true)!;
        const sendOnEmptyBool = convertToBoolean(sendOnEmpty);

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
    workItems = sortWorkItems(result, workItems);

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
            fields.push(getFieldValue(workItem, fieldName));
        }

        workItemTableStuff.push(fields);
    }

    return workItemTableStuff;
}

function getFieldValue(workItem: witif.WorkItem, fieldName: string) : string
{
    if (workItem == undefined)
    {
        tl.debug("undefined");
    }

    if (workItem.fields === undefined)
    {
        tl.debug('UNDEFINED');
        tl.debug(JSON.stringify(workItem));
    }

    tl.debug(`Field name: ${fieldName}`);
    if (!workItem.fields![fieldName]) {
        tl.debug(`Returning empty`);
        return "";
    } else {
        if (fieldName == "System.Id") {
            tl.debug(`System.Id field!!`);
            const orgUrl = getOrgUrl();
            const projectId = getProjectId();
            return `<a href="${orgUrl}${projectId}/_workItems/edit/${workItem.fields![fieldName]}">${workItem.fields![fieldName]}</a>`;
        } else {
            tl.debug(`else`);
            const fieldValue = workItem.fields![fieldName];
            if (fieldValue.hasOwnProperty("displayName")) {
                return fieldValue.displayName;
            } else {
                return fieldValue;
            }
        }
    }
}

function getColumnNames(result: witif.WorkItemQueryResult) : string[]
{
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

    return columnNames;
}

async function getTree(result: witif.WorkItemQueryResult, query: witif.QueryHierarchyItem, wit: witapi.IWorkItemTrackingApi) : Promise<any[][]>
{
    let threads: any[][] = [];
    let thread: number[] = [];
    for (let item of result.workItemRelations as Array<witif.WorkItemLink>)
    {
        if (item.rel == null) {
            // Start new "Thread"
            thread = [];
            threads.push([thread, item.target?.id!]);
        } else {
            thread.push(item.target?.id!);
        }
    }

    for (let i = 0; i < threads.length; i++) {
        let thread = threads[i][0];
        tl.debug("Thread: " + threads[i][1]);
        for (let y = 0; y < thread.length; y++) {
            tl.debug("Content: " + thread[y]);
        }
    }

    return threads;
}

run();