# Azure DevOps Schedule Work Item Query

## Introduction

The Goal of this extension was to keep as much functionality as possible inside Azure DevOps / Team Foundation Server.
We have opted to realize this by using a standard Build Pipeline in Azure DevOps / Team Foundation Server as our "Scheduling" Engine.

The actual functionality is realized as a Build Pipeline Task called "Scheduled Work Item Query".
This task executes a query that is saved in either the "My Queries" or "Shared Queries" folders, and sends the results by e-mail using either [SendGrid](https://www.sendgrid.com) or standard SMTP.

## Setup

### Introduction

Setup of the Task should be straight forward. You should probably create a new Build Pipeline, and then add as many "Scheduled Work Item Query" Tasks as needed.
Each Task is able to send a single Query to multiple people.
The Build Pipeline does not have to live in the same Project as the Query.

You should then add a Schedule Trigger to the Build Pipline, that it gets executed automatically whenever you need it to.

### Endpoints

#### SendGrid Endpoint

The SendGrid endpoint can be used if you want to send e-mails using [SendGrid](https://www.sendgrid.com).
All you need is a SendGrid API Token with the permissions to send e-mails.
For Documentation about this, see [here](https://sendgrid.com/docs/ui/account-and-settings/api-keys/)

#### SMTP Endpoint

The SMTP endpoint can be used to send e-mails using standard SMTP interfaces.
It supports common TLS configurations as well as authentication.

#### Azure Repos / Team Foundation Server Endpoint

This Endpoint is needed when you want to send a query that exists only in someone's "My Queries" folder.
It is not needed if you want to send out queries living in the "Shared Queries" folder.

### Task

#### Shared Queries

Shared queries can be scheduled without any other authentication, it uses the authentication/authorization of the Agent to connect to Azure DevOps Service / Azure DevOps Server.

#### Personal (My) Queries

"My Queries" requires a personal access token or username/password from the person which owns the query. You need to configure an Azure DevOps Endpoint for this to work.
Note: We recommend generating a Personal Access Token with minimal rights (e.g. only work_item.read scope), and to limit the created Endpoint to a specific pipeline.
Make sure you don't expose the endpoint for others to use, as this might be a security risk.