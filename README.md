# Send-M365ApiRequest
Send requests to Microsoft 365 Cloud APIs. The script currently supports Graph API and Office 365 Management API. You don't need anything like libraries or SDKs to run a simple test and either confirm if an app permission is working or if a request is being returned as expected. It is necessary to have only the minimum requirement: an app registration in Azure AD with the appropriate permission to request the respective Api.

### .SYNOPSIS
Send requests to Microsoft 365 Cloud APIs. The script currently supports Graph API and Office 365 Management API. You don't need anything like libraries or SDKs
to run a simple test and either confirm if an app permission is working or if a request is being returned as expected. It is necessary to have only the minimum
requirement: an app registration in Azure AD with the appropriate permission to request the respective Api.

###  .DESCRIPTION
This script will make a request to the API informed in the appropriate parameter and return the results. You just need to fullfil all parameters
as expected by the API documentation to receive the response. Obviously, it is expected that you have a application registration in a Azure AD tenant already created.
You can use both client credentials flow (Application Permission Type) and authorization code flow (Delegated Permission Type).

### .PARAMETER ClientID
The client id attribute from Application Registration in Azure AD.

### .PARAMETER TenantID
The tenant id attibute from Application Registratoin in Azure AD.

### .PARAMETER ClientSecret
The client secret from Applition Registration in Azure AD.

### .PARAMETER Method
HTTP method to be used in the request.

### .PARAMETER Path
Path that will be appended in the Graph API root URI (https://graph.microsoft.com/v1.0)

### .PARAMETER Json
Optional json file to be send as HTTP body in the request.

### .PARAMETER Api
The API to request. Supported APIs are Graph, GraphBeta and ManagementApi

### .PARAMETER Operation
Operation that will be used when contacting ManagementApi. Supported operations are Start, Stop, List and Content. The script doesn't support the usage of
Start operation with a webhook.

### .PARAMETER ContentType
The content type to be returned when using ManagementApi. Supported content types are Audit.AzureActiveDirectory, Audit.Exchange, Audit.SharePoint, Audit.General and Dlp.All.

### .INPUTS
It is possible to use the Json parameter to send inputs to the Api, for instance, when creating a user or sending an email.

### .OUTPUTS
There are requests that is expected to return a body. Please follow the Api documentation to check if the request you're making is supposed to return a response.

### .EXAMPLE
The following example will get all users from Graph 1.0 endpoint in a tenant using client credentials grant flow:

Send-M365ApiRequest.ps1 -ClientID XYZ -TenantID XYZ -ClientSecret XYZ -Api Graph -Method Get -Path /users

### .EXAMPLE
The example below will start a subscription for Exchange Audit logs and after list all log entries in each blob:

Send-M365ApiRequest.ps1 -ClientID XYZ -TenantID XYZ -ClientSecret XYZ -Api ManagementApi -Method Post -Operation Start -ContentType Audit.Exchange
Send-M365ApiRequest.ps1 -ClientID XYZ -TenantID XYZ -ClientSecret XYZ -Api ManagementApi -Method Get -Operation Content -ContentType Audit.Exchange

### .EXAMPLE
This example shows how to list currently subscriptions in Management Api:

Send-M365ApiRequest.ps1 -ClientID XYZ -TenantID XYZ -ClientSecret XYZ -Api ManagementApi -Method Get -Operation List

### .LINK
https://github.com/deividfoggi/Send-M365ApiRequest
#>