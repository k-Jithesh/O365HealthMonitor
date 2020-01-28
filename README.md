# O365 Health Monitor
 
This sample Azure Function App, monitores for O365 service degradation and notifies to a Teams Channel.

Application queries O365 service management API.
https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference 
writes the json to Azure blob. PowerBI Report can be used to visualise the information. 

## Setup Instructions.

1. Create a new azure app registration, and create client secrets.
	1. Copy Domain Name, Client Id, Client Secret from the created app registration.
2. Grant ServiceHealth.Read permission to the newly created applicaiton.
![Api Permission](https://user-images.githubusercontent.com/20592381/72151738-05656780-33cf-11ea-8282-c596e9e8a632.png)
3. Create a new Storage account (general purpose v1) and copy the connection string (from under Access Keys).
4. Create a Teams webhook to write notifications to Teams
https://docs.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook
5. Download the source code, build and publish the project to create a new Azure function app/or use existing function app.
6. On the Newly created function app, provide/create the following application settings
	1. ClientId
	2. ClientSecret
	3. Domain
	4. TeamsWebhookURL
	5. StorageConnectionString
	6. Env

![image](https://user-images.githubusercontent.com/20592381/72153099-ea94f200-33d2-11ea-9699-e2822f6288a6.png)

##### use Dev as Env application settings value only for execution from Visual Studio. 

7. Open O365HealthMonitor.pbix and modify the blob storage location strings.
![image](https://user-images.githubusercontent.com/20592381/72153638-88d58780-33d4-11ea-97a0-89e335d848e4.png)


## -- Powerbi Reports --
![O365Health](https://user-images.githubusercontent.com/20592381/72154448-a73c8280-33d6-11ea-9e03-fe2d47a51d29.jpg)

## -- Teams Notification --
![TeamsNotification](https://user-images.githubusercontent.com/20592381/72154462-b15e8100-33d6-11ea-8e30-639323679542.jpg) 
