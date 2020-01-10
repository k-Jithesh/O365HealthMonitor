# O365HealthMonitor
 
This sample Azure Function App, monitores for O365 service degradation and notifies to a Teams Channel.

Application queries O365 service management API.
https://docs.microsoft.com/en-us/office/office-365-management-api/office-365-service-communications-api-reference 
writes the json to Azure blob. PowerBI Report can be used to visualise the information. 
