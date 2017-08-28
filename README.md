# ibera-document-service

## Prerequisites
Turn the Windows Identity Foundation Features on: 
Open Windows Features and turn `Windows Identity Foundation 3.5` on

Install the [Microsoft.Identity.Extensions](http://go.microsoft.com/fwlink/?LinkID=252368)

## Configuration
Before running the solution, ensure you set `StorageAccountConnectionString` and     `OutlookIntegrationHostUri` (e.g. https://localhost:8443/Functions.html)

When running local you can set it in your [web.config](https://github.com/CatalystCode/ibera-document-service/blob/master/DocumentService/Web.config) or in AppSettings when deploying to an Azure WebApp.
