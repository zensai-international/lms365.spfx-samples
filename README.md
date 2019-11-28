## lms365-spfx-samples

The repository contains SharePoint Framework solution demonstrates how to use LMS365 API using modern SharePoint web parts.

### Before you begin

* Please checkout [SharePoint framework documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview) and **make sure** you [setup dev environment](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment) for SPFx solutions.  
* Clone this repository

### Configure access to LMS365 API
Access to LMS365 API should be configured using `package-solution.json` 
* You need to open `https://portal.azure.com > Azure Active Directory > Enterprise Applications` and search lms365 api application. it could be found by one of two search criterias: `lms365-api-prod` or `LMS365 API`
![](https://i.imgur.com/oJSDhMr.png)  
OR  
![](https://i.imgur.com/IHQz8P2.png)

* Once application found copy the name `lms365-api-prod` or `LMS365 API`

* Open `repository-folder/config/package-solution.json`
and make sure that section `webApiPermissionRequests` has right name of resource from previous step
```json
"webApiPermissionRequests": [
    {
        "resource": "LMS365 API", //or "lms365-api-prod"
        "scope": "user_impersonation"
    }
]
```

### Build solution in debug mode

```bash
cd repository-folder
npm i
gulp bundle
gulp package-solution
gulp serve
```

### Deploy solution in debug mode

* Upload the package file from `repository-folder\sharepoint\solution\lms365-spfx-samples.sppkg` to the SharePoint app catalog
* Once package is uploaded you need to approve LMS365 API permission request. It could be done using steps from [documentation](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient#manage-permission-requests)  
![](https://i.imgur.com/XAmL2J9.png)  

the request is approved one time. After request is approved web part will able to connect to the LMS365 API and ready to run.
* Open any modern site you have admin access to and then navigate to hosted workbench `https://contoso.sharepoint.com/sites/modernsite/_layouts/15/workbench.aspx`  
**`IMPORTANT: workbench should be hosted and NOT localhost based`**
* On the page you can add `lms365-statistics` web part it should look like that  
IMAGE
