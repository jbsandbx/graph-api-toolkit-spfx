# Build a SharePoint web part with the Microsoft Graph Toolkit

This topic covers how to use Microsoft Graph Toolkit components in a SharePoint client-side web part. Getting started involves the following steps:

## Set up your development environment and create a web part.
Add the Microsoft Graph Toolkit SharePoint Framework package.
Add the SharePoint Provider.
Add components.
Configure permissions.
Deploy the Microsoft Graph Toolkit SharePoint Framework package.
Build and deploy your web part.
Test your web part.

## Set up your SharePoint Framework development environment and create a new web part
Follow the steps to Set up your SharePoint Framework development environment and then create a new web part.

## Add the Microsoft Graph Toolkit SharePoint Framework package
To prevent multiple SharePoint Framework components from registering their own set of Microsoft Graph Toolkit components on the page, you should deploy the Microsoft Graph Toolkit SharePoint Framework package to your tenant and reference Microsoft Graph Toolkit components that you use in your solution from this package.

The Microsoft Graph Toolkit SharePoint Framework package contains a SharePoint Framework library that registers a single instance of Microsoft Graph Toolkit components in SharePoint.

Install the Microsoft Graph Toolkit SharePoint Framework npm package using the following command:

npm install @microsoft/mgt-spfx

## Add the SharePoint Provider
The Microsoft Graph Toolkit providers enable authentication and access to Microsoft Graph for the components. To learn more, see Using the providers. SharePoint web parts always exist in an authenticated context because the user has already had to sign in in order to get to the page that hosts your web part. Use this context to initialize the SharePoint provider.

First, add the provider to your web part. Locate the src\webparts\<your-project>\<your-web-part>.ts file in your project folder, and add the following line to the top of your file, right below the existing import statements:

import { Providers, SharePointProvider } from '@microsoft/mgt-spfx';
Next, you need to initialize the provider with the authenticated context inside the onInit() method of your web part. In the same file, add the following code right before the public render(): void { line:

protected async onInit() {
  if (!Providers.globalProvider) {
    Providers.globalProvider = new SharePointProvider(this.context);
  }
}

## Add components
Now, you can start adding components to your web part. Simply add the components to the HTML inside of the render() method, and the components will use the SharePoint context to access Microsoft Graph. For example, to add the Person component, your code will look like:

public render(): void {
    this.domElement.innerHTML = `
      <mgt-person person-query="me" view="twolines"></mgt-person>
    `;
}

### Note

If you're building a web part using React, see the @microsoft/mgt-spfx docs to learn how to use @microsoft/mgt-react.

## Configure permissions
To call Microsoft Graph from your SharePoint Framework application, you need to request the needed permissions in your solution package and a Microsoft 365 tenant administrator needs to approve the requested permissions.

To add the permissions to your solution package, locate and open the config\package-solution.json file and set:

"isDomainIsolated": false,
Just below that line, add the following:

"webApiPermissionRequests":[]

Determine which Microsoft Graph API permissions you need depending on the components you're using. Each component's documentation page provides a list of the permissions that component requires. You will need to add each permission required to webApiPermissionRequests. For example, if you're using the Person component and the Agenda component, your webApiPermissionRequests might look like:

"webApiPermissionRequests": [
  {
    "resource": "Microsoft Graph",
    "scope": "User.Read"
  },
  {
    "resource": "Microsoft Graph",
    "scope": "Calendars.Read"
  }
]

## Deploy the Microsoft Graph Toolkit SharePoint Framework package
Before deploying your SharePoint Framework package to your tenant, you will need to deploy the Microsoft Graph Toolkit SharePoint Framework package to your tenant. You can download the package corresponding to the version of Microsoft Graph Toolkit that you used in your project, from the Releases section on GitHub.

 ### Important
Because only one version of the SharePoint Framework library for Microsoft Graph Toolkit can be installed in the tenant, before you use the Microsoft Graph Toolkit in your solution, determine whether your organization or customer already has a version of the SharePoint Framework library deployed and use the same version.

After downloading the Microsoft Graph Toolkit SharePoint Framework .sppkg package, upload it to your SharePoint Online App Catalog. Go to the More features page of your SharePoint admin center. Select Open under Apps, then click App Catalog, and Distribute apps for SharePoint. Upload your .sppkg file, and click Deploy.

## Build and deploy your web part
Now, you will build your application and deploy it to SharePoint. Build your application by running the following commands:

gulp build
gulp bundle
gulp package-solution
In the sharepoint/solution folder, there will be a new .sppkg file. You will need to upload this file to your SharePoint Online App Catalog. Go to the More features page of your SharePoint admin center. Select Open under Apps, then click App Catalog, and Distribute apps for SharePoint. Upload your .sppkg file, and click Deploy.
Next, you need to approve the permissions as an administrator.
Go to your SharePoint Admin center. In the left-hand navigation, select Advanced and then API Access. You should see pending requests for each of the permissions you added in your config\package-solution.json file. Select and approve each permission.

## Test your web part
You're now ready to add your web part to a SharePoint page and test it out. You will need to use the hosted workbench to test web parts that use the Microsoft Graph Toolkit because the components need the authenticated context in order to call Microsoft Graph. You can find your hosted workbench at https://<YOUR_TENANT>.sharepoint.com/_layouts/15/workbench.aspx.

Open the config\serve.json file in your project and replace the value of initialPage with the url for your hosted workbench:
"initialPage": "https://<YOUR_TENANT>.sharepoint.com/_layouts/15/workbench.aspx",
Save the file and then run the following command in the console to build and preview your web part:

gulp serve
Your hosted workbench will automatically open in your browser. Add your web part to the page and you should see your web part with the Microsoft Graph Toolkit components in action! As long as the gulp serve command is still running in your console, you can continue to make edits to your code and then just refresh your browser to see your changes.


## Please refer the below documentation
(https://docs.microsoft.com/en-us/graph/toolkit/get-started/build-a-sharepoint-web-part#deploy-the-microsoft-graph-toolkit-sharepoint-framework-package)

<hr/>

# graph-tookit-sp-fx

## Summary

Short summary on functionality and used technologies.

[picture of the solution in action, if possible]

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.13-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

Solution|Author(s)
--------|---------
folder name | Author details (name, company, twitter alias with link)

## Version history

Version|Date|Comments
-------|----|--------
1.1|March 10, 2021|Update comment
1.0|January 29, 2021|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development