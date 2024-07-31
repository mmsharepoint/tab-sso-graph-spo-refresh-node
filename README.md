# Tab SSO Graph SPO Refresh

## Summary

This app describes how to establish a Teams Tab as SSO solution including the generation of a second SharePoint access token.
Although Microsoft Graph is the way to go there are still gaps to the original SharePoint rest API. One example is shown in this sample as it is the /_api/web/ensureuser endpoint

Teams Tab 

![Teams Tab](assets/)

For further details see the author's [blog post](https://mmsharepoint.wordpress.com/2024/)

## Applies to

This sample was created [using the Teams Toolkit with Visual Studio Code](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/teams-toolkit-fundamentals?pivots=visual-studio&WT.mc_id=M365-MVP-5004617). The same sample will also realized with the [Teams Toolkit for Visual Studio 2022](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/toolkit-v4/teams-toolkit-fundamentals-vs?WT.mc_id=M365-MVP-5004617) and ...

## Version history

Version|Date|Author|Comments
-------|----|----|--------
1.0|Aug 08, 2024|[Markus Moeller](https://twitter.com/moeller2_0)|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## Minimal Path to Awesome
- Clone the repository
    ```bash






**Congratulations**! You are running an application that can now show a beautiful web page in Teams, Outlook and the Microsoft 365 app.

![Personal tab demo](https://github.com/OfficeDev/TeamsFx/assets/63089166/9599b53c-8f89-493f-9f7e-9edae1f9be54)

## What's included in the template

| Folder       | Contents                                                                                                               |
| ------------ | ---------------------------------------------------------------------------------------------------------------------- |
| `.vscode`    | VSCode files for debugging                                                                                             |
| `appPackage` | Templates for the Teams application manifest                                                                           |
| `env`        | Environment files                                                                                                      |
| `infra`      | Templates for provisioning Azure resources                                                                             |
| `src`        | The source code for the frontend of the Tab application. Implemented with Fluent UI Framework.                         |
| `api`        | The source code for the backend of the Tab application. Implemented single-sign-on with OBO flow using Azure Functions. |

The following are Teams Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Teams Toolkit works.

| File                 | Contents                                                                                                                                                                                                                                                |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `teamsapp.yml`       | This is the main Teams Toolkit project file. The project file defines two primary things: Properties and configuration Stage definitions.                                                                                                               |
| `teamsapp.local.yml` | This overrides `teamsapp.yml` with actions that enable local execution and debugging.                                                                                                                                                                   |
| `aad.manifest.json`  | This file defines the configuration of Microsoft Entra app. This template will only provision [single tenant](https://learn.microsoft.com/azure/active-directory/develop/single-and-multi-tenant-apps#who-can-sign-in-to-your-app) Microsoft Entra app. |

## Extend the React with Fluent UI template

Following documentation will help you to extend the React with Fluent UI template.

- [Add or manage the environment](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Create multi-capability app](https://learn.microsoft.com/microsoftteams/platform/toolkit/add-capability)
- [Use an existing Microsoft Entra application](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-existing-aad-app)
- [Customize the Teams app manifest](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- [Collaborate on app development](https://learn.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Set up the CI/CD pipeline](https://learn.microsoft.com/microsoftteams/platform/toolkit/use-cicd-template)
- [Publish the app to your organization or the Microsoft Teams app store](https://learn.microsoft.com/microsoftteams/platform/toolkit/publish)
- [Enable the app for multi-tenant](https://github.com/OfficeDev/TeamsFx/wiki/Multi-tenancy-Support-for-Azure-AD-app)
- [Preview the app on mobile clients](https://github.com/OfficeDev/TeamsFx/wiki/Run-and-debug-your-Teams-application-on-iOS-or-Android-client)
