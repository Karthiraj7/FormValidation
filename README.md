#spfx

## Summary

Overview
1. Project Components
ContentTab: Manages the display and navigation of various content tabs within the application.
EditContainer: Handles the editing of form data, including form validation.
ViewContainer: Displays the content in a view-only mode and manages approval processes.
Form Validation using PnP (Patterns and Practices)
PnP.js Overview
PnP.js is a reusable library that simplifies SharePoint REST API calls. It can be used to interact with SharePoint lists, libraries, and other SharePoint entities.

Form Validation Steps
Form Design: Create the form structure within the EditContainer component using React.

Input Handling: Use React state to manage form input values and changes.

Validation Rules: Implement validation rules (e.g., required fields, data formats) using functions or a validation library.

PnP.js Integration:

Use PnP.js to fetch necessary data (like dropdown options) and submit form data to the SharePoint list.
Example: import { sp } from "@pnp/sp";
Error Handling: Display validation errors dynamically as users interact with the form.

Approval Levels Using SharePoint List
Approval Process Overview
An approval process typically involves multiple stages where each stage requires a designated person or group to approve or reject the item.

Steps to Implement Approval Levels
SharePoint List Setup:

Create a SharePoint list to store form data and approval status.
Add columns such as Status, Approver, Comments, and Approval Level.
EditContainer:

On form submission, save the form data to the SharePoint list with an initial status (e.g., "Pending").
Use PnP.js to submit and update list items.
ViewContainer:

Display the form data along with current approval status.
Provide an interface for approvers to approve/reject and add comments.
Update the status in the SharePoint list based on approver's action using PnP.js.
ContentTab:

Navigate between different stages of content or different approval levels.
Display relevant data and actions based on the user's role (e.g., viewer, editor, approver).
Approval Workflow:

Define the logic to move items through various approval levels.
Update the item in the SharePoint list after each approval stage.
Send notifications to next approver using Power Automate or custom code if necessary.
Summary
In this SPFx React project, the ContentTab, EditContainer, and ViewContainer components are used to create a seamless experience for form data entry, validation, and approval processing. By leveraging PnP.js, the application can efficiently interact with SharePoint lists to manage form data and approval workflows. Form validation ensures data integrity, while the approval levels implemented via SharePoint lists enable structured and trackable approval processes.








## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.18.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Any special pre-requisites?

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| folder name | Author details (name, company, twitter alias with link) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

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
