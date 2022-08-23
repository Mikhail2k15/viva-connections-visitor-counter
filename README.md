# Visitor Counter for Viva Connections Dashboard

## Summary

Visitor Counter for Viva Connections Dashboard is a simple widget (Adaptive Card Extension) which provides to you insights about how users interact with your intranet. For example, you can view the number of people who have visited your intranet portal via a brand-new Viva Connections App in Teams vs browser (old way).

![image](https://user-images.githubusercontent.com/11201670/178820760-b4d50cb8-9649-4fc7-9d2c-1b6ae1de44e2.png)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.15.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

> Application Insights

## Solution

| Solution            | Author(s)                |
| ------------------- | ------------------------ |
| Michael Bondarevsky | bondarevsky at gmail.com |

## Version history

| Version | Date            | Comments                                           |
| ------- | --------------- | -------------------------------------------------- |
| 1.1.1   | August 22, 2022 | Upgrade to 1.15.2 sfpx, fixes some eslint warnings |
| 1.0.4   | March 26, 2022  | Fixes                                              |
| 1.0     | March 24, 2022  | Initial release                                    |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- In the command-line run:
  - **npm install**
  - **gulp serve --nobrowser**
  - **gulp bundle --ship**
  - **gulp package-solution --ship**

> Log in to the Azure Portal,

- Create an Application Insights resource
- In the sidebar, navigate to Configure > API Access on the sidebar
- Create API key with a read telemertry permission
- Copy the Application ID and API Key
