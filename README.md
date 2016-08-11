# Microsoft Graph Excel REST API ASP.NET Excel as a Calculation Service sample

This sample shows how to call an Excel built-in function (PMT) from within an Excel document stored in your OneDrive for Business account using the Excel REST APIs.

## Prerequisites

To use the Microsoft Graph Excel REST API ASP.NET Excel as a Calculation Service, you need the following:
* Visual Studio 2015 installed and working on your development computer. 

     > Note: This sample is written using Visual Studio 2015. If you're using Visual Studio 2013, make sure to change the compiler language version to 5 in the Web.config file:  **compilerOptions="/langversion:5**
* A Microsoft Office 365 account. You can sign up for [an Office 365 Developer subscription](https://aka.ms/devprogramsignup) that includes the resources that you need to start building apps.

     > Note: If you already have a subscription, the previous link sends you to a page with the message *Sorry, you can’t add that to your current account*. In that case, use an account from your current Office 365 subscription.
* A Microsoft Azure Tenant to register your application. Azure Active Directory (AD) provides identity services that applications use for authentication and authorization. A trial subscription can be acquired here: [Microsoft Azure](https://account.windowsazure.com/SignUp).

     > Important: You also need to make sure your Azure subscription is bound to your Office 365 tenant. To do this, see the Active Directory team's blog post, [Creating and Managing Multiple Windows Azure Active Directories](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx). The section **Adding a new directory** will explain how to do this. You can also see [Set up your Office 365 development environment](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription) and the section **Associate your Office 365 account with Azure AD to create and manage apps** for more information.
* The client ID and redirect URI values of an application registered in Azure. This sample application must be granted the **Have full access to user files and files shared with user** permission for **Microsoft Graph**. [Add a web application in Azure](https://msdn.microsoft.com/office/office365/HowTo/add-common-consent-manually#bk_RegisterWebApp) and grant the proper permissions to it:
	* In the [Azure Management Portal](https://manage.windowsazure.com/), select the **Active Directory** tab and an Office 365 tenant.
	* Select the **Applications** tab and choose the application that you want to configure.
	* In the **permissions to other applications** section, add the **Microsoft Graph** application.
	* For the **Microsoft Graph** application, add the following delegated permissions: **Have full access to user files, read and write access to user profile and sign in to user profile**.
	* Save the changes.

     > Note: During the app registration process, make sure to specify **http://localhost:44347** as the **Sign-on URL** and the **Reply URL**.  

## Configure the app
1. Open **ExcelAsACalculationService.sln** file. 
2. In Solution Explorer, open the **Web.config** file. 
3. Replace *ENTER_YOUR_CLIENT_ID* with the client ID of your registered Azure application.
4. Replace *ENTER_YOUR_SECRET* with the key of your registered Azure application.

## Run the app
1. Press F5 to build and debug. Run the solution and sign in with your organizational account. The application launches on your local host and shows the starter page. 
     > Note: Copy and paste the start page URL address **http://localhost:44347/home/index** to a different browser if you get the following error during sign in: **AADSTS70001: Application with identifier ad533dcf-ccad-469a-abed-acd1c8cc0d7d was not found in the directory**.
2. Select the `Excel` link from the top menu bar.
4. The application relies on you having an Excel workbook called 'Book.xlsx' in the root OneDrive folder of your O365 account. If this file does not exist, please manually add to your OneDrive by navigating to **https://yourtenant.sharepoint.com**, clicking on the App Launcher "Waffle" at the top left of the page, and then choosing the OneDrive application - add a new Excel file called Book.xlsx.
5. Enter the 3 values rate, nper, and pv and press the calculate button - the payment per month should be calculated. 


