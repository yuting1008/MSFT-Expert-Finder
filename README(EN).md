# Expert Finder with SSO in Copilot as a Teams AI-based Message Extension

> *This English verion of README was translated by AI and may contain inaccuracies. If needed, please refer to the Chinese version for the most accurate and official content.*

This is a Microsoft 365 Copilot integrated Teams application that uses [Microsoft Graph](https://developer.microsoft.com/en-us/graph) to search for experts based on keywords such as skills and location. This application supports Single Sign-On (SSO), providing a better user experience and authentication functionality.

1. Users can invoke the Expert Finder agent in Microsoft 365 Copilot and search for experts using natural language, such as: "Find AI experts for me," and the system will return a list of matching experts.
2. Users can also launch the Expert Finder application in the [Teams Message Extension](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-message-extension-bot) and input keywords like Azure or AI to retrieve a list of experts that match the criteria.

<img src="images/demo.gif" alt="img-alt-text" width="700">


### Table of Contents

- [Step 1. Prerequisites](#step-1.-prerequisites)  
- [Step 2. Prepare the Data](#step-2.-prepare-the-data)  
    - [Step 2.1 Update User Profile Photos and Office Locations via Microsoft Entra Admin Center](#step-2.1-update-user-photo-and-office-location-via-microsoft-entra-admin-center)  
    - [Step 2.2 Update User Skills via API in Graph Explorer](#step-2.2-update-user-skills-via-api-in-graph-explorer)  
- [Step 3. Set Up Local Development Environment](#step-3.-set-up-local-development-environment)  
- [Step 4. Deploy the App to Azure](#step-4.-deploy-the-app-to-azure)  
- [Step 5. Publish the App to Teams](#step-5.-publish-the-app-to-teams)  
- [Step 6. Use the App in Teams and Microsoft 365 Copilot](#step-6.-use-the-app-in-teams-and-microsoft-365-copilot)  
    - [Step 6.1 Search by Skill and Office Location in Copilot](#step-6.1-search-by-skill-and-office-location-in-copilot)  
    - [Step 6.2 Use Message Extension in Teams Chat](#step-6.2-use-message-extension-in-teams-chat)  
- [Step 7. Bulk Update User Information via PowerShell (Optional)](#step-7.-bulk-update-user-information-via-powershell-(optional))  
  - [Step 7.1 Clone the Repository](#step-7.1-clone-the-repository)  
  - [Step 7.2 Prepare a CSV File with User Info to be Updated](#step-7.2-prepare-a-csv-file-with-user-info-to-be-updated)  
  - [Step 7.3 Install Microsoft Graph PowerShell SDK](#step-7.3-install-microsoft-graph-powershell-sdk)  
  - [Step 7.4 Create a Self-signed Public Certificate for App Authentication](#step-7.4-create-a-self-signed-public-certificate-to-authenticate-the-app)  
  - [Step 7.5 Register the Microsoft Entra App](#step-7.5-register-a-microsoft-entra-application)  
  - [Step 7.6 Grant API Permissions to the App by Global Admin](#step-7.6-grant-api-permissions-to-the-app-by-global-admin)  
  - [Step 7.7 Bulk Update User Info Using PowerShell](#step-7.7-bulk-update-user-info-using-powershell)

---

# Step 1. Prerequisites

Before testing the app locally and deploying it to Azure, please ensure that the required applications and packages are installed, and that you have accounts with the necessary permissions.

1. **An Azure account** that meets the following requirements:
   - Has [Owner](https://learn.microsoft.com/en-us/azure/role-based-access-control/built-in-roles/privileged#owner) or [Contributor](https://learn.microsoft.com/en-us/azure/role-based-access-control/built-in-roles/privileged#contributor) role permissions.
   - Has a valid **Subscription**, which will be used to create all Azure resources required later. If you don’t have a subscription, refer to [Create a Microsoft Customer Agreement subscription](https://learn.microsoft.com/en-us/azure/cost-management-billing/manage/create-subscription) to set one up.

2. **A Microsoft 365 account** that meets the following requirements:
   - Has [Application Administrator](https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/permissions-reference#application-administrator) and [Teams Administrator](https://learn.microsoft.com/en-us/microsoftteams/using-admin-roles#teams-roles-and-capabilities) permissions.
   - Has [permission to upload custom Teams apps](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant#enable-custom-teams-apps-and-turn-on-custom-app-uploading).
   - Has a **Copilot for Microsoft 365 license**. Without this license, you will not be able to use the app within Copilot, but you can still use it as a Message Extension in Teams chat.

3. [Node.js 18.x](https://nodejs.org/en/download)  
4. [Chrome](https://support.google.com/chrome/answer/95346?hl=en&co=GENIE.Platform%3DDesktop)  
5. [Visual Studio Code](https://code.visualstudio.com/)  
6. [Teams Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension)  
    1. After installing Teams Toolkit, you will see its icon in the sidebar of VS Code. Please sign in with your Azure account and Microsoft 365 account, and ensure that you have a **Copilot for Microsoft 365 license** and **permission to upload custom apps**.  
        ![account-login](images/login.png =600x)
    1. If the **upload custom Teams apps** permission is disabled, follow these steps to enable it:

        - When custom app upload is disabled, the following screen will appear:  
            ![custom-app-disabled](images/custom-app-disabled.png =400x)
        - Go to the [Teams Admin Center](https://admin.teams.microsoft.com/).
        - Navigate to **Teams apps** > **Setup policies** > **Global (Org-wide default)**.  
            ![teams-app-upload-permission](images/setup-poilcies.png =500x)
        - Enable **Upload custom apps**.  
            ![teams-app-upload-permission](images/teams-app-upload-permission-2.png =500x)
        - Then go to **Teams apps** > **Manage apps** > **Actions** > **Org-wide app settings**.  
            ![teams-app-upload-permission](images/manage-app.png =1000x)
        - Enable **Upload custom apps for personal use**.  
            ![teams-app-upload-permission](images/org-wide-app-settings.png =300x)
        - Once all steps are completed, return to VS Code. The Teams Toolkit will show that **Custom App Upload is Enabled**.  
            ![teams-app-upload-permission](images/custom-app-allowed.png =300x)


# Step 2. Prepare the Data

When a user enters a query, Expert Finder calls the [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview) to access user data and returns information such as name, profile photo, office location, and skills. This section explains how to update a user's photo and office location via the Microsoft Entra Admin Center, and how to update user skills using the Graph Explorer API.  
If you would like to update user information in bulk, please refer to [Step 7. Bulk Update User Information via PowerShell (Optional)](#step-7-bulk-update-user-information-via-powershell-optional).

## Step 2.1 Update User Photo and Office Location via Microsoft Entra Admin Center

1. Go to the [Microsoft Entra Admin Center](https://entra.microsoft.com) and sign in with your Microsoft 365 account.
2. Navigate to **Users** > **All Users**.  
   ![img-alt-text](images/entra-id-users.png =600x)
3. Select the user whose information you want to update and update their profile photo in the **Overview** section.  
   ![img-alt-text](images/photo.png =600x)
4. Click on **Properties**.  
   ![img-alt-text](images/property.png =600x)
5. Scroll down to **Job Information** and click the pencil icon to edit.  
   ![img-alt-text](images/job-info.png =600x)
6. Update the **Office location** field and save.  
   ![img-alt-text](images/office-location.png =600x)

## Step 2.2 Update User Skills via API in Graph Explorer

### 2.2.1 First-Time Setup: Grant Permissions to Graph Explorer in Microsoft Entra Admin Center

1. Open [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and sign in with your Microsoft 365 account. Your Microsoft Entra tenant will now log Graph Explorer as an identity.
2. Go to the [Microsoft Entra Admin Center](https://entra.microsoft.com) and sign in with your Microsoft 365 account to authorize Graph Explorer.
3. Navigate to **Applications**. If this is your first time using Graph Explorer in the Entra Admin Center, you will need to enable **Global Secure Access**.  
   ![img-alt-text](images/enable-application-access.png =800x)
4. In **Enterprise Applications**, locate and select **Graph Explorer**.  
   ![img-alt-text](images/entra-id-application.png =600x)
5. Go to the **Permissions** tab to grant organizational consent to Graph Explorer.  
   ![img-alt-text](images/enable-graph.png =500x)

### 2.2.2 Get User ID from Microsoft Entra Admin Center and Update Skills in Graph Explorer

1. Open the [Microsoft Entra Admin Center](https://entra.microsoft.com) and sign in with your Microsoft 365 account.
2. Navigate to **Users** > **All Users** and find the user whose data you want to update.
3. Copy the user’s **Object ID** — you’ll use this in Graph Explorer shortly.  
   ![img-alt-text](images/user-id.png =800x)
4. Open [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) and sign in.
5. Use the `PATCH` method to call:  
   ```
   https://graph.microsoft.com/v1.0/users/{user-id}
   ```  
   In the Request Body, input the skills in the following format:  
   ```json
   {
     "skills": ["Python", "JavaScript"]
   }
   ```  
   or  
   ```json
   {
     "skills": ["量化交易"]
   }
   ```
6. If successful, you’ll see an HTTP status code **204 No Content** as shown below.  
   ![img-alt-text](images/graph-explorer.jpg =800x)
7. Use the `GET` method to call:  
   ```
   https://graph.microsoft.com/v1.0/users/{user-id}/?$select=displayName,skills
   ```  
   to verify that the skills were successfully updated.  
   ![img-alt-text](images/skill-check.png =800x)

> You’ll need to repeat the [steps above](#get-user-id-from-microsoft-entra-admin-center-and-update-skills-in-graph-explorer) each time you want to update user skills.

> For more Graph API usage, refer to the [Microsoft Graph REST API v1.0 endpoint reference](https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0&preserve-view=true).


# Step 3. Set Up Local Development Environment

Before deploying and publishing the application, we’ll first set up the local environment and test the app in the Teams web client.

1. Clone the repository:

    ```bash
    git clone https://mtc-taiwan@dev.azure.com/mtc-taiwan/MTC%20Internal/_git/expert-finder
    ```

2. Navigate to the `expert-finder` directory and open it with Visual Studio Code.
3. In VS Code, go to **File > Open Folder**, and select the project folder.
4. In the left sidebar, go to **Run and Debug** > **Debug in Teams (Chrome)** to run the app locally in the Teams web version.  
   ![debug-in-teams](images/debug.png =400x)
5. A browser will automatically open and navigate to Teams web. Keep this browser tab open.
6. Go back to the `env/env.local` file in the project, and copy the `AAD_APP_CLIENT_ID` value to Notepad — you will use it shortly to locate the application.  
   ![img-alt-text](images/aad-client-id.png =600x)
7. Go to the [Microsoft Entra Admin Center](https://entra.microsoft.com) and sign in with your Microsoft 365 account. You will now grant the required Graph API permissions.
8. In **Applications**, open **App registrations**, and find the application using the copied `AAD_APP_CLIENT_ID` as the **Application (client) ID**.  
   ![img-alt-text](images/app-registration.png =600x)
9. Open the **API permissions** tab and grant **admin consent** on behalf of your organization.  
   ![img-alt-text](images/api-permission.png =700x)  
   Once done, the status will change to a green checkmark, indicating that consent has been granted.  
   ![img-alt-text](images/permission-check.png =700x)
10. Return to the browser with Teams web opened and test the app following the instructions in [this section](#step-6-use-the-app-in-teams-and-microsoft-365-copilot).
11. After testing is complete, you can close the browser.

> When launching the debugger, the following sequence is executed. For a detailed explanation, refer to [Debug (F5) in Visual Studio Code](https://github.com/OfficeDev/teams-toolkit/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#debug-f5-in-visual-studio-code):
>
> 1. The `launch.json` file acts as the entry point when starting the debugger in VS Code. It defines the `preLaunchTask` and the target URL for the browser.
> 2. Based on `launch.json`, VS Code executes tasks defined in `tasks.json`, including creating the `env` folder/files, reading environment variables, validating prerequisites, and executing the Teams app lifecycle.
> 3. It runs the app lifecycle defined in `teamsapp.local.yml`, including [Provision](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/provision), [Deploy](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy), and [Publish](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy-teams-app-to-container-service). During this process, it references the environment variables in `env/.env.local` and `env/.env.local.user`.
>
> ![f5-tasks](images/f5-tasks.png =500x)



# Step 4. Deploy the App to Azure

Make sure you have successfully run the app in your local development environment as described in the [previous step](#step-3-set-up-local-development-environment). In this step, we will deploy the application to Azure and create necessary Azure resources such as Bot Services and App Services.

The process uses the Teams app lifecycle defined in the `teamsapp.yml` file, including [Provision](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/provision), [Deploy](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy), and [Publish](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy-teams-app-to-container-service). It also references environment variables from `env/.env.dev` and `env/.env.dev.user`.

1. Open **Teams Toolkit**. We’ll be using the **Lifecycle** section as shown below:  
   ![teams-toolkit-lifecycle](images/provision-deploy.png =400x)
2. In the Lifecycle section, select **Provision**. This will create the necessary resources in Azure. First, select the Azure **subscription** in which the resources will be created.  
   ![subscription](images/subscription.png =400x)
3. Select the Azure **resource group** where the resources will be created.  
   ![resource-group](images/resource-group.png =400x)
4. Click **Provision** to complete the Azure resource creation.  
   ![provision](images/provision.png =400x)
5. Next, in the Lifecycle section, select **Deploy** to deploy the application to Azure.  
   <p><span style="color:red;font-weight:bold">⚠ Note: If you make changes to the code in the future, you can re-click **Deploy** to update the app in Azure.</span></p>

6. Go to the `env/env.dev` file in your VS Code project folder, and copy the `AAD_APP_CLIENT_ID` to Notepad — you’ll need it shortly to locate the application.  
   ![img-alt-text](images/aad-client-id-2.png =600x)
7. Open the [Microsoft Entra Admin Center](https://entra.microsoft.com) and sign in with your Microsoft 365 account. We’ll now grant the necessary Graph API permissions.
8. In **Applications**, open **App registrations**, and find the application using the `AAD_APP_CLIENT_ID` copied earlier as the **Application (client) ID**.  
   ![img-alt-text](images/app-registration2.png =600x)
9. Go to **API permissions** and grant **admin consent** on behalf of your organization.  
   ![img-alt-text](images/api-permission2.png =700x)  
   Once complete, the status will show a green checkmark, indicating that consent has been granted.  
   ![img-alt-text](images/permission-check.png =700x)
10. After completing this step, you will see the created resources in the Azure resource group, including Azure Bot, App Service, and App Service Plan.  
    ![img-alt-text](images/resource.png =600x)


# Step 5. Publish the App to Teams

After deploying the app to Azure, we will upload it to the Teams Admin Center and allow it to be published to your organization’s Teams app store.

1. Open **Teams Toolkit**, and in the **Lifecycle** section, select **Publish** to publish the app to the Teams Admin Center.  
   ![teams-toolkit-lifecycle](images/publish.png =400x)
2. Go to the [**Teams Admin Center**](https://admin.teams.microsoft.com/) and sign in with your Microsoft 365 account.
3. Navigate to **Manage apps** > search for `Expert Finder` > select the newly published app.  
   ![img-alt-text](images/teams-admin-center-1.png =800x)
4. Approve the publication of the `Expert Finder dev` app.  
   ![img-alt-text](images/teams-admin-center-2.png =600x)
5. Open your Teams client, go to the Teams **App Store**, and install the app.  
   You can now refer to the [next section](#step-6-use-the-app-in-teams-and-microsoft-365-copilot) to start using it.  
   ![install-app](images/app-store.png =700x)


# Step 6. Use the App in Teams and Microsoft 365 Copilot

The following shows how to use **Expert Finder** in both Copilot and Teams chat.

## Step 6.1 Search by Skill and Office Location in Copilot

Go to the Microsoft 365 Copilot chat interface. In the top-right corner of the chat, you’ll find the **Expert Finder agent**. Click it to begin using the app.

Here are some sample prompts:

1. `Find experts with skill in Azure.`
2. `Find experts with skill in Python and who are from Taipei.`

![demo](images/copilot-demo.gif =800x)

## Step 6.2 Use Message Extension in Teams Chat

Open a Teams chat, click the "+" icon in the bottom-right corner, and select **Expert Finder**. In the search bar, enter a skill keyword such as “Azure” to search for experts.

![demo](images/teams-message-extension-demo.gif =800x)


# Step 7. Bulk Update User Information via PowerShell (Optional)

This section guides you through using a PowerShell script to automate bulk updates to users’ personal information, including **skills** and **office location**. It covers preparing user data, installing required packages, generating a certificate, registering a Microsoft Entra application, granting API permissions, and executing the update.

<p><span style="color:red;font-weight:bold">⚠ If you need to maintain or sync a large number of user records, we recommend using this method to reduce manual operations in Microsoft Entra and Graph Explorer.</span></p>

## Step 7.1 Clone the Repository

Run the command below to clone the repository. If you already cloned it earlier, you may skip this step:

```bash
git clone https://mtc-taiwan@dev.azure.com/mtc-taiwan/MTC%20Internal/_git/expert-finder
```


## Step 7.2 Prepare a CSV File with User Info to Be Updated

Before using PowerShell to update user data, prepare a CSV file containing all the user info to be updated. You can use Microsoft Excel to edit the file.

1. Open the `UserSkills.csv` file located in the `PowerShell` folder of the project repository.
2. Delete the sample entries and fill in your own user data as follows:

| Column Name        | Required | Description                                                                 |
|--------------------|----------|-----------------------------------------------------------------------------|
| `UserPrincipalName`| Yes      | The user’s email address (e.g., `user1@contoso.com`).                       |
| `Skills`           | No       | List of user skills. Use **commas ( , )** to separate multiple values (e.g., `Data Science, Python, Azure`). |
| `OfficeLocation`   | No       | Office location of the user (e.g., `Taipei`).                              |

3. Save the file without changing its location or name.

<p><span style="color:red;font-weight:bold">⚠ Make sure the CSV file is saved in the `PowerShell` folder and named `UserSkills.csv`, so that relative paths work correctly in later steps.</span></p>



## Step 7.3 Install Microsoft Graph PowerShell SDK

To interact with Microsoft Graph via PowerShell, install the [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview?view=graph-powershell-1.0):

1. In the Start menu, search for **PowerShell**, right-click, and choose **Run as administrator**.  
   ![demo](images/open-powershell.png =500x)
2. Run the command below to install the SDK:
    ```
    Install-Module Microsoft.Graph -Scope CurrentUser
    ```
3. Verify the installation with:
    ```
    Get-Module Microsoft.Graph -ListAvailable
    ```
4. If installed successfully, you will see the module's version and path:  
   ![img](images/graph-module-intall.png =600x)



## Step 7.4 Create a Self-signed Public Certificate to Authenticate the App

When using Microsoft Graph PowerShell SDK, the app must authenticate using a certificate. This step generates a self-signed certificate and exports it as a `.cer` file for later registration.

1. Open PowerShell as administrator.
2. Navigate to the `PowerShell` folder in your project. Replace `YourFolderPath` with your actual path.

    ```powershell
    cd "YourFolderPath"
    ```

3. Generate the certificate:

    ```powershell
    $certname = "Certificate"
    $cert = New-SelfSignedCertificate -Subject "CN=$certname" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256
    Export-Certificate -Cert $cert -FilePath "./$certname.cer"
    ```

4. You should now see `Certificate.cer` in the `PowerShell` folder.

> For more information, refer to [this guide](https://learn.microsoft.com/en-us/entra/identity-platform/howto-create-self-signed-certificate).



## Step 7.5 Register a Microsoft Entra Application

To securely access Microsoft Graph API, register an app in Microsoft Entra ID. This step uses the `RegisterAppOnly.ps1` script to automate registration, upload the certificate, and save the app info in `AppInfo.json`.

1. Open PowerShell as administrator.
2. Navigate to the `PowerShell` folder. Replace `YourFolderPath` with your actual path:

    ```powershell
    cd "YourFolderPath"
    ```

3. Run the registration script:

    ```powershell
    .\RegisterAppOnly.ps1 -AppName "Graph PowerShell Script" -CertPath "Certificate.cer"
    ```

4. Upon success, you’ll see `AppInfo.json` in the folder containing your `ClientId` and `TenantId`. You’ll also see the app in [Microsoft Entra Admin Center](https://entra.microsoft.com), as shown in [Step 7.6](#step-76-grant-api-permissions-to-the-app-by-global-admin).

> Learn more from [this official doc](https://learn.microsoft.com/en-us/powershell/microsoftgraph/app-only?view=graph-powershell-1.0).



## Step 7.6 Grant API Permissions to the App by Global Admin

After registering the app, you need to grant Microsoft Graph API permissions.

<p><span style="color:red;font-weight:bold">⚠ Only a Microsoft 365 **Global Administrator** account can complete this step.</span></p>

1. Go to the [Microsoft Entra Admin Center](https://entra.microsoft.com) and sign in as a Global Admin.
2. Navigate to **Applications** > **App Registrations**, and find your app `Graph PowerShell Script`.
3. Go to **API Permissions**.  
   ![img](images/app.png =600x)
4. Grant the following permissions:
   - `Sites.ReadWrite.All`
   - `User.Read.All`
   - `User.ReadWrite.All`  
   ![img](images/grant-api-permission.png =600x)

> See more details in [this guide](https://learn.microsoft.com/zh-tw/entra/identity/role-based-access-control/permissions-reference).



## Step 7.7 Bulk Update User Info Using PowerShell

Now that your app is registered and permissions granted, you can run a script to update users' skills and office locations from the `UserSkills.csv` file.

1. Open PowerShell as administrator.
2. Navigate to the `PowerShell` folder. Replace `YourFolderPath` with your actual path.:

    ```powershell
    cd "YourFolderPath"
    ```

3. Run the update script:

    ```powershell
    .\UpdateUserSkills.ps1
    ```

4. To verify if updates were successful, run the following script. Replace `YourUserPrincipalName` with your actual user email:

    ```powershell
    $appInfo = Get-Content "./AppInfo.json" | ConvertFrom-Json
    Connect-MgGraph -ClientId $appInfo.ClientId -TenantId $appInfo.TenantId -CertificateName "CN=Certificate"
    Get-MgUser -UserId "YourUserPrincipalName" -select "Skills", "OfficeLocation"  | Select Skills, OfficeLocation
    Disconnect-MgGraph
    ```

<p><span style="color:red;font-weight:bold">
⚠ To update user info again in the future, you only need to follow  
<a href="#step-72-prepare-a-csv-file-with-user-info-to-be-updated">Step 7.2</a> to prepare the CSV file, and  
<a href="#step-77-bulk-update-user-info-using-powershell">Step 7.7</a> to run the script — no need to re-register the app or re-grant permissions.
</span></p>
