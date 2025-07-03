# Expert Finder with SSO in Copilot as a Teams AI-based Message Extension


### 目錄

- [Step 1. 先決條件](#step-1.-先決條件)
- [Step 2. 資料準備](#step-2.-資料準備)  
    - [Step 2.1 利用 Microsoft Entra Admin Center 更新用戶的照片與辦公室地點](#step-2.1-利用-microsoft-entra-admin-center-更新用戶的照片與辦公室地點)  
    - [Step 2.2 在 Graph Explorer 中利用 API 更新用戶的技能](#step-2.2-在-graph-explorer-中利用-api-更新用戶的技能)  
- [Step 3. 本地開發環境建置應用程式](#step-3.-本地開發環境建置應用程式)  
- [Step 4. 部署應用程式至 Azure](#step-4.-部署應用程式至-azure)
- [Step 5. 發佈應用程式至 Teams](#step-5.-發佈應用程式至-teams)  
- [Step 6. 在 Teams 與 Microsoft 365 Copilot 使用應用程式](#step-6.-在-teams-與-microsoft-365-copilot-使用應用程式)
    - [Step 6.1 在 Copilot 中依技能與辦公室地點搜尋](#step-6.1-在-copilot-中依技能與辦公室地點搜尋)
    - [Step 6.2 在 Teams 聊天室中使用 Message Extension](#step-6.2-在-teams-聊天室中使用-message-extension)
- [Step 7. 利用 PowerShell 大量更新用戶資訊 (Optional)](#step-7.-利用-powershell-大量更新用戶資訊-(optional))
  - [Step 7.1 複製儲存庫](#step-7.1-複製儲存庫)
  - [Step 7.2 準備待更新的用戶資訊 CSV 檔](#step-7.2-準備待更新的用戶資訊-csv-檔)
  - [Step 7.3 安裝 Microsoft Graph PowerShell SDK](#step-7.3-安裝-microsoft-graph-powershell-sdk)
  - [Step 7.4 建立 Self-signed Public Certificate 以驗證應用程式](#step-7.4-建立-self-signed-public-certificate-以驗證應用程式)
  - [Step 7.5 註冊 Microsoft Entra 應用程式](#step-7.5-註冊-microsoft-entra-應用程式)
  - [Step 7.6 全域管理員授予應用程式 API 權限](#step-7.6-全域管理員授予應用程式-api-權限)
  - [Step 7.7 利用 PowerShell 大量更新用戶資訊](#step-7.7-利用-powershell-大量更新用戶資訊)


這一個結合 Microsoft 365 Copilot 的 Teams 應用程式，透過 [Microsoft Graph](https://developer.microsoft.com/en-us/graph) 根據技能、地點等關鍵字搜尋專家。此應用程式支援單一登入 (SSO)，提供更佳的使用者體驗與身份驗證功能。
1. 使用者可在 Microsoft 365 Copilot 中開啟 Expert Finder 代理程式以自然語言查詢專家，例如：「請幫我尋找 AI 專家」，系統就會回傳符合條件的專家名單。
1. 使用者亦可在 [Teams Message Extension](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-message-extension-bot) 中開啟 Expert Finder 應用程式，並輸入技能如 Azure、AI 等關鍵字，系統就會回傳符合條件的專家名單。


![img-alt-text](images/demo.gif =700x)


# Step 1. 先決條件
在開始本地測試應用程式與部署應用程式至 Azure 前，請先確保已安裝以下提及的應用程式與套件，並具備帶有指定權限的帳號。

1. 滿足以下條件的 Azure 帳號。
    - 具備[負責人 ( Owner )](https://learn.microsoft.com/zh-tw/azure/role-based-access-control/built-in-roles/privileged#owner)或[參與者 ( Contributor )](https://learn.microsoft.com/zh-tw/azure/role-based-access-control/built-in-roles/privileged#contributor) 權限。
    - 具備訂用帳戶 ( Subscription )，後續會使用到的所有 Azure 資源將建立於此訂用帳戶中，若您尚未建立訂用帳戶，請參閱 [建立 Microsoft 客戶合約訂用帳戶](https://learn.microsoft.com/zh-tw/azure/cost-management-billing/manage/create-subscription) 以建立訂用帳戶。
1. 滿足以下條件的 Microsoft 365 帳號。
    - 具備[應用程式系統管理員 ( Application Administrator )](https://learn.microsoft.com/zh-tw/entra/identity/role-based-access-control/permissions-reference#application-administrator)和[Teams 系統管理員 ( Teams administrator )](https://learn.microsoft.com/zh-tw/microsoftteams/using-admin-roles#teams-roles-and-capabilities) 權限。
    - 具備[上傳自訂 Teams 應用程式權限](https://learn.microsoft.com/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant#enable-custom-teams-apps-and-turn-on-custom-app-uploading)。
    - 具備 Copilot for Microsoft 365 授權。若您不具備  Copilot for Microsoft 365 授權，您將無法在 Copilot 中使用此應用程式，但仍可以作為 Message Extension 在 Teams 聊天室中使用。
1. [Node.js 18.x](https://nodejs.org/en/download)
1. [Chrome](https://support.google.com/chrome/answer/95346?hl=zh-Hant&co=GENIE.Platform%3DDesktop#zippy=)
1. [Visual Studio Code](https://code.visualstudio.com/)
1. [Teams Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension)
    1. Teams Toolkit 安裝完成後你會在 VS Code 左側欄中看到圖示，請使用您的 Azure 帳戶和 Microsoft 365 帳戶登入，並確認具備 Copilot for Microsoft 365 授權及上傳自訂應用程式的權限。 \
        ![account-login](images/login.png =600x)
    1. 如果上傳自訂 Teams 應用程式權限被禁用，請依照以下步驟啟用上傳自訂應用程式的權限：

        - 當上傳自訂應用程式被禁用時，會顯示以下畫面： \
            ![custom-app-disabled](images/custom-app-disabled.png =400x)
        - 前往 [Teams 管理中心](https://admin.teams.microsoft.com/)。
        - 導覽至 Teams 應用程式 ( Teams apps ) > 設定原則 ( Setup policies ) > 全域 Global（組織範圍的應用程式預設值 Org-wide default ）。 \
            ![teams-app-upload-permission](images/setup-poilcies.png =500x)
        - 啟用「上傳自訂應用程式」 ( Upload custom apps )。 \
            ![teams-app-upload-permission](images/teams-app-upload-permission-2.png =500x)
        - 前往 Teams 應用程式 ( Teams apps ) > 管理應用程式 ( Manage apps ) > 操作 ( Actions ) > 組織範圍的應用程式設定 ( Org-wide app settings )。 \
            ![teams-app-upload-permission](images/manage-app.png =1000x)
        - 啟用「允許上傳自訂應用程式供個人使用」( Upload custom apps for personal use )。 \
            ![teams-app-upload-permission](images/org-wide-app-settings.png =300x)
        - 完成以上步驟後回到 VS Code，您的 Teams Toolkit 中會顯示 Custom App Uploaded Enabled。 \
            ![teams-app-upload-permission](images/custom-app-allowed.png =300x)

# Step 2. 資料準備
當使用者輸入查詢時，Expert Finder 會呼叫 [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview) 以存取用戶資料，並回傳用戶的姓名、照片、辦公室地點、以及技能。以下將介紹如何利用 Microsoft Entra Admin Center 更新用戶的照片與辦公室地點，以及在 Graph Explorer 中利用 API 更新用戶的技能。若您想要批次更改大量使用者資訊，您可以參考 [Step 7. 利用 PowerShell 大量更新使用者資訊 (Optional)](#step-7.-利用-powershell-大量更新使用者資訊-(optional))。
<!-- <p> <span style="color:red;font-weight:bold">⚠ 注意：您必須擁有 <a href="https://learn.microsoft.com/zh-tw/entra/identity/role-based-access-control/permissions-reference#global-administrator">Microsoft Entra 全域系統管理員</a> 的權限。</span></p> -->

## Step 2.1 利用 Microsoft Entra Admin Center 更新用戶的照片與辦公室地點
1. 開啟 [Microsoft Entra Admin Center](https://entra.microsoft.com) 並以您的 Microsoft 365 帳號登入。
1. 瀏覽使用者 ( Users ) > 所有使用者 ( All Users )。 \
    ![img-alt-text](images/entra-id-users.png =600x)
1. 選擇待更新資訊的用戶，並在概觀 ( Overview ) 中更新用戶照片。 \
    ![img-alt-text](images/photo.png =600x)
1. 點選屬性 ( Properties ) 。 \
    ![img-alt-text](images/property.png =600x)
1. 請下滑至工作資訊 ( Job Information )，並選擇鉛筆符號進行編輯。 \
    ![img-alt-text](images/job-info.png =600x)
1. 更新辦公室位置 ( Office location ) 的資料並儲存。 \
    ![img-alt-text](images/office-location.png =600x)

## Step 2.2 在 Graph Explorer 中利用 API 更新用戶的技能
### 2.2.1 首次使用 Graph Explorer 時需先至 Microsoft Entra Admin Center 授予權限
1.	開啟 [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) 並以您的 Microsoft 365 帳號登入，此時您的 Microsoft Entra 租用戶會作為身分識別，在 Microsoft Entra Admin Center 中紀錄 Graph Explorer。
1.	前往 [Microsoft Entra Admin Center](https://entra.microsoft.com) 並以您的 Microsoft 365 帳號登入。接下來將會授權剛才記錄到的 Graph Explorer。
1.  前往應用程式 ( Application )，若您是首次在 Microsoft Entra Admin Center 中開啟應用程式，需要先啟用全域安全存取。 \
    ![img-alt-text](images/enable-application-access.png =800x)
1.	選擇企業應用程式 ( Enterprise applications ) 中的 Graph Explorer。 \
    ![img-alt-text](images/entra-id-application.png =600x)
1.	選擇權限 ( Permission ) 代表您的組織授予 Graph Explorer 權限。 \
    ![img-alt-text](images/enable-graph.png =500x)

### 2.2.2 於 Microsoft Entra Admin Center 取得用戶 ID 並在 Graph Explorer 更新用戶技能
1.	開啟 [Microsoft Entra Admin Center](https://entra.microsoft.com) 以您的 Microsoft 365 帳號登入。
1.  導覽至使用者 ( Users ) > 所有使用者 ( All users ) 中尋找欲更新資料的用戶。
1.  複製其物件識別碼 ( Object ID )，稍後會在 Graph Explorer 中使用此 id 更改用戶資料。 \
    ![img-alt-text](images/user-id.png =800x)
1.	開啟 [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) 並以您的 Microsoft 365 帳號登入。
1.	使用 PACTH 功能呼叫 `https://graph.microsoft.com/v1.0/users/{user-id}`，並在 Request Body 中依照格式輸入欲更新的內容，格式如`{"skills": ["Python", "JavaScript"]}`或`{"skills": ["量化交易"]}`等。
1.  成功執行後會看到 204 No Content 的 HTTP 狀態碼如下圖所示。 \
    ![img-alt-text](images/graph-explorer.jpg =800x)
1.  您可以使用 GET 功能呼叫 ``https://graph.microsoft.com/v1.0/users/{user-id}/?$select=displayName,skills`` 來檢查用戶的技能是否成功更新。 \
    ![img-alt-text](images/skill-check.png =800x)

> 每次要更新用戶的技能時，皆需重複[以上步驟](#於-microsoft-entra-admin-center-取得用戶-id-並在-graph-explorer-更新用戶技能)。

> 你可以參考 [Microsoft Graph REST API v1.0 endpoint reference](https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0&preserve-view=true) 來使用更多 Graph API。

# Step 3. 本地開發環境建置應用程式
在正式部署與發佈應用程式前，我們將建立本地環境並在 Teams 網頁版中測試應用程式。

1. 複製儲存資料庫：

    ```bash
    git clone https://mtc-taiwan@dev.azure.com/mtc-taiwan/MTC%20Internal/_git/expert-finder
    ```
1. 將目錄切換至 `expert-finder`，並使用 Visual Studio Code 開啟。
1. 在 VS Code 中選擇檔案 > 開啟資料夾，並選擇範例專案的資料夾。
1. 選擇執行與偵錯 ( Run and Debug ) > Debug in Teams (Chrome) 以在 Teams 網頁版本地端執行應用程式。 \
    ![debug-in-teams](images/debug.png =400x)
1. 瀏覽器將自動開啟並進入 Teams 網頁版，請保持此瀏覽器開啟。
1. 回到 VS Code 目錄中的 ``env/env.local``，複製 ``AAD_APP_CLIENT_ID``到記事本，稍後將會用於尋找應用程式。 \
    ![img-alt-text](images/aad-client-id.png =600x)
1. 前往 [Microsoft Entra Admin Center](https://entra.microsoft.com) 並以您的 Microsoft 365 帳號登入。接下來要授權 Graph API 所需要的權限。
1. 開啟應用程式 ( Application ) 中的應用程式註冊  ( Application registration )，找到以剛才複製的 ``AAD_APP_CLIENT_ID`` 作為應用程式 ( 用戶端 ) 識別碼的應用程式。 \
    ![img-alt-text](images/app-registration.png =600x)
1. 開啟 API 權限，並代表您的組織授與管理員同意。 \
    ![img-alt-text](images/api-permission.png =700x) \
    完成後你會看到狀態改為綠色打勾符號，並顯示已授與。 \
    ![img-alt-text](images/permission-check.png =700x)
1. 回到開啟 Teams 網頁板的瀏覽器，根據[此章節](#在-teams-與-microsoft-365-copilot-使用應用程式)測試應用程式。
1. 完成測試後即可關閉瀏覽器。

> 啟動偵錯後，實際上會歷經以下流程，詳細說明可參考 [Debug (F5) in Visual Studio Code](https://github.com/OfficeDev/teams-toolkit/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#debug-f5-in-visual-studio-code)
> 1. 以 ``launch.json`` 作為 VSCode 啟動偵錯後的入口位置，定義 preLaunchTask 以及開啟瀏覽器後要導向的網址。
> 2. 根據 ``launch.json``，接續執行 ``tasks.json`` 中定義的任務，包含確認並建立 ``env`` 資料夾與檔案、讀取環境變數、驗證先備條件、執行 Teams 應用程式生命週期等。
> 3. 執行 ``teamsapp.local.yml`` 中定義的 Teams 應用程式的生命週期，包含 [Provision](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/provision)、[Deploy](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy)、[Publish](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy-teams-app-to-container-service)，過程中會參考 ``env/.env.local`` 與 ``env/.env.local.user`` 中的環境變數。
> 
> ![f5-tasks](images/f5-tasks.png =500x)


# Step 4. 部署應用程式至 Azure
請確保已照[上述步驟](#本地測試應用程式)在本地開發環境成功運行，本階段我們將部署應用程式至 Azure，並建立 Azure 資源如 Bot Services 和 App Services。過程中會使用到 ``teamsapp.yml`` 定義的 Teams 應用程式生命週期，包含 [Provision](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/provision)、[Deploy](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy)、[Publish](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy-teams-app-to-container-service)，同時會參考 ``env/.env.dev`` 與 ``env/.env.dev.user`` 中的環境變數。

1. 開啟 Teams Toolkit，接下來會利用到 Lifecycle 的功能如下圖所示。 \
    ![teams-toolkit-lifecycle](images/provision-deploy.png =400x)
1. 於 Lifecycle 中選擇 **Provision**（佈建），此動作將會在 Azure 中建立必要的資源。首先，選擇即將建立資源的 Azure 訂閱 ( Subscription )。 \
    ![subscription](images/subscription.png =400x)
1. 選擇即將建立資源的 Azure 資源群組 ( Resource Group )。 \
    ![resource-group](images/resource-group.png =400x)
1. 按下 **Provision** 完成 Azure 資源建立。 \
    ![provision](images/provision.png =400x)
1. 於 Lifecycle 中選擇 **Deploy**（部署），將應用程式部署至 Azure。
    <p> <span style="color:red;font-weight:bold">⚠ 注意：如果您未來有修改程式碼，您可以重新點擊 Deploy 以將變更部署至 Azure。</span></p>

1. 回到 VS Code 目錄中的 ``env/env.dev``，複製 ``AAD_APP_CLIENT_ID``到記事本，稍後將會用於尋找應用程式。 \
    ![img-alt-text](images/aad-client-id-2.png =600x)
1. 前往 [Microsoft Entra Admin Center](https://entra.microsoft.com) 並以您的 Microsoft 365 帳號登入。接下來要授權 Graph API 所需要的權限。
1. 開啟應用程式 ( Application ) 中的應用程式註冊  ( Application registration )，找到以剛才複製的 ``AAD_APP_CLIENT_ID`` 作為應用程式 ( 用戶端 ) 識別碼的應用程式。 \
    ![img-alt-text](images/app-registration2.png =600x)
1. 開啟 API 權限，並代表您的組織授與管理員同意。 \
    ![img-alt-text](images/api-permission2.png =700x) \
    完成後你會看到狀態改為綠色打勾符號，並顯示已授與。 \
    ![img-alt-text](images/permission-check.png =700x)
1. 完成本階段後，您會在 Azure 資源群組中看到資源已成功建立，包括 Azure Bot、App Service、App Service 方案。 \
    ![img-alt-text](images/resource.png =600x)
<!-- 1. 如果在使用應用程式時遇到任何錯誤，可在 Azure App Services 中查看錯誤日誌。
    1. 開啟 [Azure Portal](https://ms.portal.azure.com/#home) 以您的 Azure 帳號登入。
    1. 前往已建立好的 Web App，選擇 **監視 ( Monitor ) > App Service 記錄 ( App Service logs )**。
    1. 啟用 **應用程式記錄 Application logging ( 檔案系統 Filesystem )** 並點擊 **儲存**。
    1. 接著可在 **記錄資料流 ( Log stream )** 中查看應用程式日誌。  
    
        <p> <span style="color:red;font-weight:bold">⚠ 注意：此設定會在 12 小時後自動關閉。若設定已自動關閉，請重新執行上述步驟。</span></p>

    ![enable-error-log](images/enable-error-log.png =700x) -->

# Step 5. 發佈應用程式至 Teams
將應用程式部署至 Azure 之後，我們會將應用程式上傳到 Teams 系統管理中心，並允許應用程式發佈到您組織內的 Teams 應用程式商店。

1. 開啟 Teams Toolkit，於  Lifecycle 中選擇 **Publish**（發佈），將應用程式發佈到 Teams 系統管理中心。 \
    ![teams-toolkit-lifecycle](images/publish.png =400x)
1. 前往 [**Teams 系統管理中心**](https://admin.teams.microsoft.com/) 以您的 Microsoft 365 帳號登入。
1. 選擇管理應用程式 > 查詢 ``Expert Finder`` > 選擇剛發佈的應用程式。 \
    ![img-alt-text](images/teams-admin-center-1.png =800x)
1. 允許發佈 ``Expert Finder dev`` 應用程式。 \
    ![img-alt-text](images/teams-admin-center-2.png =600x)
1. 開啟您的 Teams，前往 Teams 應用程式商店並安裝應用程式。 接著你就可以參考[下個章節](#在-teams-與-microsoft-365-copilot-使用應用程式)來使用此應用程式。\
    ![install-app](images/app-store.png =700x)

# Step 6. 在 Teams 與 Microsoft 365 Copilot 使用應用程式
以下示範如何在 Copilot 以及 Teams 聊天室中使用 Expert Finder。

## Step 6.1 在 Copilot 中依技能與辦公室地點搜尋
前往 Microsoft 365 Copilot 聊天介面。在介面右上角，您可以看到 Expert Finder 代理程式。點擊並開始使用 Expert Finder。

以下是一些範例提示：
1. `Find experts with skill in Azure.`
2. `Find experts with skill in Python and who are from Taipei.`

![demo](images/copilot-demo.gif =800x)

## Step 6.2 在 Teams 聊天室中使用 Message Extension
開啟 Teams 聊天室，點擊右下角「+」符號並選擇開啟 Expert Finder，在搜尋框中以技能作為關鍵字查詢專家，例如：Azure。

![demo](images/teams-message-extension-demo.gif =800x)

# Step 7. 利用 PowerShell 大量更新用戶資訊 (Optional)
本章節的目標為透過 PowerShell 腳本，批次自動化更新用戶的個人資訊，包含技能（Skills）與辦公室地點（OfficeLocation）。以下將依序說明如何準備用戶資料、安裝必要套件、建立驗證憑證、註冊 Microsoft Entra 應用程式、設定 API 權限，到最後執行批次更新指令。
<p> <span style="color:red;font-weight:bold">⚠ 若您有大量的使用者資訊需同步或維護，建議您使用此方法，以減少手動操作 Microsoft Entra 及 Graph Explorer 的負擔。</span></p>

## Step 7.1 複製儲存庫
請執行以下指令以複製儲存庫，若您先前已複製過此儲存庫，可以跳過此步驟。
```bash
git clone https://mtc-taiwan@dev.azure.com/mtc-taiwan/MTC%20Internal/_git/expert-finder
```

## Step 7.2 準備待更新的用戶資訊 CSV 檔
在利用 PowerShell 進行更新前，您需要先準備一份包含所有要更新的用戶資訊 CSV 檔。這份檔案可以使用 Microsoft Excel 製作。
1. 利用 Excel 開啟本專案儲存庫 `PowerShell` 資料夾中的 `UserSkills.csv`，其中有兩個範例資料您可以將其刪除。
1. 依照以下規範填寫待更新的用戶資料。

    | 欄位名稱             | 必填 | 說明                                                                 |
    |---------------------|------|----------------------------------------------------------------------|
    | `UserPrincipalName` | 是 | 用戶的電子郵件，例如 `user1@contoso.com`。 |
    | `Skills`            | 否 ( 若無需更新此欄位，請保持空白 ) | 用戶的技能清單。允許多個值，若需輸入多個技能，請以 **英文逗號 `,`** 分隔，例如：`Data Science, Python, Azure`。 |
    | `OfficeLocation`    | 否 ( 若無需更新此欄位，請保持空白 ) | 用戶的辦公室地點，例如：`Taipei`。

1. 完成填寫後，點選儲存檔案，無須更改檔案位置與名稱。
<p> <span style="color:red;font-weight:bold">⚠ 注意：請確保此 CSV 檔儲存於 PowerShell 資料夾中且命名為 UserSkills.csv，以確保後續步驟中使用到的相對路徑正確。</span></p>

## Step 7.3 安裝 Microsoft Graph PowerShell SDK
為了使用 PowerShell 與 Microsoft Graph 進行互動，您需要先安裝 [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview?view=graph-powershell-1.0)。此套件提供一系列指令，讓您可以在 PowerShell 中呼叫 Microsoft Graph API 並存取 Microsoft 365 資料。
1. 在應用程式搜尋列中輸入 PowerShell。右鍵點擊該項目，選擇 「以系統管理員身分執行」PowerShell。\
    ![demo](images/open-powershell.png =500x)
1. 執行以下指令以安裝 Microsoft Graph PowerShell SDK。
    ```
    Install-Module Microsoft.Graph -Scope CurrentUser
    ```
1. 利用以下指令確認模組是否安裝成功。
    ```
    Get-Module Microsoft.Graph -ListAvailable
    ```
1. 若成功安裝您將會看到以下畫面，顯示您安裝的位置與版本。 \
    ![img](images/graph-module-intall.png =600x)

## Step 7.4 建立 Self-signed Public Certificate 以驗證應用程式
在使用 Microsoft Graph PowerShell SDK 呼叫 API 時，應用程式需要透過憑證（Certificate）進行驗證，以確保存取安全且符合企業授權機制。本步驟將利用 PowerShell 建立一組自我簽署的公開憑證（Self-signed Public Certificate），並匯出成 `.cer` 檔案，供後續註冊至 Microsoft Entra 應用程式使用。
1. 在應用程式搜尋列中輸入 PowerShell。右鍵點擊該項目，選擇 「以系統管理員身分執行」PowerShell。
1. 利用以下指令切換到本專案儲存庫中的 `PowerShell` 資料夾。請把 `YourFolderPath` 替換成您自己的路徑， 例如: `cd "C:\Users\a-sallychen\Documents\expert-finder\PowerShell"`。
    ```
    cd "YourFolderPath"
    ```
1. 路徑切換完畢後，執行以下指令以生成憑證。
    ```
    $certname = "Certificate"
    $cert = New-SelfSignedCertificate -Subject "CN=$certname" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256
    Export-Certificate -Cert $cert -FilePath "./$certname.cer"
    ```
1. 若成功建立憑證，您會在 `PowerShell` 資料夾中看到 `Certificate.cer` 檔案。

> 若想了解更多關於如何建立 Self-signed Public Certificate 的資訊，可參考此[文件](https://learn.microsoft.com/en-us/entra/identity-platform/howto-create-self-signed-certificate)

## Step 7.5 註冊 Microsoft Entra 應用程式
為了能夠安全地存取 Microsoft Graph API，我們需要在 Microsoft Entra ID 中註冊一個應用程式，透過此應用程式，您可以指定 API 權限範圍，並將憑證與應用程式關聯，實現安全的 API 存取。本步驟將使用 PowerShell 腳本 ``RegisterAppOnly.ps1`` 自動完成註冊流程，包含建立應用程式、上傳 `.cer` 憑證檔案、並儲存應用程式的重要資訊至 `AppInfo.json`，方便後續連線使用。
1. 在應用程式搜尋列中輸入 PowerShell。右鍵點擊該項目，選擇 「以系統管理員身分執行」PowerShell。
1. 利用以下指令切換到本專案儲存庫中的 `PowerShell` 資料夾。請把 `YourFolderPath` 替換成您自己的路徑， 例如: `cd "C:\Users\a-sallychen\Documents\expert-finder\PowerShell"`。
    ```
    cd "YourFolderPath"
    ```
1. 路徑切換完畢後，執行以下指令，使用 `RegisterAppOnly.ps1` 腳本來自動註冊應用程式，並將憑證上傳至 Microsoft Entra。
    ```
    .\RegisterAppOnly.ps1 -AppName "Graph PowerShell Script" -CertPath "Certificate.cer"
    ```
1. 若應用程式成功註冊後，您會在 `PowerShell` 資料夾中看到 `AppInfo.json` 檔案，裡面儲存應用程式的 `ClientId` 與 `TenentId`。另外您也會在 [Microsoft Entra Admin Center](https://entra.microsoft.com) 中看到您剛建立的應用程式，如 [Step 7.6](#step-7.6-全域管理員授予應用程式-api-權限) 中圖一所示。

> 若想了解更多以 PowerShell 註冊 Microsoft Entra 應用程式的資訊，可參考此[文件](https://learn.microsoft.com/en-us/powershell/microsoftgraph/app-only?view=graph-powershell-1.0)。

##　Step 7.6 全域管理員授予應用程式 API 權限
完成應用程式註冊後，您需要為該應用程式授予適當的 Microsoft Graph API 權限。這些權限將決定應用程式是否能夠讀取或更新使用者資料、網站內容等。
<p> <span style="color:red;font-weight:bold">⚠ 注意：只有具備 Microsoft 365 全域管理員（Global Administrator）權限的帳號才能完成此授權步驟。</span></p>

1. 開啟 [Microsoft Entra Admin Center](https://entra.microsoft.com) 以具有全域管理員 ( Global Administrator ) 權限的 Microsoft 365 帳號登入。
1. 前往應用程式 ( Applications ) > 應用程式註冊 ( Application Registration )，找到剛才建立的應用程式 `Graph PowerShell Script`。
1. 開啟 `Graph PowerShell Script` 中的 API 權限 ( API Permissions )。 \
    ![img](images/app.png =600x)
1. 授予 `Sites.ReadWrite.All`、`User.Read.All` 以及 `User.ReadWrite.All` 的權限。 \
    ![img](images/grant-api-permission.png =600x)

> 若想了解更多關於 Microsoft Entra 內建角色與權限的資訊，可參考此[文件](https://learn.microsoft.com/zh-tw/entra/identity/role-based-access-control/permissions-reference)。

## Step 7.7 利用 PowerShell 大量更新用戶資訊
在完成應用程式註冊與授權後，您就可以透過 PowerShell 腳本批次更新用戶的資訊，包括技能（Skills）與辦公室地點（OfficeLocation）。本步驟將執行 `UpdateUserSkills.ps1` 腳本，自動讀取 `UserSkills.csv` 檔案中的資料，並依照內容逐一更新指定用戶的資訊。最後，透過指令驗證更新結果，確保資訊已成功寫入。
1. 在應用程式搜尋列中輸入 PowerShell。右鍵點擊該項目，選擇 「以系統管理員身分執行」PowerShell。
1. 利用以下指令切換到本專案儲存庫中的 `PowerShell` 資料夾。請把 `YourFolderPath` 替換成您自己的路徑， 例如: `cd "C:\Users\a-sallychen\Documents\expert-finder\PowerShell"`。
    ```
    cd "YourFolderPath"
    ```
1. 路徑切換完畢後，執行以下指令，使用 `UpdateUserSkills.ps1` 腳本來更新用戶資料。
    ```
    .\UpdateUserSkills.ps1
    ```
1. 您可以利用以下指令，確定用戶的資料是否更新成功。請把 `YourUserPrincipalName` 替換成您想查詢之用戶的電子信箱，例如 user1@contoso.com。
    ```
    $appInfo = Get-Content "./AppInfo.json" | ConvertFrom-Json
    Connect-MgGraph -ClientId $appInfo.ClientId -TenantId $appInfo.TenantId -CertificateName "CN=Certificate"
    Get-MgUser -UserId "YourUserPrincipalName" -select "Skills", "OfficeLocation"  | Select Skills, OfficeLocation
    Disconnect-MgGraph
    ```
<p>
  <span style="color:red;font-weight:bold">
    ⚠ 注意：每次要更新用戶資訊時，您只需要依照 
    <a href="#step-7-2-準備待更新的用戶資訊-csv-檔">Step 7.2 </a>準備待更新的用戶資訊 CSV 檔，再根據 
    <a href="#step-7-7-利用-powershell-大量更新用戶資訊">Step 7.7 </a>利用 PowerShell 大量更新用戶資訊即可，不需要重新註冊應用程式或授予權限。
  </span>
</p>