# 內容說明

此資料夾內的檔案是關於如何利用 PowerShell 批次更新 Microsoft Entra 用戶中的技能（Skills）與辦公室地點（OfficeLocation）。
更詳細的說明請參考 ``expert-finder/README.md`` 中的 **Step 7**。

| 檔案名稱              | 說明 |
|-----------------------|------|
| `RegisterAppOnly.ps1` | 註冊一個具憑證驗證的 Microsoft Entra Application，供後續腳本使用。 |
| `UpdateUserSkills.ps1` | 讀取 `UserSkills.csv` 並透過 Microsoft Graph 更新使用者的 Skills 和 OfficeLocation。 |
| `UserSkills.csv`      | 使用者資料的範例檔案，包含用戶信箱、技能與辦公室地點。 |


---

# Description

This folder contains files related to how to use PowerShell to batch update Skills and OfficeLocation fields for users in Microsoft Entra.  
For more detailed instructions, please refer to **Step 7** in `expert-finder/README.md`.

| File Name              | Description |
|------------------------|-------------|
| `RegisterAppOnly.ps1`  | Registers a Microsoft Entra application with certificate-based authentication for use in subsequent scripts. |
| `UpdateUserSkills.ps1` | Reads from `UserSkills.csv` and updates users' Skills and OfficeLocation fields via Microsoft Graph. |
| `UserSkills.csv`       | Sample user data file containing user email addresses, skills, and office locations. |