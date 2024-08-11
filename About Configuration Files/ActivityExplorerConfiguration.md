## Activity Explorer Configuration file

This file is used to configure the connection to a Logs Analytics workspace.
The values required are:
- **EncryptedKeys** : This value is set on `False` by default, if you encrypt the key you need to change the value to `True`. 
- **TenantGUID** : This is your Tenant ID
- **TenantDomain** : This is your primary domain set in your Tenant
- **Workspace_ID** : This is the Logs Analytics workspace ID
- **WorkspacePrimaryKey** : This is the Logs Analytics Primary Key

<details>
<summary>ActivityExplorerConfiguration.json</summary>
  
```JSON
{
    "EncryptedKeys": "False",
    "TenantGUID": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
    "TenantDomain": "kazlivedemos.com",
    "Workspace_ID": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
    "WorkspacePrimaryKey": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
}
```
</details>

### How to Encrypt keys

The easy way to encrypt your key is using the next cmdlet, this cmdlet uses the Logged user and the Machine ID to hash the password
```PowerShell
"your_secret" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
```

![image](https://github.com/user-attachments/assets/c2b3b4a8-47f0-4b7d-b75c-d4b580f79dbb)

The value result from the cmdlet need to be paste in the field `WorkspacePrimaryKey`

To see how to get that primarykey, check the section related to that point in this same document.

### How to get my Tenant ID and Primary Domain

At https://portal.azure.com you need to find Microsoft Entra ID menu, you can do it in the search section at top center, and then selecting Microsoft Entra ID

![image](https://github.com/user-attachments/assets/3a6a8973-1934-465f-949f-79a96b489c05)

In the first page that correspond to the Tenant Overview, you can get the Tenant ID and the Primary domain, you need to add those values in the configuration file in the values `TenantGUID` and `TenantDomain`in that same oder.

![image](https://github.com/user-attachments/assets/61e33825-c790-4706-996f-b22434bac5b5)

<details>
<summary>ActivityExplorerConfiguration.json with the values recently aquired</summary>
  
```JSON
{
    "EncryptedKeys": "False",
    "TenantGUID": "ac1dff03-7e0e-4ac8-a4c9-9b38d24f062c",
    "TenantDomain": "kazdemos.org",
    "Workspace_ID": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
    "WorkspacePrimaryKey": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
}
```
</details>

### How to create a Logs Analytics workspace

At https://portal.azure.com you need to find Logs Analytics workspaces menu, you can do it in the search section at top center, and then selecting Logs Analytics workspaces

![image](https://github.com/user-attachments/assets/39155308-958e-46ae-9b9d-33a5330f2e61)

After selecting this option, you will be redirect to Logs Analytics workspaces menu, where you can press "+ Create" to create a new workspace

![image](https://github.com/user-attachments/assets/04357c55-36a9-4d4b-84af-0e922256e040)

In the new interface you need to select the right Azure subscription, then select a Resource Group or Create a new one. The last step is set a name for our Logs Analytics workspace and the region where this one will be located. To finish pressing "Review + Create" button at the bottom of the page.

![image](https://github.com/user-attachments/assets/2eca94bd-e28a-4455-b90d-ea52d7becdaf)

Finally press "Create" and wait until the process finish, this process normally can take a couple of minutes.

![image](https://github.com/user-attachments/assets/25dddc10-273a-4224-a455-3c3aa66ce4dc)

### How to get Logs Analytics workspace ID and Primary Key

At the Logs Analytics workspaces menu, you need to find your workspace name and select it.
In the workspace interface,  at left you need to extend Settings and select Agents.

![image](https://github.com/user-attachments/assets/49422970-f881-4c86-87a9-0bda0ae8b32b)

In the new windows you need to extend the submenu called "Log Analytics agent instructions"
Now the information required was displayed, you need to get the values for Workspace ID and Primary key and replace the values `Workspace_ID`and `WorkspacePrimaryKey` on the configuration file.

![image](https://github.com/user-attachments/assets/8ce26eb4-41a7-4e49-8493-8645d18a5f9b)

<details>
<summary>ActivityExplorerConfiguration.json with the values recently aquired</summary>
  
```JSON
{
    "EncryptedKeys": "False",
    "TenantGUID": "ac1dff03-7e0e-4ac8-a4c9-9b38d24f062c",
    "TenantDomain": "kazdemos.org",
    "Workspace_ID": "f6f62ab8-4f92-4397-9543-99a23c4e8747",
    "WorkspacePrimaryKey": "/wg+g8dEkd34MeW7RMb8OuXkik3oFFsQds/Ae5W68FXdKj5KV2NaOj8K86ii7QmZ2Uo0llw++JF60/ThAa8tlQ=="
}
```
</details>

<br>
<br>

> Now You have your file ready to use
<br>
<br>
