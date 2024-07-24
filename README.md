# Activity Explorer Export script

PowerShell script to simplify the execution of Export-ActivityExplorerData cmdlet.

Ways to use the script:

> **NOTES**
>
>  **By default**, the data is exported in JSON format to a folder named **"ExportedData."** This folder is automatically created in the same directory where the script is executed, if it doesn't already exist.
> **By default**, The configuration file called **"ActivityExplorerActivitiesSelector.json"** contains a list of 26 activities. This list is manually created and may not be complete. The file is provided to help focus on specific activities, such as **"DLPRuleMatch"** or **"LabelApplied"**. By default, all values are set to "True." You can set any activities you don't want to "False."
> If the configuration file is not present when the script is executed, it will include all possible activities. You can achieve the same result by using the **"-NotApplyFilters"** attribute. Additionally, this configuration file can be edited to add new activities as needed.
<br>  
<br>  


💠**Description:** Using only the script by default, you'll be able to capture and calculate data from the past 24 hours, based on the exact time of execution.
```
.\ActivityExplorerExport.ps1
```
<br>  
<br>  

💠**Description:** With this next execution, you'll be able to capture and calculate data from the past X days, based on the exact time of execution..
```
.\ActivityExplorerExport.ps1 -PastDays 31
```
<br>  
<br>  

💠**Description:** This script uses a configuration file located in the same folder as the script. The file, called ActivityExplorerActivitiesSelector.json, contains a list of common activities and is used by default with the values set to "True." It's possible that some activities are not listed; you can manually add those. If the file is not present, the script will take all possible activities. Alternatively, you can use the "-NotApplyFilters" attribute to retrieve all activities.
```
.\ActivityExplorerExport.ps1 -NotApplyFilters
```
<br>  
<br>  

💠**Description:** You can request a specific time frame for data retrieval. For example, to get all the data for June 18, 2024, you can set the desired interval. This can be a short interval or an extended period spanning several days.
```
.\ActivityExplorerExport.ps1 -UseCustomParameters -AEStartTime "06/18/2024 00:00:00" -AEEndTime "06/18/2024 23:59:59"
```
<br>  
<br>  

💠**Description:** If you are not comfortable working with JSON format, you can use the attribute "-ExportToCsv" to export the data in CSV format.
```
.\ActivityExplorerExport.ps1 -ExportToCsv
```
<br>  
<br>  

💠**Description:** The PowerShell cmdlet used in this script supports a maximum Page Size of 5,000, which is the default value. However, if you want to change this Page Size, you can use the attribute -InitialPageSize and set a smaller value.
```
.\ActivityExplorerExport.ps1 -InitialPageSize 300
```
<br>  
<br>  

💠**Description:** You can combine different attributes available in the script to customize its functionality. For example:
```
.\ActivityExplorerExport.ps1 -NotApplyFilters -PastDays 15
```
<br>  
<br>  

When the script is executed, a summary will be displayed in the console, showing the configuration settings for that execution.
![image](https://github.com/user-attachments/assets/05d84980-8af5-44f4-94db-fa2bf00cfdc3)

