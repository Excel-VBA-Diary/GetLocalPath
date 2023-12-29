# GetLocalPath
# Convert the URL returned by Workbook.Path Property in Excel VBA on OneDrive to a local path.  

## Problem to be solved  
  
There is a problem with the Workbook.Path property returning a URL when Excel VBA runs on OneDrive. This makes it impossible to get the local path for that book and even FileSystemObject is not available.     
  
Several methods have been proposed to solve this problem. For personal OneDrive, there is a way to convert the URL path to a local path by processing the URL path as a string.
For a personal OneDrive, the URL path returned by the Workbook.Path property has the following form. \<CID> is the 16-digit number assigned to the individual, followed by the subfolder path \<FOLDER-PATH>..  
```  
https://d.docs.live.net/<CID>/<FOLDER-PATH>
```  
At this time, the OneDrive local path can be converted as follows:    
```  
C:\Users\<USER-NAME>\OneDrive\<FOLDER-PATH>
```    
In OneDrive for Business, however, this URL path can be complicated. Here is an example:  
```
https://<TENANT>.sharepoint.com/sites/<SITE-NAME>/Shared Documents/<FOLDER-PATH>
```
```    
https://<TENANT>-my.sharepoint.com/personal/<USER-PRINCIPAL-NAME>/Documents/<FOLDER-PATH>
```  
The URL paths listed here are only an example, and it is not easy to convert them to local paths using only string conversion. For example, the \<TENANT> in the URL path is not the same as the \<tenant name> in the locale path, so it cannot be used as is. Also, in SharePoint and Teams, you can add folders or shortcuts to OneDrive via "Sync" or "Add Shortcut to OneDrive", but if you have many of these folders or shortcuts, it is difficult to determine which URL path corresponds to which shortcut.   
  
For this reason, the URL path returned by the Workbook.Path property may not be converted to a local path using only string processing.  
  
## Proposed Solution  

### OneDrive mounting information
  
OneDrive mount information is located under the following registry key
```
\HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive
```
Under this registry key are the entries (subkeys) that are mounted in OneDrive. The hierarchical structure is shown in the following figure in the Registry Editor.  

![OneDrive-Registory-1](OneDrive-Registry-1.png)  
    
Each entry is registered with a pair of UrlNameSpace and MountPoint.    

![OneDrive-Registory-1](OneDrive-Registry-2.png) 
  
UrlNameSpace is the URL path to the SharePoint document library, and MountPoint is the local path under OneDrive. if there is an UrlNameSpace that matches the upper portion of the URL path returned by Workbook.Path the corresponding MountPoint for the UrlNameSpace can be found.
For example, assume the following case. 
```
UrlNameSpace ： https://xxxx.sharepoint.com/sites/Test/Shared Documents/  
MountPoint   ： c:\Users\diary\OneDrive - MyCompany\General - Work  
Workbook.Path： https://xxxx.sharepoint.com/sites/Test/Shared Documents/General/folder1 
```
Since the UrlNameSpace matches the upper portion of the URL path returned by the Workbook.Path property, we can determine that the Workbook exists in or under the local path of MountPoint.
From the structure and notational relationship of the document library on the SharePoint site, we know that /General in the URL path returned by the Workbook.Path property corresponds to MountPoint's \General - Work. 
Based on these, the URL path returned by Workbook.Path can be converted to the following local path
```
c:\Users\diary\OneDrive - MyCompany\General - Work\folder1
```
  
## GetLocalPath Function

The GetLocalPath function converts URL paths to local paths using OneDrive mount information.
If the argument is a local path, it returns the local path as it is without conversion, so it can be used universally by replacing ThisWorkbook.Path in the code with GetLocalPath(ThisWorkbook.Path), for example.
Module_GetLocalPath.bas is an exported VBA module, which contains the Get "LocalPath function.
You can import Module_GetLocalPath.bas as is or copy and paste the necessary parts.  
  
### Syntax
GetLocalPath(UrlPath, [UseCache])  

|Part|Description|
----|----
|UrlPath|Required.  String expression of URL path returned by Workbook.Path property.|
|UseCache|Optional. Specify True to use the cache or False to not use the cache. The GetLocalPath function reads the OneDrive mount information from the registry and stores it in the cache (Static variable), which is used on the second and subsequent calls to the GetLocalPath function to speed up processing. The cache is valid until the Excel book of the VBA macro is closed. Regardless of the UseCache setting, if 30 seconds have elapsed since the last time the cache was read, the registry is read again and the cache is updated.

### Return values

GetLocalPath returns a local path.

### Examples of Use
```
Dim localPath As String
localPath = GetLocalPath(ThisWorkbook.Path) 
'''
  
## Known Issue
  
The local path shown by MountPoint contains only the name of a target folder under the document library on the SharePoint site. For example, if the name of a target folder is the same as that of an upper-level folder, the upper-level folder may be mistakenly identified as the target folder. This issue will not happen if there is a subordinate folder with the same name as the target folder. Now investigating a workaround for this issue.
  
