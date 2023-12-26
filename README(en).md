# GetLocalPath
# Convert the URL returned by Workbook.Path Property in Excel VBA on OneDrive to a local path.  

## Problem to be solved  
  
There is a problem with the Workbook.Path property returning a URL when Excel VBA runs on OneDrive. This makes it impossible to get the local path for that book and even FileSystemObject is not available.     
  
Several methods have been proposed to solve this problem. For personal OneDrive, there is a way to convert the URL path to a local path by processing the URL path as a string.
For a personal OneDrive, the URL path returned by the Workbook.Path property has the following form. \<CID> is the 16-digit number assigned to the individual, followed by the subfolder path \<FOLDER-PATH>..  
  
    https://d.docs.live.net/<CID>/<FOLDER-PATH>
  
At this time, the OneDrive local path can be converted as follows:    
  
    C:\Users\<USER-NAME>\OneDrive\<FOLDER-PATH>
    
In OneDrive for Business, however, this URL path can be complicated. Here is an example:  

    https://<TENANT>.sharepoint.com/sites/<SITE-NAME>/Shared Documents/<FOLDER-PATH>
    
    https://<TENANT>-my.sharepoint.com/personal/<USER-PRINCIPAL-NAME>/Documents/<FOLDER-PATH>
  
The URL paths listed here are only an example, and it is not easy to convert them to local paths using only string conversion. For example, the \<TENANT> in the URL path is not the same as the \<tenant name> in the locale path, so it cannot be used as is. Also, in SharePoint and Teams, you can add folders or shortcuts to OneDrive via "Sync" or "Add Shortcut to OneDrive", but if you have many of these folders or shortcuts, it is difficult to determine which URL path corresponds to which shortcut.   
  
For this reason, the URL path returned by the Workbook.Path property may not be converted to a local path using only string processing.  
  
## Proposed Solution  
