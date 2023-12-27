# GetLocalPath
# OneDrive上のExcel VBAでWorkbook.Pathプロパティが返すURLをローカルパスに変換する。  
#### README(en).md for English version    

## 解決したい問題  
  
OneDrive上のExcel VBAを動かすとWorkbook.Path プロパティがURLを返す問題が起きます。そのブックのローカルパスを取得できず、FileSystemObjectまで使えなくなるという不便な状態になります。  
  
この問題の解決にはいくつかの方法が提案されています。個人用OneDriveであればURLパスを文字列処理してローカルパスに変換する方法があります。
個人用OneDriveの場合、Workbook.Path プロパティが返すURLは次の形式となります。\<CID>は個人用に割り当てられた16桁の番号で、その後にサブフォルダのパス\<FOLDER-PATH>が続きます。  
  
    https://d.docs.live.net/<CID>/<FOLDER-PATH>
  
この時、OneDriveのローカルパスは次のように変換できます。  
  
    C:\Users\<USER-NAME>\OneDrive\<FOLDER-PATH>
    
しかし、OneDrive for Business においては、このURLパスが複雑になります。以下はその例です。  

    https://<TENANT-NAME>.sharepoint.com/sites/<SITE-NAME>/Shared Documents/<FOLDER-PATH>
    
    https://<TENANT-NAME>-my.sharepoint.com/personal/<UPN>/Documents/<FOLDER-PATH>
  
ここに挙げたURLパスは一例に過ぎず、これを文字列変換だけでローカルパスに変換するのは簡単ではありません。例えば、URLパスに含まれる\<TENANT-NAME>はロカールパスに含まれる<テナント名>とは異なるのでそのまま使えません。また、SharePointやTeamsでは「同期」または「OneDriveへのショートカットの追加」によってOneDriveにフォルダーまたはショートカットフォルダーを追加できますが、一般的にローカルパスに含まれる<テナント名>の表記が微妙に異なります。  
  
「同期」の場合：  
    C:\Users\<USER-NAME>\<テナント名>\<フォルダー名>
  
「OneDriveへのショートカットの追加」の場合：  
    C:\Users\<USER-NAME>\OneDrive - <テナント名>\<フォルダー名>
  
このように「同期」と「OneDriveへのショートカットの追加」が混在したり、フォルダーが多数ある場合、URLパスがどのフォルダーに対応するかURLパスから判別するのは大変難しくなります。  
このような理由から、Workbook.Path プロパティが返すURLパスを文字列処理だけでローカルパスに変換できないことがあります。  
  
## 提案する解決策  

