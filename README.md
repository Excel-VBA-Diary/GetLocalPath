# GetLocalPath
# OneDrive上のExcel VBAでWorkbook.Pathプロパティが返すURLをローカルパスに変換する。  
#### README(en).md for English version    

## 解決したい問題  
  
OneDrive上のExcel VBAを動かすとWorkbook.Path プロパティがURLを返す問題が起きます。そのブックのローカルパスを取得できず、URLのままではDir関数が実行時エラーになったり、FileSystemObjectが使えなくなるなど不便な状態になります。  
  
この問題の解決にはいくつかの方法が提案されています。個人用OneDriveであればURLパスを文字列処理してローカルパスに変換する方法があります。
個人用OneDriveの場合、Workbook.Path プロパティが返すURLは次の形式となります。\<CID>は個人用に割り当てられた16桁の番号で、その後にOneDrive配下のフォルダのパス\<FOLDER-PATH>が続きます。  
```  
https://d.docs.live.net/<CID>/<FOLDER-PATH>
```  
この時、OneDriveのローカルパスは次のように変換できます。  
```  
C:\Users\<USER-NAME>\OneDrive\<FOLDER-PATH>
```    
個人用OneDriveの場合、ローカルパスへの変換は比較的容易です。しかし、OneDrive for Business においては、このURLパスが複雑になります。以下はその典型例です。  
```
https://<TENANT-NAME>.sharepoint.com/sites/<SITE-NAME>/Shared Documents/<FOLDER-PATH>
```
```    
https://<TENANT-NAME>-my.sharepoint.com/personal/<UPN>/Documents/<FOLDER-PATH>
```  
エクスプローラーを使ってSharePointやTeamsのファイルにアクセスする場合、「同期」と「OneDriveへのショートカットの追加」の二つの方法があります。生成されるローカルパスは次のとおりです。 
  
「同期」の場合：  
```
C:\Users\<USER-NAME>\<テナント名>\<フォルダーパス>
```  
「OneDriveへのショートカットの追加」の場合：  
```
C:\Users\<USER-NAME>\OneDrive - <テナント名>\<フォルダーパス>
```
  
「同期」と「OneDriveへのショートカットの追加」ではローカルパスの表記が微妙に異なります。また、ロカールパスに含まれる<テナント名>はURLパスに含まれる\<TENANT-NAME>とは異なります。さらにロカールパスに含まれる<フォルダーパス>は
URLパスに含まれる\<FOLDER-PATH>と必ずしも一致しません。ここに挙げたURLパスもローカルパスも一例に過ぎず、文字列変換だけでURLパスをローカルパスに変換するのは事実上無理です。  
  
## 提案する解決策 

# OneDriveのマウント情報
  
OneDriveのマウント情報は次のレジストキー配下にあります。
```
\HKEY_CURRENT_USER\Software\SyncEngines\Providers\OneDrive
```
このレジストリーキーの配下にはOneDriveにマウントされている数だけのサブキー（エントリー）があり、それぞれのエントリーにはUrlNameSpaceとMountPointが対になって登録されています。
UrlNameSpaceはSharePointのドキュメントライブラリーのURLパス、MountPointはOneDrive配下のフォルダーパスを示しています。Workbook.Pathが返すURLパスの上位部分と一致するUrlNameSpaceがあれば、そのUrlNameSpaceに対応するMountPointがわかります。

## 既知の問題
  
マウントポイント（MountPoint）は、SharePointサイトのフルパスではなくフィルダー名だけなので、マウントしたフォルダーが上位フォルダーと同一名の場合、誤って認識する場合があります。  


