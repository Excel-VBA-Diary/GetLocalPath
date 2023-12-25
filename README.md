# GetLocalPath
# OneDrive上のExcel VBAでWorkbook.Pathプロパティが返すURLをローカルパスに変換する。  
# Convert the URL returned by Workbook.Path Property in Excel VBA on OneDrive to a local path.  

## 解決したい問題  
  
OneDrive上のExcel VBAを動かすとWorkbook.Path プロパティがURLを返す問題が起きます。自分自身のローカルパスを取得できず、FileSystemObjectまで使えなくなるという不便な状態になります。  
  
この問題の解決にはいくつかの方法が提案されていますが、URLパスを文字列処理してローカルパスに変換する方法はうまくいかない場合があります。特に OneDrive for Business においてはURLに含まれるテナントコードをテナント名に変換するなどの処理が必要で、文字列処理による方法では解決できません。  
  
SharePointやTeamsでは「OneDriveへのショートカットの追加」によってOneDriveにショートカットを追加できますが、URLパスがどのショートカットに対応するか判別することは困難です。  
  
このような理由からThisWorkbook.Pathが返すURLを文字列処理によってローカルパスに変換する方法には、事実上無理があります。

## 提案する解決策  

