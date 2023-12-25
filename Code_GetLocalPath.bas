Attribute VB_Name = "Code_GetLocalPath"
Option Explicit

'-------------------------------------------------------------------------------
' OneDrive上のVBAでWorkbook.PathがURLを返す問題を解決する.
' レジストリのOneDriveマウントポイント情報を参照してローカルパスを取得する.
' Resolve the problem with Workbook.Path returning URL in VBA on OneDrive.
' Refer to the OneDrive mount point information in the registry to get the local path.
' Arguments:
'   UrlPath: URL Path (String)
'   UseCache: Use Mount Point Information Cache from Registry (Boolean)
'             True  = Use cache (Cache Enable)
'             False = Do not use cache (Cache Disable)
' Return Value:
'   Local Path (String)
'   Return null string if conversion to local path fails
'
' Author: Excel VBA Diary (@excelvba_diary)
' Last Update: December 25, 2023
' License: MIT
'-------------------------------------------------------------------------------

Public Function GetLocalPath(UrlPath As String, _
                             Optional UseCache As Boolean = True) As String
    
    If Not UrlPath Like "http*" Then
        GetLocalPath = UrlPath
        Exit Function
    End If
   
    'すべてのOneDriveマウントポイント情報を収集する
    'Collect all OneDrive mount point information
    
    Static mpiCache As Collection
    Dim mpi As Dictionary
    
    If UseCache And (Not mpiCache Is Nothing) Then GoTo Already_Collected
    
    Set mpiCache = New Collection
    
    Const HKEY_CURRENT_USER = &H80000001
    Const S_HKEY_CURRENT_USER = "HKEY_CURRENT_USER\"
    Const TARGETKEY = "SOFTWARE\SyncEngines\Providers\OneDrive"
    
    Dim objReg As Object
    Set objReg = CreateObject("WbemScripting.SWbemLocator"). _
                 ConnectServer(vbNullString, "root\default").Get("StdRegProv")
    
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    
    Dim subKeySet As Variant, subKey As Variant
    Dim entryNameSet As Variant, entryName As Variant, entryValue As Variant
    Dim entryPath As String
    
    objReg.EnumKey HKEY_CURRENT_USER, TARGETKEY, subKeySet
    If IsNull(subKeySet) Then Exit Function
    
    For Each subKey In subKeySet
        'すべてのエントリー名とその値を取得する
        'Get all entry names and their values
        objReg.EnumValues HKEY_CURRENT_USER, TARGETKEY & "\" & subKey, entryNameSet
        If Not IsNull(entryNameSet) Then
            Set mpi = New Dictionary
            mpi.Add "GUID", subKey
            For Each entryName In entryNameSet
                entryPath = S_HKEY_CURRENT_USER & TARGETKEY & "\" & subKey & "\" & entryName
                entryValue = objShell.regRead(entryPath)
                mpi.Add entryName, entryValue
            Next
            mpiCache.Add mpi
            Set mpi = Nothing
        End If
    Next
    
Already_Collected:
    
    '有効なOneDriveマウントポイント情報が無ければ終了する
    'Exit if no valid OneDrive mount point information
    If mpiCache.Count = 0 Then Exit Function
   
    Dim strUrlNamespace As String, strMountPoint As String, strLibraryType As String
    Dim tmpLocalPath As String, tmpSubPath As String
    
    '個人用OneDriveのURLパスをローカルパスに変換する
    'Convert personal OneDrive URL path to local path

    If UrlPath Like "https://d.docs.live.net/????????????????*" Then
        'Remove CID from personal OneDrive URL path for comparison with mount point information
        UrlPath = Left(UrlPath, 23) & Mid(UrlPath, 41)
        For Each mpi In mpiCache
            strUrlNamespace = mpi.Item("UrlNamespace")
            strMountPoint = mpi.Item("MountPoint")
            If UrlPath Like strUrlNamespace & "*" Then
                tmpSubPath = Replace(UrlPath, strUrlNamespace, "")
                tmpSubPath = Replace(tmpSubPath, "/", "\")
                tmpLocalPath = strMountPoint & tmpSubPath
                GoTo Verify_Folder_Exists
            End If
        Next
        'No corresponding NameSpace for UrlPath
        Exit Function
    End If
    
    '個人用OneDrive以外のURLパスをローカルパスに変換する
    'Convert non-personal OneDrive URL path to local path
    
    Dim pathTree As Variant, i As Long
    
    For Each mpi In mpiCache
        strUrlNamespace = mpi.Item("UrlNamespace")
        strMountPoint = mpi.Item("MountPoint")
        strLibraryType = mpi.Item("LibraryType")
        Select Case True
            Case UrlPath & "/" = strUrlNamespace
                tmpLocalPath = strMountPoint
                GoTo Verify_Folder_Exists

            Case UrlPath = strUrlNamespace & "General"
                 If strLibraryType = "mysite" Or strLibraryType = "personal" Then
                    tmpLocalPath = strMountPoint & "\General"
                 Else
                    tmpLocalPath = strMountPoint
                 End If
                 GoTo Verify_Folder_Exists
            
            Case UrlPath Like strUrlNamespace & "*"
                tmpSubPath = "/" & Replace(UrlPath, strUrlNamespace, "")
                If tmpSubPath Like "/General/*" Then tmpSubPath = Mid(tmpSubPath, 9)
                tmpSubPath = Replace(tmpSubPath, "/", "\")
                pathTree = Split(strMountPoint, "\")
                If UBound(pathTree) = 4 Then
                    i = InStr(1, tmpSubPath, "\" & pathTree(4))
                    If i = 0 Then
                        tmpLocalPath = strMountPoint & tmpSubPath
                        GoTo Verify_Folder_Exists
                    End If
                    Do
                        tmpSubPath = Mid(tmpSubPath, i + Len(pathTree(4)) + 1)
                        tmpLocalPath = strMountPoint & tmpSubPath
                        If Dir(tmpLocalPath, vbDirectory) <> "" Then
                            GetLocalPath = tmpLocalPath
                            Exit Function
                        End If
                        i = InStr(i, tmpSubPath, "\" & pathTree(4))
                    Loop While i > 0
                Else
                    tmpLocalPath = strMountPoint & tmpSubPath
                    GoTo Verify_Folder_Exists
                End If
        
        End Select
    Next

    'No corresponding NameSpace for UrlPath
    Exit Function

Verify_Folder_Exists:
                   
    If Dir(tmpLocalPath, vbDirectory) <> "" Then GetLocalPath = tmpLocalPath

End Function


'-------------------------------------------------------------------------------
' テストコード
' Test code for GetLocalPath
'-------------------------------------------------------------------------------

Private Sub Functinal_Test_GetLocalPath()
    Debug.Print "URL Path", ThisWorkbook.path
    Debug.Print "Local Path", GetLocalPath(ThisWorkbook.path)
End Sub


Private Sub Speed_Test_GetLocalPath()
    Dim i As Long, t1 As Single
    t1 = Timer
    For i = 1 To 100
        Call GetLocalPath(ThisWorkbook.path, False)
    Next
    Debug.Print "UseCache Disable:"; Timer - t1; "[Sec]"
    t1 = Timer
    For i = 1 To 100
        Call GetLocalPath(ThisWorkbook.path, True)
    Next
    Debug.Print "UseCache Enable: "; Timer - t1; "[Sec]"
End Sub

'-------------------------------------------------------------------------------
' このモジュールはここで終わり
' The script for this module ends here
'-------------------------------------------------------------------------------

