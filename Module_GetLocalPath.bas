Attribute VB_Name = "Module_GetLocalPath"
Option Explicit

'-------------------------------------------------------------------------------
' OneDrive上のVBAでWorkbook.PathがURLを返す問題を解決する.
' レジストリのOneDriveマウントポイント情報を参照してローカルパスを取得する.
' Resolve the problem with Workbook.Path returning URL in VBA on OneDrive.
' Refer to the OneDrive mount point information in the registry to get the local path.
'
' 参照設定で「Microsoft Scripting Runtime」を有効にする.
' Check the "Microsoft Scripting Runtime" in the references dialog box.
'
' Arguments:
'   UrlPath: URL Path (Required, String)
'   UseCache: Use Mount Point Information Cache from Registry (Optional, Boolean)
'             True  = Use cache (Default)
'             False = Do not use cache
'   DebugMode: Debug mode switch (Optional, Boolean)
'             True  = Return mpiCache instead of local path (Debug Mode)
'             False = Return local path (Default, Normal Mode)
'
' Return Value:
'   Local Path (String)
'   Return null string if conversion to local path fails
'
' Usage:
'   Dim lp As String
'   lp = GetLocalPath(ThisWorkbook.Path)
'
' Author: Excel VBA Diary (@excelvba_diary)
' Created: December 29, 2023
' Last Updated: March 13, 2024
' Version: 1.006
' License: MIT
'-------------------------------------------------------------------------------

Public Function GetLocalPath(UrlPath As String, _
                             Optional UseCache As Boolean = True, _
                             Optional DebugMode As Boolean = False) As Variant
    
    If Not UrlPath Like "http*" And DebugMode = False Then
        GetLocalPath = UrlPath
        Exit Function
    End If
    
    'キャッシュがない場合、キャッシュ収集から30秒を超えた場合は、キャッシュを更新する
    'If no cache or more than 30 seconds since last update, the cache is updated
    
    Static mpiCache As Collection, lastUpdated As Date
    
    If Not (mpiCache Is Nothing Or Now - lastUpdated > 30 / 86400 Or UseCache = False) Then
        GoTo Already_Updated
    End If
    
    'STEP-1
    'レジストリーからすべてのOneDriveマウント情報を収集する
    'Collect all OneDrive mount information from registory
    
    Const HKEY_CURRENT_USER As Long = &H80000001
    Const S_HKEY_CURRENT_USER As String = "HKEY_CURRENT_USER\"
    Const TARGET_KEY As String = "SOFTWARE\SyncEngines\Providers\OneDrive"
    
    Dim mpi As Dictionary
    Set mpiCache = New Collection
    
    Dim objReg As Object
    Set objReg = CreateObject("WbemScripting.SWbemLocator"). _
                 ConnectServer(, "root\default").Get("StdRegProv")
    
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    
    Dim subKeySet As Variant, subKey As Variant
    Dim entryNameSet As Variant, entryName As Variant, entryValue As Variant
    Dim entryPath As String
    
    objReg.EnumKey HKEY_CURRENT_USER, TARGET_KEY, subKeySet
    If IsNull(subKeySet) Then Exit Function
    
    For Each subKey In subKeySet
        'すべてのエントリー名とその値を取得する
        'collect all entry names and their values
        objReg.EnumValues HKEY_CURRENT_USER, TARGET_KEY & "\" & subKey, entryNameSet
        If Not IsNull(entryNameSet) Then
            Set mpi = New Dictionary
            mpi.Add "ID", subKey
            For Each entryName In entryNameSet
                entryPath = S_HKEY_CURRENT_USER & TARGET_KEY & "\" & subKey & "\" & entryName
                entryValue = objShell.regRead(entryPath)
                mpi.Add entryName, entryValue
            Next
            mpiCache.Add mpi
            Set mpi = Nothing
        End If
    Next
    
    'STEP-2
    'OneDriveの同期情報を取得してマウント情報を補完する
    'Get OneDrive synchronization information to supplement mount information

    Const SETTINGS_PATH As String = "\AppData\Local\Microsoft\OneDrive\Settings"
    Const ODS_FOLDERS As String = "Business1,Business2,Personal"

    Dim odsFolder As Variant, odsPath As String, odsIndex As Long, cid As String
    Dim tempArray As Variant, tempFolder As String, j As Long, k As Long
    For Each odsFolder In Split(ODS_FOLDERS, ",")
        odsPath = Environ("USERPROFILE") & SETTINGS_PATH & "\" & odsFolder & "\"
        If Dir(odsPath) = "" Then GoTo Skip_Supplement
        cid = IniKeyValue(odsPath & "global.ini", "cid")
        If cid = "" Then GoTo Skip_Supplement
        tempArray = IniToArray(odsPath & cid & ".ini")
        If IsEmpty(tempArray) Then GoTo Skip_Supplement
        
        For j = LBound(tempArray) To UBound(tempArray)
            Select Case tempArray(j)(0)
                ' "Sync" and Root folder Case
                Case "libraryScope"
                    For Each mpi In mpiCache
                        If tempArray(j)(2) = mpi.Item("ID") Then
                            If tempArray(j)(13) = mpi.Item("MountPoint") Then
                                mpi.Add "MountFolder", ""
                                Exit For
                            End If
                        End If
                    Next
                ' "Sync" and subfolder Case
                Case "libraryFolder"
                    For Each mpi In mpiCache
                        If tempArray(j)(3) = mpi.Item("ID") Then
                            If tempArray(j)(5) = mpi.Item("MountPoint") Then
                                mpi.Add "MountFolder", Trim(tempArray(j)(7))
                                Exit For
                            End If
                        End If
                    Next
                ' "Add shortcut to OneDrive" Case
                Case "AddedScope"
                    For Each mpi In mpiCache
                        If tempArray(j)(2) = mpi.Item("ID") Then
                            If mpi.Item("UrlNamespace") Like tempArray(j)(4) & "*" Then
                                mpi.Add "FolderPath", Trim(tempArray(j)(10))
                                Exit For
                            End If
                        End If
                    Next
            End Select
        Next

Skip_Supplement:
    Next
    
    lastUpdated = Now
    

Already_Updated:
    
    'デバッグモードの場合はMPIキャッシュを返す
    'Return mpi cache (mpiCache) if in debug mode
    If DebugMode Then Set GetLocalPath = mpiCache: Exit Function
    
    '有効なOneDriveマウント情報が無ければ終了する
    'Exit if no valid OneDrive mount information
    If mpiCache.Count = 0 Then Exit Function
   
    'STEP-3
    'OneDriveマウント情報をもとにURLパスをローカルパスに変換する
    'Convert URL path to local path based on OneDrive mount information
   
    Dim strUrlNamespace As String, strLibraryType As String, strMountPoint As String
    Dim subPath As String, tmpLocalPath As String
    Dim mountFolderName As String, mountFolderPath As String
    Dim returnDir As String, errNum As Long
    Dim i As Long
    
    For Each mpi In mpiCache
        
        strUrlNamespace = mpi.Item("UrlNamespace")
        strLibraryType = LCase(mpi.Item("LibraryType"))
        
        If Right(strUrlNamespace, 1) = "/" Then
            strUrlNamespace = Left(strUrlNamespace, Len(strUrlNamespace) - 1)
        End If
        If LCase(mpi.Item("ID")) = "personal" Then
            strUrlNamespace = strUrlNamespace & "/" & mpi.Item("CID")
        End If
        
        If Not (UrlPath Like strUrlNamespace & "*") Then GoTo Skip_To_Next
        
        subPath = Replace(UrlPath, strUrlNamespace, "")
        subPath = Replace(subPath, "/", "\")
        strMountPoint = mpi.Item("MountPoint")
        If subPath = "" Or subPath = "\" Then
            tmpLocalPath = strMountPoint
            GoTo Verify_Folder_Exists
        End If
        
        Select Case True
        
            ' In case of MySite or personal
            Case LCase(mpi.Item("ID")) = "personal" Or _
                 LCase(mpi.Item("ID")) Like "business#"
                tmpLocalPath = strMountPoint & subPath
                GoTo Verify_Folder_Exists
        
            ' In case of mounting by "Add shortcut to OneDrive"
            Case mpi.Exists("FolderPath")
                mountFolderPath = mpi.Item("FolderPath")
                If mountFolderPath <> "" Then
                    mountFolderPath = "\" & Replace(mpi.Item("FolderPath"), "/", "\")
                End If
                If subPath Like mountFolderPath & "*" Then
                    tmpLocalPath = strMountPoint & Mid(subPath, Len(mountFolderPath) + 1)
                    GoTo Verify_Folder_Exists
                End If
                GoTo Skip_To_Next
     
            ' In case of mounting by "Sync"
            Case mpi.Exists("MountFolder")
                mountFolderName = mpi.Item("MountFolder")
                If mountFolderName = "" Then
                    tmpLocalPath = strMountPoint & subPath
                    GoTo Verify_Folder_Exists
                End If
                i = InStr(1, subPath & "\", "\" & mountFolderName & "\")
                If i > 0 Then
                    subPath = Mid(subPath, i + Len(mountFolderName) + 1)
                    tmpLocalPath = strMountPoint & subPath
                    GoTo Verify_Folder_Exists
                End If
                GoTo Skip_To_Next

            ' In case of unexpected
            Case Else
                Exit Function
        
        End Select

Skip_To_Next:
    Next

    'URLパスに該当するマウント情報がないためローカルパスへの変換に失敗した
    'Conversion to local path failed due to lack of mount information corresponding to URL path
    Exit Function

Verify_Folder_Exists:
                   
    'STEP-4
    '実際にフォルダーが存在するか確認する
    'Verify that the folder actually exists
    
    On Error Resume Next
    returnDir = Dir(tmpLocalPath, vbDirectory)
    errNum = Err.Number
    On Error GoTo 0
    If errNum <> 0 Or returnDir = "" Then Exit Function

    GetLocalPath = tmpLocalPath

End Function


'-------------------------------------------------------------------------------
' INIファイルから指定されたキーの値を取得する
' Get the value of a specified key from an INI file
'-------------------------------------------------------------------------------

Public Function IniKeyValue(IniFilePath As String, keyName As String) As String
    
    If IniFilePath = "" Or Dir(IniFilePath) = "" Then
        IniKeyValue = ""
        Exit Function
    End If
    
    'Reads INI file in UTF-16 format
    Dim fileNumber As Long, byteBuf() As Byte, lineBuf As Variant
    fileNumber = FreeFile
    Open IniFilePath For Binary Access Read As #fileNumber
    ReDim byteBuf(LOF(fileNumber) - 1)
    Get #fileNumber, , byteBuf
    lineBuf = Split(CStr(byteBuf), vbCrLf)
    Close #fileNumber
    
    Dim i As Long, p As Long
    For i = LBound(lineBuf) To UBound(lineBuf)
        If LCase(lineBuf(i)) Like LCase(keyName) & "*" Then
            p = InStr(lineBuf(i), "=")
            If p = 0 Then Exit Function
            IniKeyValue = Trim(Mid(lineBuf(i), p + 1))
            Exit Function
        End If
    Next
    
    'The specified key was not found.
    IniKeyValue = vbNullString

End Function


'-------------------------------------------------------------------------------
' INIファイルを開きOneDrive設定情報をジャグ配列に取り込む
' Open INI file and import OneDrive setting information into jag array
'-------------------------------------------------------------------------------

Public Function IniToArray(IniFilePath As String) As Variant
    
    If IniFilePath = "" Or Dir(IniFilePath) = "" Then
        IniToArray = Empty
        Exit Function
    End If
    
    'Reads INI file in UTF-16 format
    Dim fileNumber As Long, byteBuf() As Byte, lineBuf As Variant
    fileNumber = FreeFile
    Open IniFilePath For Binary Access Read As #fileNumber
    ReDim byteBuf(LOF(fileNumber) - 1)
    Get #fileNumber, , byteBuf
    lineBuf = Split(CStr(byteBuf), vbCrLf)
    Close #fileNumber
    
    Const TARGET_KEYS As String = "libraryScope,libraryFolder,AddedScope,library,dummy"
    
    Dim i As Long, j As Long, p As Long, entryBuf As String, keyName As String, keyBuf As Variant
    Dim pArray() As Variant, pCount As Long
    ReDim pArray(UBound(lineBuf)): pCount = 0
    Dim cArray() As Variant, cCount As Long
    
    For i = 0 To UBound(lineBuf)
        entryBuf = lineBuf(i)
        p = InStr(entryBuf, "=")
        If p = 0 Then GoTo Skip_Convert
        keyName = Trim(Left(entryBuf, p - 1))
        If InStr(TARGET_KEYS, keyName & ",") = 0 Then GoTo Skip_Convert
        
        ReDim cArray(0): cArray(0) = keyName
        cCount = 1
        j = p + 1
        Do While j <= Len(entryBuf)
            Select Case Mid(entryBuf, j, 1)
                Case """"
                    p = InStr(j + 1, entryBuf, """")
                    If p = 0 Then p = Len(entryBuf) + 1
                    ReDim Preserve cArray(cCount)
                    cArray(cCount) = Mid(entryBuf, j + 1, p - j - 1)
                    cCount = cCount + 1
                    j = p + 1
                Case " "
                    j = j + 1
                Case Else
                    p = InStr(j + 1, entryBuf, " ")
                    If p = 0 Then p = Len(entryBuf) + 1
                    ReDim Preserve cArray(cCount)
                    cArray(cCount) = Mid(entryBuf, j, p - j)
                    cCount = cCount + 1
                    j = p + 1
            End Select
        Loop
        ReDim Preserve pArray(pCount)
        pArray(pCount) = cArray
        pCount = pCount + 1
Skip_Convert:
    Next
    
    IniToArray = pArray
    
End Function


'-------------------------------------------------------------------------------
' テストコード
' Test code for GetLocalPath
'-------------------------------------------------------------------------------

Private Sub Functional_Test_GetLocalPath()
    Debug.Print "URL Path", ThisWorkbook.Path
    Debug.Print "Local Path", GetLocalPath(ThisWorkbook.Path, False)
End Sub

Private Sub Speed_Test_GetLocalPath()
    Dim i As Long, t1 As Single
    t1 = Timer
    For i = 1 To 100
        Call GetLocalPath(ThisWorkbook.Path, False)     'Cache Disable
    Next
    Debug.Print "UseCache Disable: "; Format(Timer - t1, "#0.0000000"); " [Sec]"
    t1 = Timer
    For i = 1 To 100
        Call GetLocalPath(ThisWorkbook.Path, True)      'Cache Enable
    Next
    Debug.Print "UseCache Enable:  "; Format(Timer - t1, "#0.0000000"); " [Sec]"
End Sub

'-------------------------------------------------------------------------------
' このモジュールはここで終わり
' The script for this module ends here
'-------------------------------------------------------------------------------


