Attribute VB_Name = "Module_GetLocalPath"
Option Explicit

'-------------------------------------------------------------------------------
' OneDrive���VBA��Workbook.Path��URL��Ԃ�������������.
' ���W�X�g����OneDrive�}�E���g�|�C���g�����Q�Ƃ��ă��[�J���p�X���擾����.
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
' Usage:
'   Dim lp As String
'   lp = GetLocalPath(ThisWorkbook.Path)
'
' Author: Excel VBA Diary (@excelvba_diary)
' Last Update: December 26, 2023
' License: MIT
'-------------------------------------------------------------------------------

Public Function GetLocalPath(UrlPath As String, _
                             Optional UseCache As Boolean = True) As String
    
    If Not UrlPath Like "http*" Then
        GetLocalPath = UrlPath
        Exit Function
    End If
   
    '���ׂĂ�OneDrive�}�E���g�|�C���g�������W����
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
        '���ׂẴG���g���[���Ƃ��̒l���擾����
        'Get all entry names and their values
        objReg.EnumValues HKEY_CURRENT_USER, TARGETKEY & "\" & subKey, entryNameSet
        If Not IsNull(entryNameSet) Then
            Set mpi = New Dictionary
            mpi.Add "ID", subKey
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
    
    '�L����OneDrive�}�E���g�|�C���g��񂪖�����ΏI������
    'Exit if no valid OneDrive mount point information
    If mpiCache.Count = 0 Then Exit Function
   
    Dim strUrlNamespace As String, strMountPoint As String
    Dim strLibraryType As String, isFolderScope As Boolean
    Dim tmpLocalPath As String, tmpSubPath As String
    Dim returnDir As String, errNum As Long
    
    '�l�pOneDrive��URL�p�X�����[�J���p�X�ɕϊ�����
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
    
    '�l�pOneDrive�ȊO��URL�p�X�����[�J���p�X�ɕϊ�����
    'Convert non-personal OneDrive URL path to local path
    
    Dim pathTree As Variant, i As Long
    
    For Each mpi In mpiCache
        strUrlNamespace = mpi.Item("UrlNamespace")
        If strUrlNamespace Like "*/" Then
            strUrlNamespace = Left(strUrlNamespace, Len(strUrlNamespace) - 1)
        End If
        If Not (UrlPath Like strUrlNamespace & "*") Then GoTo Skip_To_Next
        
        strLibraryType = mpi.Item("LibraryType")
        strMountPoint = mpi.Item("MountPoint")
        isFolderScope = CBool(mpi.Item("IsFolderScope"))
            
        If strLibraryType = "mysite" Or strLibraryType = "personal" Then
            tmpSubPath = Replace(UrlPath, strUrlNamespace, "")
            tmpSubPath = Replace(tmpSubPath, "/", "\")
            If tmpSubPath = "" Or tmpSubPath = "\" Then
                tmpLocalPath = strMountPoint
            Else
                tmpLocalPath = strMountPoint & tmpSubPath
            End If
            GoTo Verify_Folder_Exists
        Else
            tmpSubPath = Replace(UrlPath, strUrlNamespace, "")
            tmpSubPath = Replace(tmpSubPath, "/", "\")
            If tmpSubPath = "" Or tmpSubPath = "\" Then
                tmpLocalPath = strMountPoint
                GoTo Verify_Folder_Exists
            End If
            If isFolderScope Then
                If tmpSubPath Like "\General*" Then tmpSubPath = Mid(tmpSubPath, 9)
            End If
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
                    On Error Resume Next
                    returnDir = Dir(tmpLocalPath, vbDirectory)
                    errNum = Err.Number
                    On Error GoTo 0
                    If errNum = 0 And returnDir <> "" Then
                        GetLocalPath = tmpLocalPath
                        Exit Function
                    End If
                    i = InStr(i, tmpSubPath, "\" & pathTree(4))
                Loop While i > 0
            Else
                tmpLocalPath = strMountPoint & tmpSubPath
                GoTo Verify_Folder_Exists
            End If
        End If
Skip_To_Next:
    Next

    'No corresponding NameSpace for UrlPath
    Exit Function

Verify_Folder_Exists:
                   
    '���ۂɃt�H���_�[�����݂��邩�m�F����
    'Verify that the folder actually exists
    
    On Error Resume Next
    returnDir = Dir(tmpLocalPath, vbDirectory)
    errNum = Err.Number
    On Error GoTo 0
    If errNum = 0 And returnDir <> "" Then GetLocalPath = tmpLocalPath

End Function


'-------------------------------------------------------------------------------
' �e�X�g�R�[�h
' Test code for GetLocalPath
'-------------------------------------------------------------------------------

Private Sub Test_GetLocalPath_Function()
    Debug.Print "URL Path", ThisWorkbook.Path
    Debug.Print "Local Path", GetLocalPath(ThisWorkbook.Path)
End Sub

Private Sub Test_GetLocalPath_Speed()
    Dim i As Long, t1 As Single
    t1 = Timer
    For i = 1 To 100
        Call GetLocalPath(ThisWorkbook.Path, False)
    Next
    Debug.Print "UseCache Disable:"; Timer - t1; "[Sec]"
    t1 = Timer
    For i = 1 To 100
        Call GetLocalPath(ThisWorkbook.Path, True)
    Next
    Debug.Print "UseCache Enable: "; Timer - t1; "[Sec]"
End Sub

'-------------------------------------------------------------------------------
' ���̃��W���[���͂����ŏI���
' The script for this module ends here
'-------------------------------------------------------------------------------

