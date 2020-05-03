Attribute VB_Name = "mod_global"
Option Explicit
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpfile_name As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSizeEx Lib "kernel32" (ByVal hFile As Long, lpFileSize As Currency) As Boolean
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Const GENERIC_READ     As Long = &H80000000
Private Const FILE_SHARE_READ  As Long = &H1
Private Const OPEN_EXISTING    As Long = 3
Private Const OF_EXIST         As Long = &H4000
Private Const OFS_MAXPATHNAME  As Long = 128
Private Const HFILE_ERROR      As Long = -1

Public Const BUFFER_SIZE       As Currency = 4096
Public Const ONE_TB            As Currency = 1099511627776@
Public Const ONE_GB            As Currency = 1073741824
Public Const FOUR_GB_FAT32     As Currency = 4294901760@
Public Const ONE_MB            As Currency = 1048576
Public Const ONE_KB            As Currency = 1024

Public Const DEFAULT_FILTER    As String = "All Files (*.*)" + vbNullChar + "*.*" + vbNullChar + "XCI Files (*.xci)" + vbNullChar + "*.xci" + vbNullChar + "NSP Files (*.nsp)" + vbNullChar + "*.nsp"


Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type OpenFileName
    lStructSize    As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
Public Function file_exists(ByVal m_file_name As String) As Boolean
 
    Dim lRetVal As Long
    Dim OfSt As OFSTRUCT
    
    lRetVal = OpenFile(m_file_name, OfSt, OF_EXIST)
    If lRetVal <> HFILE_ERROR Then
        file_exists = True
    Else
        file_exists = False
    End If
    
End Function

Public Function delete_file(m_file_name As String) As Boolean
    On Error GoTo ERRORWAY
    DeleteFile m_file_name
    delete_file = True
    Exit Function
ERRORWAY:
    delete_file = False
End Function
Public Function create_dir(dir_path As String) As Long
    Dim Security As SECURITY_ATTRIBUTES
    create_dir = CreateDirectory(dir_path, Security)
    
End Function

Public Function get_path_of_file(file_name As String) As String
    Dim posn As Integer
    posn = InStrRev(file_name, "\")
    If posn > 0 Then
        get_path_of_file = Left$(file_name, posn)
    Else
        get_path_of_file = ""
    End If
    
End Function

Public Function get_file_size(file_name As String) As Currency
    Dim hFile As Long, nSize As Currency
    'open the file
    hFile = CreateFile(file_name, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    'get the filesize
    GetFileSizeEx hFile, nSize
    'close the file
    CloseHandle hFile
    
    get_file_size = nSize * 10000
End Function

Public Function get_file_extention(m_file_name As String) As String
    Dim split_file_name() As String
    Dim max As Integer
    split_file_name = Split(m_file_name, "\")
    max = UBound(split_file_name)
    split_file_name = Split(split_file_name(max), ".")
    max = UBound(split_file_name)
    
    If (max = 0) Then
        get_file_extention = ""
    Else
        get_file_extention = split_file_name(max)
    End If
End Function

Public Function convert_from_bytes(size As Currency) As String()
    Dim conv_file_size As Single
    Dim conv_file_unit As String
    If (size >= ONE_TB) Then
        conv_file_size = size / ONE_TB
        conv_file_unit = "TB"
    ElseIf (size >= ONE_GB) Then
        conv_file_size = size / ONE_GB
        conv_file_unit = "GB"
    ElseIf (size >= ONE_MB) Then
        conv_file_size = size / ONE_MB
        conv_file_unit = "MB"
    ElseIf (size >= ONE_KB) Then
        conv_file_size = size / ONE_KB
        conv_file_unit = "KB"
    Else
        conv_file_size = size
        conv_file_unit = "B"
    End If
    
    conv_file_size = Round(conv_file_size, 2)
    Dim arr(2) As String
    arr(1) = Str$(conv_file_size)
    arr(2) = conv_file_unit
    convert_from_bytes = arr
End Function

Public Function convert_to_bytes(size As Currency, unit As String) As Currency
    If (unit = "KB") Then
        convert_to_bytes = size * ONE_KB
    ElseIf (unit = "MB") Then
        convert_to_bytes = size * ONE_MB
    ElseIf (unit = "GB") Then
        convert_to_bytes = size * ONE_GB
    ElseIf (unit = "TB") Then
        convert_to_bytes = size * ONE_TB
    Else
        convert_to_bytes = size
    End If
End Function

Public Function open_file(handle As Long, Optional start_path As String = "C:\", Optional filters As String = DEFAULT_FILTER) As String
    Dim OFName As OpenFileName
    OFName.lStructSize = Len(OFName)
    
    'Set the parent window
    OFName.hwndOwner = handle
    
    'Set the application's instance
    OFName.hInstance = App.hInstance
    
    'Select a filter
    OFName.lpstrFilter = filters
    
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    
    'set the maximum length of a returned file
    OFName.nMaxFile = 255
    
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(254)
    
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 255
    
    'Set the initial directory
    OFName.lpstrInitialDir = start_path

    'Set the title
    OFName.lpstrTitle = "Open File"
    
    'No flags
    OFName.flags = 0
    
    If (GetOpenFileName(OFName)) Then
        open_file = Replace(Trim$(OFName.lpstrFile), vbNullChar, "")
    Else
        open_file = ""
    End If
    
    
End Function
