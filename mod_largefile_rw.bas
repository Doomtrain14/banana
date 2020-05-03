Attribute VB_Name = "mod_largefile_rw"
Option Explicit
Private Const GENERIC_WRITE            As Long = &H40000000
Private Const GENERIC_READ             As Long = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL    As Long = &H80&
Private Const CREATE_ALWAYS            As Long = 2
Private Const OPEN_ALWAYS              As Long = 4

Private Const FILE_BEGIN               As Long = 0
Private Const FILE_CURRENT             As Long = 1
Private Const FILE_END                 As Long = 2

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
          lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
          lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long
    
Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long
    
Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
          lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
          lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lDistanceToMove As Long, _
          lpDistanceToMoveHigh As Long, _
    ByVal dwMoveMethod As Long) As Long
    
Public Function file_open(ByVal file_name As String) As Long

    file_open = CreateFile(file_name, GENERIC_WRITE Or GENERIC_READ, 0, _
                       0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If file_open = -1 Then
        Call notify_error(1)
    End If
End Function

Public Sub file_close(ByRef file_handle As Long)
    If (file_handle <> -1) Then
        CloseHandle file_handle
    Else
        Call notify_error(4)
    End If
End Sub

Public Function read_data(ByRef file_handle As Long, ByRef data() As Byte) As Long

    Call ReadFile(file_handle, data(LBound(data)), UBound(data) - LBound(data) + 1, read_data, 0)

End Function

Public Function write_data(ByRef file_handle As Long, data() As Byte) As Long
    Call WriteFile(file_handle, data(LBound(data)), UBound(data) - LBound(data) + 1, write_data, 0)
End Function

Public Function file_seek_pos(ByRef file_handle As Long, ByVal pos As Currency) As Long
    Dim lo_val As Long
    Dim hi_val As Long
    hi_val = 0
    lo_val = pos
    file_seek_pos = SetFilePointer(file_handle, lo_val, hi_val, FILE_BEGIN)
    
    If file_seek_pos = -1 Then
        Call notify_error(2)
    End If
End Function

Public Function file_seek_end(ByRef file_handle As Long) As Long

    file_seek_end = SetFilePointer(file_handle, 0&, ByVal 0&, FILE_END)
            
    If file_seek_end = -1 Then
        Call notify_error(3)
    End If
End Function
Public Sub notify_error(err_num As Long)
    Dim message As String
    If (err_num = 1) Then
        message = "Could not open file specified!"
    ElseIf (err_num = 2) Then
        message = "Error seeking to specified position"
    ElseIf (err_num = 3) Then
        message = "Error seeking end of file"
    ElseIf (err_num = 4) Then
        message = "Error closing the file"
    Else
        message = "Unknown error has occured"
    End If
    
    Call MsgBox(message, vbExclamation + vbOKOnly, "Error")
End Sub
