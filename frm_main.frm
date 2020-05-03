VERSION 5.00
Begin VB.Form frm_main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Splitter"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pb_bg 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7425
      ScaleWidth      =   12825
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.CommandButton cmd_split 
         Caption         =   "&SPLIT"
         Height          =   690
         Left            =   90
         TabIndex        =   15
         Top             =   2565
         Width           =   5415
      End
      Begin VB.Frame frame_info 
         Caption         =   "File Information"
         Height          =   1500
         Left            =   90
         TabIndex        =   9
         Top             =   990
         Width           =   2850
         Begin VB.Label lbl_split 
            AutoSize        =   -1  'True
            Caption         =   "File will be split in: "
            Height          =   195
            Left            =   225
            TabIndex        =   18
            Top             =   990
            Width           =   1320
         End
         Begin VB.Label lbl_part_count 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0 Part(s)"
            Height          =   195
            Left            =   1575
            TabIndex        =   17
            Top             =   990
            Width           =   945
         End
         Begin VB.Label label_file_size 
            Alignment       =   1  'Right Justify
            Caption         =   "<No file selected>"
            Height          =   195
            Left            =   945
            TabIndex        =   13
            Top             =   675
            Width           =   1575
         End
         Begin VB.Label label_file_type 
            Alignment       =   1  'Right Justify
            Caption         =   "<No file selected>"
            Height          =   195
            Left            =   1035
            TabIndex        =   12
            Top             =   360
            Width           =   1485
         End
         Begin VB.Label label_size 
            AutoSize        =   -1  'True
            Caption         =   "File Size:"
            Height          =   195
            Left            =   225
            TabIndex        =   11
            Top             =   675
            Width           =   630
         End
         Begin VB.Label label_type 
            AutoSize        =   -1  'True
            Caption         =   "File Type:"
            Height          =   195
            Left            =   225
            TabIndex        =   10
            Top             =   360
            Width           =   705
         End
      End
      Begin VB.Frame frame_options 
         Caption         =   "Settings"
         Height          =   1500
         Left            =   3105
         TabIndex        =   5
         Top             =   990
         Width           =   2400
         Begin VB.CheckBox chk_fourgb 
            Caption         =   "Split in 4GB for FAT32"
            Enabled         =   0   'False
            Height          =   240
            Left            =   180
            TabIndex        =   14
            Top             =   1080
            Width           =   1860
         End
         Begin VB.ComboBox cmb_unit 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frm_main.frx":0000
            Left            =   1530
            List            =   "frm_main.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   585
            Width           =   690
         End
         Begin VB.TextBox txt_split_size 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   330
            Left            =   180
            TabIndex        =   7
            Text            =   "1"
            Top             =   585
            Width           =   1320
         End
         Begin VB.Label label 
            AutoSize        =   -1  'True
            Caption         =   "Size: "
            Height          =   195
            Left            =   225
            TabIndex        =   6
            Top             =   315
            Width           =   390
         End
      End
      Begin VB.Frame frame_main 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   75
         TabIndex        =   1
         Top             =   75
         Width           =   5430
         Begin VB.CommandButton cmd_browse 
            Caption         =   "..."
            Height          =   330
            Left            =   4770
            TabIndex        =   4
            Top             =   315
            Width           =   465
         End
         Begin VB.TextBox txt_file_name 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   945
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "C:\"
            Top             =   315
            Width           =   3795
         End
         Begin VB.Label lbl_file 
            AutoSize        =   -1  'True
            Caption         =   "Filename:"
            Height          =   195
            Left            =   135
            TabIndex        =   2
            Top             =   360
            Width           =   690
         End
      End
      Begin VB.Label lbl_by 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "-TheYeti"
         Height          =   195
         Left            =   2970
         TabIndex        =   19
         Top             =   3375
         Width           =   2445
      End
      Begin VB.Label lbl_file_progress 
         AutoSize        =   -1  'True
         Caption         =   "File Progress: "
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   3375
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const APP_VERSION = "1.0"

Dim g_file_name As String
Dim g_file_size As Currency
Dim g_file_type As String
Dim g_split_size As Currency
Dim g_split_unit As String
Private Function update_part_count() As Currency
    g_split_size = txt_split_size.Text
    g_split_unit = cmb_unit.Text
    
    'Convert split size settings Ex: 1KB to 1024 bytes
    Dim conv_split_size As Currency
    conv_split_size = convert_to_bytes(Int(g_split_size), g_split_unit)
    
    '4GB for FAT32 actually less than 4GB
    If (chk_fourgb.Value = vbChecked) Then
        conv_split_size = FOUR_GB_FAT32
    End If

    'Get how many parts will be created
    Dim file_part_count As Currency
    file_part_count = 1 + Int(g_file_size / conv_split_size)
    
    'update part count in gui
    lbl_part_count.Caption = Str$(file_part_count) + " Part(s)"
    
    If (file_part_count < 2) Then
        cmd_split.Enabled = False
    Else
        cmd_split.Enabled = True
    End If
    
    update_part_count = file_part_count
End Function
Public Function large_modulo(num As Currency, div As Currency) As Currency
    large_modulo = num - div * Int(num / div)
End Function
Private Sub chk_fourgb_Click()
    If (chk_fourgb.Value = vbChecked) Then
        txt_split_size.Text = "4"
        cmb_unit.Text = "GB"
        txt_split_size.Enabled = False
        cmb_unit.Enabled = False
    Else
        txt_split_size.Enabled = True
        cmb_unit.Enabled = True
        txt_split_size.Text = "1"
        cmb_unit.Text = "B"
    End If
End Sub

Private Sub cmb_unit_Click()
    Call update_part_count
End Sub

Private Sub cmd_browse_Click()
    Dim file_name As String
    Dim file_ext As String
    Dim file_size As Currency
    
    'open file dialog
    file_name = open_file(Me.hWnd, txt_file_name.Text)

    'if valid file name
    If (file_name <> "") Then
        'update file name textbox
        txt_file_name.Text = file_name
        
        'get file size
        file_size = get_file_size(file_name)
        
        'convert it to non-byte(s) format if possible ex: 109 MB
        '109 will be at index 1
        ' MB will be at index 2
        Dim conv_file_size() As String
        conv_file_size = convert_from_bytes(file_size)
        
        'Update the file size in file information
        label_file_size.Caption = conv_file_size(1) + " " + conv_file_size(2)
                
        
        'Get file extension
        file_ext = get_file_extention(file_name)
        If (file_ext = "") Then
            label_file_type.Caption = "No file extension"
        Else
            label_file_type.Caption = LCase$(file_ext) + " file"
        End If
        
        'Enable controls
        If (conv_file_size(2) = "GB" And conv_file_size(1) >= 4) Then
            chk_fourgb.Enabled = True
        Else
            chk_fourgb.Enabled = False
        End If
        
        'Enable and set initial split size
        txt_split_size.Enabled = True
        cmb_unit.Enabled = True
        txt_split_size.Text = Int(conv_file_size(1))
        cmb_unit.Text = conv_file_size(2)
        
        'Update global variables
        g_file_name = file_name
        g_file_type = file_ext
        g_file_size = file_size
        g_split_size = txt_split_size.Text
        g_split_unit = cmb_unit.Text
        
        Call update_part_count
        
    End If
End Sub

Private Sub cmd_split_Click()
    Dim Ret As Long
    g_split_size = txt_split_size.Text
    g_split_unit = cmb_unit.Text
    
    'Convert split size settings Ex: 1KB to 1024 bytes
    Dim conv_split_size As Currency
    conv_split_size = convert_to_bytes(Int(g_split_size), g_split_unit)
    
    '4GB for FAT32 actually less than 4GB
    If (chk_fourgb.Value = vbChecked) Then
        conv_split_size = FOUR_GB_FAT32
    End If
    
    'Get how many parts will be created
    Dim file_part_count As Currency
    file_part_count = 1 + Int(g_file_size / conv_split_size)
    
    'update part count in gui
    lbl_part_count.Caption = Str$(file_part_count) + " Part(s)"
    
    If (file_part_count < 2) Then
        cmd_split.Enabled = False
    Else
        cmd_split.Enabled = True
    End If
    
    If (file_part_count < 2) Then
        Call MsgBox("Selected split file size is greater than or equal to original file size", vbCritical + vbOKOnly)
        txt_split_size.SelStart = 0
        txt_split_size.SelLength = Len(txt_split_size.Text)
        Exit Sub
    End If
    
    
    
    'set working dir to path of file
    Dim working_dir As String
    working_dir = get_path_of_file(g_file_name)
    
    'get filename without path
    Dim no_path_file_name As String
    no_path_file_name = Replace(g_file_name, working_dir, "")
    
    'get file name without path and file extension
    Dim no_path_ext_file_name As String
    no_path_ext_file_name = no_path_file_name
    
    If (g_file_type <> "") Then
        no_path_ext_file_name = Mid$(no_path_file_name, 1, Len(no_path_file_name) - Len(g_file_type) - 1)
    End If

    'set split folder path
    Dim split_folder_path As String
    split_folder_path = working_dir + no_path_ext_file_name + " [SPLIT]"
    
    'add .nsp to folder name for nsp files
    If (LCase$(g_file_type) = "nsp") Then
        split_folder_path = split_folder_path + ".nsp"
    End If
    
    If Dir(split_folder_path, vbDirectory) <> "" Then
        If (MsgBox("Output folder already exists" + vbNewLine + "File(s) inside can be overwritten by the process" + vbNewLine + vbNewLine + "Do you want to proceed?", vbExclamation + vbYesNo) = vbNo) Then
            Exit Sub
        End If
    Else
        If (create_dir(split_folder_path) = 0) Then
            Call MsgBox("Error : Couldn't create directory !", vbCritical + vbOKOnly)
            Exit Sub
        End If
    End If
    

    
    Dim split_file_names() As String
    ReDim split_file_names(file_part_count - 1)
    
    Dim i As Integer
    If (LCase$(g_file_type) = "nsp") Then
        For i = 0 To file_part_count - 1
            split_file_names(i) = Format(i, "00")
        Next
    ElseIf (LCase$(g_file_type) = "xci") Then
        For i = 0 To file_part_count - 1
            split_file_names(i) = no_path_ext_file_name + ".xc" + Str$(i)
        Next
    Else
        For i = 0 To file_part_count - 1
            split_file_names(i) = no_path_ext_file_name + "." + Format(i, "00")
        Next
    End If
    
    'Disable split button
    cmd_split.Enabled = False
    'On Error GoTo EXITPATH

    Dim data() As Byte
    ReDim data(BUFFER_SIZE - 1)
    
    Dim fh_out As Long
    Dim fh_in As Long
    
    fh_in = file_open(g_file_name)
    
    For i = 0 To UBound(split_file_names)
    
        Dim output_file As String
        output_file = split_folder_path + "\" + split_file_names(i)
        If (file_exists(output_file) = True) Then
            If delete_file(output_file) = False Then
                Call MsgBox("Unable to delete existing files", vbExclamation + vbOKOnly, "Error")

                file_close (fh_out)
                file_close (fh_in)

                Exit Sub
            End If
        End If
        

        fh_out = file_open(output_file)
        
        Dim read_ctr As Currency
        Dim pos As Currency


        If (i = UBound(split_file_names)) Then
            read_ctr = Int(large_modulo(g_file_size, conv_split_size) / BUFFER_SIZE)
        Else
            read_ctr = Int(conv_split_size / BUFFER_SIZE)
        End If
        
        
        Dim update_freq As Currency
        update_freq = Int((read_ctr - 1) * 0.05)
        
        pos = 0
        While (pos < read_ctr)
            Call file_seek_pos(fh_in, (conv_split_size * i) + (pos * BUFFER_SIZE))
            Call read_data(fh_in, data)
            
            Call file_seek_end(fh_out)
            Call write_data(fh_out, data)
            
            If (large_modulo(pos, update_freq) < 1) Then
                lbl_file_progress.Caption = "File Progress: (" + Trim$(i + 1) + "/" + Trim$(1 + UBound(split_file_names)) + ") " + Format$(Round(100 * pos / (read_ctr - 1), 2), "00.00") + "%"
                DoEvents
            End If
            pos = pos + 1
        Wend
        
        If (i = UBound(split_file_names)) Then
            Dim new_byte_size As Currency
            Dim mod_file_size As Currency
            mod_file_size = large_modulo(g_file_size, conv_split_size)
            new_byte_size = large_modulo(mod_file_size, BUFFER_SIZE)
            ReDim data(new_byte_size - 1)
            
            Call file_seek_pos(fh_in, (conv_split_size * i) + (read_ctr * BUFFER_SIZE))
            Call read_data(fh_in, data)

            Call file_seek_end(fh_out)
            Call write_data(fh_out, data)
        End If


        file_close (fh_out)
    Next i
    file_close (fh_in)

    file_close (fh_out)

    
    If (LCase$(g_file_type) = "nsp") Then
        SetAttr split_folder_path, vbArchive
    End If
    
    lbl_file_progress.Caption = "File Progress: Done!"
    Call MsgBox("Split operation complete!", vbInformation + vbOKOnly, "Complete")
    lbl_file_progress.Caption = "File Progress: "
    
    'Enable split button
    cmd_split.Enabled = True
    Exit Sub
EXITPATH:
    'If (bin_file_out.IsOpen = True) Then: bin_file_out.CloseFile
    file_close (fh_out)
    file_close (fh_in)
    
    Call MsgBox("Split operation aborted due to error!", vbExclamation + vbOKOnly, "Error")
    'Enable split button
    cmd_split.Enabled = True
End Sub

Private Sub Form_Load()
    frm_main.Caption = "Banana v" + APP_VERSION
    label_file_size.Caption = "<No file selected>"
    label_file_type.Caption = "<No file selected>"
End Sub

Private Sub txt_split_size_Change()
    Call update_part_count
End Sub

Private Sub txt_split_size_GotFocus()
    txt_split_size.SelStart = 0
    txt_split_size.SelLength = Len(txt_split_size.Text)
End Sub

Private Sub txt_split_size_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 8) Then
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt_split_size_LostFocus()
    If (txt_split_size.Text = "" Or txt_split_size.Text = "0") Then
        txt_split_size.Text = 1
    End If
    
    
End Sub

