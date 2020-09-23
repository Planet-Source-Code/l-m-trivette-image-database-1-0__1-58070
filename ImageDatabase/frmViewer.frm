VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmViewer 
   Caption         =   " Database Image Viewer"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   13
      Top             =   6225
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6668
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clipboard"
      Height          =   255
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Export"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   2175
      Begin VB.CheckBox Check2 
         Caption         =   "Delete after exporting"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Delete after importing"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Image Type Filter:"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Import"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Delete"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   2640
      ScaleHeight     =   7.84
      ScaleMode       =   0  'User
      ScaleWidth      =   7.268
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   6240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Stored Images"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   2175
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5295
      Left            =   2640
      MouseIcon       =   "frmViewer.frx":030A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   5760
      Width           =   5775
   End
   Begin VB.Label Label2 
      Caption         =   "Image Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   150
      Width           =   1935
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Written by L. "Mike" Trivette
' Please send me comments at mtrivette@yahoo.com
'
'
' Last Revised 1/02/05
'
' Sorry for any snippets i used to did nto give credit.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Public m_CRC As clsCRC
Public bCancel As Boolean
Dim srcDB As String
Dim ask As Boolean
Dim response As Integer
Const Blocksize = 32768


Private Sub loadtitles(Optional strsql As String)
    Dim dbs As Database
    Dim rst As Recordset
    Dim i As Long
    List1.Clear
    Image1.Picture = Nothing
    Label4.Caption = ""
    If strsql = "" Then
        strsql = "Select * from icons;"
        Combo1.Text = "All"
    End If
    
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset(strsql)
    If rst.RecordCount > 0 Then
        rst.MoveLast: rst.MoveFirst
        StatusBar1.Panels(1).Text = rst.RecordCount & " images found"
        For i = 1 To rst.RecordCount
            List1.AddItem rst.Fields("title")
            rst.MoveNext
        Next i
    End If
    rst.Close
    dbs.Close
End Sub

Private Sub Combo1_Click()
    Dim strtemp As String
    Dim temp As String
    
    temp = Combo1.Text
    If Combo1.Text = "All" Then
        strtemp = ""
    Else
        strtemp = "SELECT * FROM icons WHERE type = '" & Combo1.Text & "';"
    End If
    
    loadtitles strtemp
End Sub

Private Sub Command1_Click()
    Dim i As Long
    Dim X As String
    Dim lastindex As Long
    
    ' Exit sub if no images chosen
    If List1.SelCount = 0 Then Exit Sub
    
    ' Process chosen images
    If List1.SelCount = 1 Then
        ' If only one selection is made
        DelImage List1.Text
        lastindex = List1.ListIndex
        List1.RemoveItem lastindex
        Image1.Picture = Nothing
        If lastindex >= List1.ListCount Then lastindex = 0
        If lastindex <> 0 Then
            List1.ListIndex = lastindex
            List1.Selected(lastindex) = True
        End If
        loadtypes
    Else
        ' Cycle through each selection if multiple items chosen
        ProgressBar1.Max = List1.SelCount
        For i = 0 To List1.ListCount - 1
            If List1.Selected(i) = True Then
                DelImage List1.List(i)
                ProgressBar1.Value = ProgressBar1.Value + 1
            End If
        Next i
        ProgressBar1.Value = 0
        loadtitles
        loadtypes
        If List1.ListCount > 0 Then List1.Selected(0) = True
    End If
    

End Sub

Private Sub Command3_Click()
    Dim i As Long
    Dim X As String
    
    ' Exit sub if no images chosen
    If List1.SelCount = 0 Then Exit Sub
    
    ' Get folder to export images into
    X = GetFolder("Choose a directory to export images to.") & "\"
    If X = "" Then Exit Sub
    
    ' Cycle through chosen images
    ProgressBar1.Max = List1.SelCount
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) = True Then
            ProgressBar1.Value = ProgressBar1.Value + 1
            WriteImage List1.List(i), X
            StatusBar1.Panels(2).Text = List1.List(i) & " exported"
            If Check2.Value = 1 Then DelImage List1.List(i)
        End If
        DoEvents
    Next i
    'StatusBar1.Panels(2).Text = ""
    ProgressBar1.Value = 0
End Sub

Private Sub Command4_Click()
    ' Make sure there any images shown
    If List1.ListCount = 0 Then Exit Sub
    ' Make sure image type is clipboard compatible
    If Image1.Picture.Type <> 1 Then
        response = MsgBox("Cannot copy selected image to clipboard." & vbCrLf & vbCrLf & "Wrong image type.", vbOKOnly + vbExclamation, " Clipboard Error")
        Exit Sub
    End If
    ' Set image to clipboard
    Clipboard.SetData Image1.Picture
    ' Update the status bar to notify user
    StatusBar1.Panels(2).Text = List1.Text & " copied to clipboard"
End Sub

Private Sub Form_Load()
    srcDB = App.path & "\images.mdb"
    loadtitles
    loadtypes
    AddProgBar ProgressBar1, StatusBar1, 3
    Me.Caption = Me.Caption & " (Beta version " & App.Major & "." & App.Minor & "." & App.Revision & ")"
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    Me.Height = 6930
    Me.Width = 8655
End Sub

Private Sub Image1_DblClick()
    frmPreview.Show
    frmPreview.Picture1 = Image1.Picture
    frmPreview.Caption = " Preview - [" & Label4.Caption & "]"
End Sub

Private Sub List1_Click()
    On Error Resume Next
    Dim dbs As Database
    Dim rst As Recordset
    Dim strFile As String
    
    StatusBar1.Panels(2).Text = ""
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("Select * from icons where title = '" & List1.Text & "';")
    Image1.Picture = Nothing
    strFile = "_" & rst.Fields("title")
    Call WriteBLOB(rst, "BinData", strFile)
    StatusBar1.Panels(2).Text = "Image: " & rst.Fields("width") & " x " & rst.Fields("height") & " - " & SetBytes(rst.Fields("size"))
    Label4.Caption = rst.Fields("title")
    Image1.Picture = LoadPicture(strFile)
    rst.Close
    dbs.Close
    Kill strFile ' Remove temp file
End Sub

Private Sub List1_DblClick()
    frmPreview.Show
    frmPreview.Picture1 = Image1.Picture
    frmPreview.Caption = " Preview - [" & Label4.Caption & "]"
End Sub

'**********************************************************************
'FUNCTION: WriteBLOB()
'
'PURPOSE:
'WritesBLOB information stored in the specified table and field to the
'specified disk file.
'
'PREREQUISITES:
'
'ARGUMENTS:
'Destination - the path and filename of the file to be extracted.
'T - the table object the data is stored in.
'Field - the OLE object to store the data in.
'
'RETURN:
'0 on fail 1 on success
'**********************************************************************

Public Function WriteBLOB(T As Recordset, sField As String, Destination As String)
        On Error GoTo Err_WriteBLOB
        Dim NumBlocks As Integer, DestFile As Integer, i As Integer
        Dim FileLength As Long, LeftOver As Long
        Dim FileData() As Byte, retval As Variant

        ' Get the length of the file.
        FileLength = T(sField).FieldSize()
        If FileLength <> 0 Then
            DestFile = FreeFile
            NumBlocks = FileLength \ Blocksize
            LeftOver = FileLength Mod Blocksize 'reminder appended first
            'initialize status bar meter
            'RetVal = SysCmd(acSysCmdInitMeter, "Writing BLOB", NumBlocks)

            Open Destination For Binary Access Write Lock Write As DestFile
            ReDim FileData(LeftOver)
            FileData() = T(sField).GetChunk(0, LeftOver)
            Put DestFile, , FileData() 'write first chunk
            
            ReDim FileData(Blocksize)
            
            For i = 1 To NumBlocks
                FileData() = T(sField).GetChunk((i - 1) * Blocksize _
                + LeftOver, Blocksize)
                Put DestFile, , FileData() 'write remaining chunks
                'update status bar meter
                'RetVal = SysCmd(acSysCmdUpdateMeter, i)
            Next i
            Close DestFile
        End If
        
        'remove status bar meter
        'RetVal = SysCmd(acSysCmdRemoveMeter)
        WriteBLOB = 1
        Exit Function

Err_WriteBLOB:
        MsgBox Err.Description
        WriteBLOB = 0
        Exit Function
End Function

Private Sub loadtypes()
    Dim dbs As Database
    Dim rst As Recordset
    Dim i As Long
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("Select type from icons group by type;")
    Combo1.Clear
    
    If rst.RecordCount > 0 Then
    rst.MoveLast: rst.MoveFirst
    Combo1.AddItem "All"
    For i = 1 To rst.RecordCount
        Combo1.AddItem rst.Fields("type")
        rst.MoveNext
    Next i
    End If
    
    rst.Close
    dbs.Close
End Sub

Private Sub Command2_Click()
    Dim BufferFileArray() As String
    Dim i As Integer
    
    ask = True
    
    With CommonDialog1
        .DialogTitle = "Add Multiple files..."
        .Filter = "All Image Files|*.jpg;*.jpeg;*.gif;*.bmp;*.ico;*.wmf"
        .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
        .InitDir = CurDir
        .MaxFileSize = 32767
        .Filename = ""
        .ShowOpen
        BufferFileArray = Split(.Filename, Chr(0))
    End With
    
    ' If no files are selected
    If UBound(BufferFileArray) = -1 Then Exit Sub
    
    ' If only one file was chosen.
    If UBound(BufferFileArray) = 0 Then
        saveimage CommonDialog1.Filename
        Exit Sub
    End If
    
    ' If multiple files chosen.
    ProgressBar1.Max = UBound(BufferFileArray)
    For i = LBound(BufferFileArray) + 1 To UBound(BufferFileArray)
        ProgressBar1.Value = i
        saveimage CurDir & "\" & BufferFileArray(i)
    Next i
    ProgressBar1.Value = 0
    List1.Selected(List1.ListCount - 1) = True
End Sub

Private Sub saveimage(strImage As String)
    ' Save image to database
    On Error Resume Next
    
    Dim NumBlocks As Integer, SourceFile As Integer, i As Integer
    Dim FileLength As Long, LeftOver As Long
    Dim FileData() As Byte, retval As Variant
    Dim dbs As Database
    Dim rst As Recordset
    Dim strHex As String
    
    Set m_CRC = New clsCRC
    
    Picture1.Cls
    Picture1.Picture = LoadPicture(strImage)
    
    StatusBar1.Panels(2).Text = "Importing " & GetFileName(Replace(strImage, "'", ""))
    
    strHex = Hex(m_CRC.CalculateFile(strImage))
    
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("SELECT * FROM icons where crc = '" & strHex & "';")
           
    m_CRC.Algorithm = CRC32
    
    If rst.RecordCount = 0 Then
    rst.AddNew
        rst.Fields("title") = GetFileName(Replace(strImage, "'", ""))
        rst.Fields("crc") = strHex
        rst.Fields("size") = FileLen(strImage)
        rst.Fields("width") = Picture1.Width / 15
        rst.Fields("height") = Picture1.Height / 15
        rst.Fields("type") = LCase(GetFileExtension(strImage))
        
        SourceFile = FreeFile
        Open strImage For Binary Access Read As SourceFile
        FileLength = LOF(SourceFile)
            NumBlocks = FileLength \ Blocksize
            LeftOver = FileLength Mod Blocksize 'remainder appended first
            ReDim FileData(LeftOver)
            Get SourceFile, , FileData()
            rst.Fields("BinData").AppendChunk FileData() 'store the first image chunk
            ReDim FileData(Blocksize)
            For i = 1 To NumBlocks
                Get SourceFile, , FileData()
                rst.Fields("BinData").AppendChunk FileData() 'remaining chunks
                DoEvents
            Next i
        Close SourceFile
        rst.Update
        List1.AddItem GetFileName(Replace(strImage, "'", ""))
        List1.ListIndex = List1.ListCount - 1
    Else
        ' duplicate image found
        If ask = True Then response = MsgBox("This image was already found in database." & vbCrLf & vbCrLf & "Source: " & GetFileName(Replace(strImage, "'", "")) & vbCrLf & "Found: " & rst.Fields("title") & vbCrLf & vbCrLf & "Would you like to continue to receive duplicate warnings?", vbYesNo + vbInformation, "Duplicate")
        If response = 7 Or response = 0 Then
            ask = False
        Else
            ask = True
        End If
    End If
    
    rst.Close
    dbs.Close
    ' Delete the source file if user wants
    If Check1.Value = 1 Then Kill strImage
    StatusBar1.Panels(2).Text = ""
    
    loadtypes
End Sub

Private Sub DelImage(strImage As String)
    Dim dbs As Database
    Dim rst As Recordset
    Dim strTitle As String

    StatusBar1.Panels(2).Text = "Deleting " & strImage
    
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("Select * from icons where title = '" & strImage & "';")
    ' Delete image if found
    If rst.RecordCount = 0 Then
        MsgBox "Error deleting " & strImage & " from database."
    Else
        rst.Delete
    End If
    rst.Close
    dbs.Close
    
    StatusBar1.Panels(2).Text = ""
    Label4.Caption = ""
End Sub

Private Sub WriteImage(strImage As String, Optional strfolder As String)
    ' Get image from database
    Dim dbs As Database
    Dim rst As Recordset
    
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("SELECT * FROM icons where title = '" & strImage & "';")
    
    ' Write image to disk is found
    If rst.RecordCount > 0 Then
        Call WriteBLOB(rst, "BinData", strfolder & rst.Fields("title"))
        Picture1.Picture = LoadPicture(strfolder & rst.Fields("title"))
        Picture1.Refresh
    Else
        response = MsgBox("Image not found in database.", vbOKOnly + vbExclamation, "Error")
    End If
    
    rst.Close
    dbs.Close
    
    Set rst = Nothing
    Set dbs = Nothing
End Sub

