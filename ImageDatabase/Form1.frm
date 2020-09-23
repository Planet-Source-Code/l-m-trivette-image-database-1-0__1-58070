VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   4080
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "     Show Preview"
      Height          =   4695
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   5535
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   0
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   120
         ScaleHeight     =   6.419
         ScaleMode       =   0  'User
         ScaleWidth      =   6.664
         TabIndex        =   10
         Top             =   240
         Width           =   5295
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4965
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5556
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5556
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5556
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Import"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   4575
      Left            =   2160
      MultiSelect     =   2  'Extended
      Pattern         =   "*.ico;*.jpg;*.gif;*.jpeg"
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "c:\icon.ico"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'modBlob reads and writes binary data to and from an OLE field in a table
'Source: http://support.microsoft.com/default.aspx?scid=KB;EN-US;Q103257&

''''Option Compare Database
Option Explicit

Public m_CRC As clsCRC
Public bCancel As Boolean
Const Blocksize = 32768
Const srcDB = "D:\VB Code\Icon Projects\IconDB\images.mdb"


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

Private Function WriteBLOB(T As Recordset, sField As String, Destination As String)
        
        Dim NumBlocks As Integer, DestFile As Integer, i As Integer
        Dim FileLength As Long, LeftOver As Long
        Dim FileData() As Byte, retval As Variant

        ' On Error GoTo Err_WriteBLOB

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

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Picture1.Visible = True
    Else
        Picture1.Visible = False
    End If
End Sub

Private Sub Command1_Click()
    saveimage txtPath.Text
    'MsgBox txtPath.Text
End Sub

Private Sub saveimage(strImage As String)
    ' Save image to database
    Dim NumBlocks As Integer, SourceFile As Integer, i As Integer
    Dim FileLength As Long, LeftOver As Long
    Dim FileData() As Byte, retval As Variant
    
    Dim dbs As Database
    Dim rst As Recordset
    Dim strHex As String
    Dim response As String
    
    'Strip apostraphes from title string
    
    strHex = Hex(m_CRC.CalculateFile(strImage))
    
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("SELECT * FROM icons where crc = '" & strHex & "';")
           
    m_CRC.Algorithm = CRC32
    
    If rst.RecordCount = 0 Then
    rst.AddNew
        StatusBar1.Panels(3).Text = Picture1.Width / 15 & " x " & Picture1.Height / 15
        'Call ReadBLOB(strFile, rst, "BinData")
        rst.Fields("title") = GetFileName(Replace(strImage, "'", ""))
        rst.Fields("crc") = strHex
        rst.Fields("size") = GetFileSize(strImage)
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
    Else
        StatusBar1.Panels(3).Text = "Duplicate!"
    End If
    
    rst.Close
    dbs.Close
    
    On Error Resume Next
    Picture1.Cls
    Picture1.Picture = LoadPicture(strImage)
End Sub

Private Sub Command2_Click()
Dim i As Long
bCancel = False

For i = 0 To File1.ListCount - 1
    StatusBar1.Panels(1).Text = "Importing " & i & " of " & ProgressBar1.Max
    saveimage Dir1.path & "\" & File1.List(i)
    ProgressBar1.Value = i
    If bCancel = True Then
        ProgressBar1.Value = 0
        Exit Sub
    End If
Next i

StatusBar1.Panels(1).Text = ""
ProgressBar1.Value = 0

For i = 0 To File1.ListCount - 1
    StatusBar1.Panels(1).Text = "Deleting " & i & " of " & ProgressBar1.Max
    'Kill Dir1.path & "\" & File1.List(i)
    ProgressBar1.Value = i
Next i

End Sub

Private Sub Command3_Click()
    ' Get image from database
    ' Right now this sub dumps out all the images out
    ' of the database.
    '
    ' Not pretty, but it works....
    
    Dim dbs As Database
    Dim rst As Recordset
    Dim i As Long
    
    Set dbs = OpenDatabase(srcDB)
    Set rst = dbs.OpenRecordset("SELECT * FROM icons;")
    
    rst.MoveLast: rst.MoveFirst
    ProgressBar1.Max = rst.RecordCount
    For i = 1 To rst.RecordCount
        ProgressBar1.Value = i
        StatusBar1.Panels(1).Text = "Writing " & rst.Fields("title")
        Call WriteBLOB(rst, "BinData", rst.Fields("title"))
        Picture1.Picture = LoadPicture(rst.Fields("title"))
        Picture1.Refresh
    rst.MoveNext
    Next i
    
    rst.Close
    dbs.Close
    
    Set rst = Nothing
    Set dbs = Nothing
    
    ProgressBar1.Value = 0
End Sub

Private Sub Command4_Click()
    bCancel = True
End Sub

Private Sub Dir1_Change()
    File1.path = Dir1.path
End Sub

Private Sub Drive1_Change()
    Dir1.path = Drive1.Drive
End Sub

Private Sub File1_Click()
    If Check1.Value = 1 Then
        Picture1.Picture = LoadPicture(Dir1.path & "\" & File1.FileName)
        StatusBar1.Panels(3).Text = Picture1.Width / 15 & " x " & Picture1.Height / 15
        txtPath.Text = Dir1.path & "\" & File1.FileName
    End If
End Sub

Private Sub File1_PathChange()
    If File1.ListCount > 1 Then
        Picture1.Visible = True
        ProgressBar1.Max = File1.ListCount - 1
        StatusBar1.Panels(2).Text = File1.ListCount & " items listed."
    End If
End Sub

Private Sub Form_Load()
    'File1.Path = Dir1.Path
    'File1.Refresh
    Set m_CRC = New clsCRC
End Sub

Function GetFileName(path As String) As String
    Dim i As Integer
    For i = (Len(path)) To 1 Step -1
        If Mid(path, i, 1) = "\" Then
            GetFileName = Mid(path, i + 1, Len(path) - i + 1)
            Exit For
        End If
    Next
End Function

Public Function GetFileSize(FileName) As String
    On Error GoTo Gfserror
    Dim TempStr As String
    TempStr = FileLen(FileName)


    If TempStr >= "1024" Then
        'KB
        TempStr = CCur(TempStr / 1024) & "KB"
    Else


        If TempStr >= "1048576" Then
            'MB
            TempStr = CCur(TempStr / (1024 * 1024)) & "KB"
        Else
            TempStr = CCur(TempStr) & "B"
        End If
    End If
    GetFileSize = TempStr
    Exit Function
Gfserror:
    GetFileSize = "0B"
    Resume
End Function

Public Function GetFileExtension(FileName As String)
    On Error Resume Next
    Dim TempStr As String
    TempStr = Right(FileName, 2)


    If Left(TempStr, 1) = "." Then
        GetFileExtension = Right(FileName, 1)
        Exit Function
    Else
        TempStr = Right(FileName, 3)


        If Left(TempStr, 1) = "." Then
            GetFileExtension = Right(FileName, 2)
            Exit Function
        Else
            TempStr = Right(FileName, 4)


            If Left(TempStr, 1) = "." Then
                GetFileExtension = Right(FileName, 3)
                Exit Function
            Else
                TempStr = Right(FileName, 5)


                If Left(TempStr, 1) = "." Then
                    GetFileExtension = Right(FileName, 4)
                    Exit Function
                Else
                    GetFileExtension = "Unknown"
                End If
            End If
        End If
    End If
    
End Function


