Attribute VB_Name = "modCommonDialog"
'''''''''''''''''''''''''''''''''''''''''''
' Windows API/Global Declarations for :
' Browse Folder Dialog
'
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const BIF_BROWSEFORCOMPUTER = &H1000

Private Const MAX_PATH = 260


Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As BrowseInfo) As Long


Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long


Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long


Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Function GetFolder(Optional title As String) As String
'Opens a Treeview control that displays
    '     the directories in a computer
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    szTitle = title

    With tBrowseInfo
        .hwndOwner = frmViewer.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
        
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        GetFolder = sBuffer
    End If


End Function

Public Function AddProgBar(pb As ProgressBar, sb As StatusBar, lPan As Long)


' make sure that when the form is resized that the
' statusbar is rsized before we continue
sb.Align = 2
sb.Refresh

' set the properties of the progressbar
' flat with no border seems to look the best
' also set the progressbar to the top of the zorder
pb.ZOrder 0
pb.Appearance = ccFlat
pb.BorderStyle = ccNone

' now resize the progressbar1 to fit in the statusbar panel
pb.Left = sb.Panels(lPan).Left + 25
pb.Width = sb.Panels(lPan).Width - 45
pb.Top = sb.Top + 45
pb.Height = sb.Height - 75


End Function



