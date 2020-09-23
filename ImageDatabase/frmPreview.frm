VERSION 5.00
Begin VB.Form frmPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Preview"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5895
      Left            =   0
      ScaleHeight     =   5835
      ScaleWidth      =   6555
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Picture1_Change()
    ' Size form to fit picture
    Me.Width = Picture1.Width + 70
    Me.Height = Picture1.Height + 380
    ' Postion form so that it will be centered
    Me.Top = (Screen.Height * 0.85) / 2 - Me.Height / 2 + 350
    Me.Left = Screen.Width / 2 - Me.Width / 2
    ' Make sure form top is not off screen
    If Me.Top < 50 Then Me.Top = 50
End Sub

