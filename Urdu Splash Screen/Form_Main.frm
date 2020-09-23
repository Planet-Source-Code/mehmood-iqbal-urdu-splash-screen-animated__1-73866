VERSION 5.00
Begin VB.Form Form_Main 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   1440
      Picture         =   "Form_Main.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   1320
      Picture         =   "Form_Main.frx":0167
      Top             =   720
      Width           =   2130
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload frmSampleViewer
Unload Me
End Sub
