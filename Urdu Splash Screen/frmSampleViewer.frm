VERSION 5.00
Begin VB.Form frmSampleViewer 
   BorderStyle     =   0  'None
   Caption         =   "Sample Multi-GIF Viewer"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBkg 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      Picture         =   "frmSampleViewer.frx":0000
      ScaleHeight     =   4815
      ScaleWidth      =   8025
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8025
      Begin VB.Timer Timer1 
         Interval        =   6000
         Left            =   240
         Top             =   240
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   6000
         Picture         =   "frmSampleViewer.frx":7E14
         Top             =   3600
         Width           =   1275
      End
      Begin VB.Shape shpMarker 
         BorderColor     =   &H80000005&
         Height          =   1095
         Index           =   0
         Left            =   6000
         Top             =   2280
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmSampleViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myGifViewer() As cGIFViewer ' array of gif viewers
Implements IGifRender               ' required when using the gif viewer


Private Sub Form_Load()

    Me.ScaleMode = vbPixels
    Me.picBkg.ScaleMode = vbPixels
    ReDim myGifViewer(shpMarker.LBound To shpMarker.UBound)


    Dim Slot As Long
    Dim Path As String
    Path = App.Path & "\Gif\Giffile.gif"


        For Slot = shpMarker.LBound To shpMarker.UBound
            ' find a free viewer to use
            If myGifViewer(Slot) Is Nothing Then Exit For
        Next
        
        Set myGifViewer(Slot) = New cGIFViewer
        If myGifViewer(Slot).LoadGIF(Path, Slot, Me, Me.hWnd) >= 1 Then

            myGifViewer(Slot).AnimationState = gfaPlaying
            
        End If


End Sub

Private Sub IGifRender_GetRenderDC(ByVal ViewerID As Long, ByVal FrameIndex As Long, _
        destDC As Long, Optional hwndRefresh As Long, Optional bAutoRedraw As Boolean = False, Optional PostNotify As Boolean = False)

destDC = picBkg.hdc
    
End Sub

Private Sub IGifRender_Rendered(ByVal ViewerID As Long, ByVal FrameIndex As Long, ByVal Message As RenderMessage, ByVal MsgValue As Long)
              
With shpMarker(ViewerID)
               
            If ViewerID = shpMarker.LBound Then
                    myGifViewer(ViewerID).SetAnimationBkg gfdBkgFromDC, picBkg.hdc, _
                    .Left + 1, .Top + 1, .Width - 2, .Height - 2, gfsCentered
            End If
                
End With


End Sub

Private Sub picBkg_Click()

End Sub

Private Sub Timer1_Timer()
frmSampleViewer.Hide
Form_Main.Show
End Sub
