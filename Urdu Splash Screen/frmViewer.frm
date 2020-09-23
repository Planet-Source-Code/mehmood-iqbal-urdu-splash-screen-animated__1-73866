VERSION 5.00
Begin VB.Form frmViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading GIF"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private myGifViewer As cGIFViewer
Implements IGifRender

Private Sub Form_Load()
    Me.ScaleMode = vbPixels
End Sub

Private Sub Form_Paint()
    If Not myGifViewer Is Nothing Then
        If Not myGifViewer.AnimationState = gfaPlaying Then myGifViewer.AnimationState = gfaRefresh
    End If
End Sub

Public Sub ViewGIF(ByVal FileName As String, ByVal bkgColor As Long)
    If Me.Visible = False Then
        Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
    End If
    Set myGifViewer = New cGIFViewer
    Me.BackColor = bkgColor
    Me.Cls
    Me.Show
    DoEvents
    If myGifViewer.LoadGIF(FileName, 1, Me, Me.hWnd) = 0 Then
        Set myGifViewer = Nothing
        MsgBox "Cannot display GIF file", vbExclamation + vbOKOnly
        Unload Me
    Else
        With myGifViewer
            'TIP. If wanting to display the first image immediately, you must
            ' call SetAnimationBkg in the IGifRender_Rendered event, otherwise
            ' you would call it here, after the successful loading of the GIF
            ' .SetAnimationBkg gfdSolidColor, Me.BackColor, 0, 0, Me.ScaleWidth, Me.ScaleHeight, gfsShrinkScaleToFit Or gfsCentered
            .AnimationState = gfaPlaying
        End With
    End If
End Sub

Public Sub ChangeBackColor(ByVal newColor As Long)
    If Not myGifViewer Is Nothing Then
        myGifViewer.SetAnimationBkgColor newColor
        Me.BackColor = newColor
        If Me.AutoRedraw = False Then myGifViewer.AnimationState = gfaRefresh
    End If
End Sub

Private Sub IGifRender_GetRenderDC(ByVal ViewerID As Long, ByVal FrameIndex As Long, destDC As Long, Optional hwndRefresh As Long, Optional bAutoRedraw As Boolean = False, Optional PostNotify As Boolean = False)
    destDC = Me.hdc             ' render to this DC
    bAutoRedraw = Me.AutoRedraw ' lets class use appropriate refresh method
    hwndRefresh = Me.hWnd   ' refresh me when frame is rendered
    PostNotify = True       ' tell me when frames are rendered
End Sub

Private Sub IGifRender_Rendered(ByVal ViewerID As Long, ByVal FrameIndex As Long, ByVal Message As RenderMessage, ByVal MsgValue As Long)
    If Message = msgProgress Then
        Me.Caption = "Loading... " & MsgValue & "% Completed"
        
        If MsgValue = 0 Then
            ' we want the viewer to cache our background for flicker free drawing
            ' And also we want viewer to scale our image for us too; basically,
            ' we want the viewer to be self-sufficient. So supply it with what it
            ' needs...
            
            ' TIP #1: By calling .SetAnimationBkg here, then the 1st frame will immediately
            ' be rendered when it is done being processed. Otherwise, .SetAnimationBkg should
            ' be called after LoadGIF and the 1st frame will be rendered after all frames are processed
            
            ' Tip #2: If this is what you want, ensure the DC being drawn to has AutoRedraw=False, otherwise
            ' the frame will still be rendered by VB may not update the screen immediately
            myGifViewer.SetAnimationBkg gfdSolidColor, Me.BackColor, 0, 0, Me.ScaleWidth, Me.ScaleHeight, gfsShrinkScaleToFit Or gfsCentered
        End If
    
    ElseIf Message = msgRendered Then
        Me.Caption = "Frame " & FrameIndex
    
    ElseIf Message = msgLoopsEnded Then
        ' for gee whiz...
        Debug.Print "Current GIF animated its required loops: "; MsgValue
    End If
End Sub
