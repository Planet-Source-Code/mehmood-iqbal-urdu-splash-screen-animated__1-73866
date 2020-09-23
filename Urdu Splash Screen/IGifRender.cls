VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGifRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum RenderMessage
    msgRendered = 0
    msgProgress = 1
    msgLoopsEnded = 2
End Enum

Public Sub GetRenderDC(ByVal ViewerID As Long, ByVal FrameIndex As Long, ByRef destDC As Long, Optional ByRef hwndRefresh As Long, _
    Optional ByRef bAutoRedraw As Boolean = False, _
    Optional ByRef PostNotify As Boolean = False)

End Sub

Public Sub Rendered(ByVal ViewerID As Long, ByVal FrameIndex As Long, ByVal Message As RenderMessage, ByVal MsgValue As Long)

End Sub
