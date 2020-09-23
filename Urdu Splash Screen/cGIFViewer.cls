VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGIFViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private z_CbMem   As Long
Private z_Cb()    As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'-------------------------------------------------------------------------------------------------

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Type PICTDESC
    Size As Long
    Type As Long
    hHandle As Long
    hPal As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hWnd As Long, ByRef lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32.dll" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, ByRef pcRGBQuad As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, ByRef lppt As POINTAPI) As Long
Private Declare Function GetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long


Private Const HALFTONE As Long = 4
Private Const OBJ_DC As Long = 3
Private Const OBJ_MEMDC As Long = 10
Private Const PALETTECOUNT = 256&

' Standard Window UDTs
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D        ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SAFEARRAYBOUND
End Type
Private Type BITMAPINFOHEADER   ' used to create a DIB
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiPalette(0 To PALETTECOUNT - 1) As Long
End Type
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


' Custom UDTs
Private Type ColorTableSTRCUT
    Index As Long               ' 0=global; 1-xxx are local tables, if any
    Tables() As Long            ' 2 dim array (0 to 255, 0 to xxx); color table(s)
End Type
Private Type GIFcoreProperties
    Width As Long               ' overall GIF width
    Height As Long              ' overall height
    Loops As Long               ' Nr loops defined in GIF (0=infinite)
    Disposal3 As Byte           ' Uses Disposal rule #3. See RenderFrame_Timer for usage
    BkgIndex As Byte            ' Global palette index for suggested window bkg color
End Type
Private Type GIFframeProperties
    Dimensions As RECT          ' bounding rectangle of image
    Delay As Long               ' length of time (ms) frame stays visible (0 or -1 is infinite)
    TblIndex As Long            ' the ColorTableSTRCUT this frame uses
    imgOffset As Long           ' file byte postion where image begins in gif file
                                ' & after decompressed, position where img begins in DIB strip
    TransIndex As Byte          ' which palette index is to be transparent
    IsTransparent As Byte       ' does frame use transparency (0=no, 1=yes)
    Disposal As Byte            ' the disposal method for this frame (0-7)
End Type
Private Type UserDCInfo         ' created in SetAnimationBkg (unused for most single frame GIFs)
    DC As Long                  ' off-screen DC
    bmp As Long                 ' bitmap for flicker-free drawing
    Brush As Long               ' only used if animation bkg is solid color
    Mode As ScaleGIFConstants   ' scaling options
    Bounds As RECT              ' target destination dimensions
    ScaleCx As Single ' pre calculated ratios when scaling/stretching is used
    ScaleCy As Single ' pre calculated ratios when scaling/stretching is used
End Type

' Custom Enumerations
Public Enum GifFrameProps        ' expose Read-only properties a coder would need
    gfpDelay = 0                 ' to animate the GIF on their own if they so choose.
    gfpLeft = 1
    gfpTop = 2
    gfpWidth = 3
    gfpHeight = 4
    gfpDisposal = 5
    gfpFrameCount = 6
End Enum
Public Enum ScaleGIFConstants
    gfsClip = 0                  ' will never scale, nor stretch
    gfsScaleAlways = 1           ' will always scale up or down as needed
    gfsStretch = 2               ' will stretch to fit, not scaled, distorted
    gfsShrinkScaleToFit = 3      ' will only scale down if needed else scale is 1:1
    gfsCentered = 8              ' can be OR'd with above when creating animation background
End Enum
Public Enum AnimationActions     ' when class is self-animating; animation setting
    gfaStopped = 0               ' stop on current frame, reset current frame to first
    gfaPlaying = 1               ' start/restart from current frame
    gfaPaused = 2                ' stop on current frame, do not reset frame nr
    gfaStepping = 3              ' show next frame only, then pause
    gfaRefresh = 4               ' refresh current frame (useful for DCs where AutoRedraw=False)
End Enum
Public Enum BackgroundTypes      ' options when class caches GIF background (see SetAnimationBkg)
    gfdSolidColor = 0            ' background will be a solid color, uses less resources
    gfdBkgFromDC = 1             ' background copied/cached from supplied DC
End Enum
Public Enum StdPictureType       ' used to extract a frame in one of two formats
    gfiPicGIF = 0                ' not used by this class; but offered as FrameImage function
    gfiPicBMP = 1
End Enum
    
' All Class-Level variables are prefixed with c_

' used when compressing bitmap to GIF (see FrameImage routine)
Private c_BitBucketCount As Long    ' bit bucket bit counter
Private c_BitBucketBuff As Long     ' bit bucket contents
Private c_lCurrentBits As Long      ' current lzw compression size (variable)
Private c_lSubBlockSize As Long     ' flag to track last byte in a data subBlock

Private c_FileNum As Integer    ' file number for opened GIF
Private c_MinDelay As Integer   ' any delay less than this value will use this value
Private c_AniLoops As Integer   ' during animation: the current loop count
Private c_aniState As AnimationActions ' current animation state
Private c_curFrame As Long      ' during animation: which frame is being rendered
Private c_UserID As Long        ' user-supplied class index/reference
Private c_tmrAddressOf As Long  ' the AddressOf of the class's timer function
Private c_tmrOwner As Long      ' the hWnd assigned to the timer
Private c_DIB As Long           ' DIB strip handle
Private c_DC As Long            ' DC used for drawing frames to other DCs
Private c_EOF As Long           ' used to prevent "read past end of file" errors

Private c_IFace As IGifRender ' user-supplied interface if class uses own timer
Private c_UserDCdata As UserDCInfo ' used for flicker free drawing (back buffer)
Private c_gifProps As GIFcoreProperties ' overall GIF properties
Private c_Frames() As GIFframeProperties ' collection of individual frame properties
Private c_ColorTables As ColorTableSTRCUT  ' collection of color tables used in the GIF

Private c_aBuff() As Byte           ' general use byte array
' following used during decompressing GIF & released immediately
Private c_DataLen() As POINTAPI     ' tracks frame file position and size
Private c_DIBarray() As Byte        ' DMA overaly of the color strip
Private c_aPOT() As Long            ' Power of 2 look up table
Private c_maskTable() As Long       ' GIF Mask palette

Private Sub Class_Initialize()
    c_MinDelay = 50     ' class minimal frame delay; can be changed by user. 10 is absolute minimum
End Sub
Private Sub Class_Terminate()
    Me.UnloadGIF    ' free memory objects & clear arrays
End Sub

Public Function LoadGIF(ByVal gifFileName As String, ByVal ViewerID As Long, _
        Optional ByRef GIFRenderInterface As IGifRender = Nothing, _
        Optional ByVal timerHwnd As Long) As Long

' This function loads a gif from file and returns the number of frames within that file.
' Return Values:
' = 0 :: invalid gif file or major errors (i.e., no system resources/memory)
' > 0 :: successful. Return value is the number of frames in the gif &
'        class will use its own timer if applicable
' < 0 :: partial success. Abs(Return value) is number of frames in the gif &
'        creation of class timer failed or timerHwnd is invalid. You must use
'        your own timers and render frames by calling Class's RenderFrame function

' \ \ \ \ \ \ \ \ \ \ \ \ PARAMETERS \ \ \ \ \ \ \ \ \ \ \ \ \
' gifFileName :: full path and filename of an existing GIF file
' ViewerID :: user-supplied reference. This ID is passed in the IGifRender's
'             GetRenderDC & Rendered events
' GIFRenderInterface :: If used, you must enter Implements IGifRender in your
'             declarations section. Pass the keyword Me as the parameter.
'             The Interface is not used if timerhWnd is not a valid window handle
'             or the GIF file failed to load properly. See IGifRender class
' timerHwnd :: must be passed if class is to use its own timer. This should be Me.hWnd
'             and is used to ensure timer is killed when that hWnd terminates.
'             Ignored if GIFRenderInterface is Nothing or gif contains only 1 frame

' If your gif is a byte array extracted from a usercontrol property bag or from a
' RES file, save the byte array to disk first then pass the file name to this routine.

' Notes about optional usage of the GIFRenderInterface. See that class's comments too.
' 1. The class can create & use its own timer, thanx to Paul Caton's excellent thunks.
'   To enable that, you must pass the GIFRenderInterface parameter as Me, and also
'   pass the timerHwnd parameter as Me.hWnd
' 2. Passing a DC for class usage isn't reliable. When a DC using the AutoRedraw property
'   is used, the DC can change by VB without warning which would cause massive leaks
'   during the drawing process. Therefore, the class will ask you for the destination
'   DC each time a frame is about to be drawn. This is the IGifRender.GetRenderDC event.
'   By returning the DC's AutoRedraw and hWnd, this class can auto-refresh the DC for you.
' 4. If you opt not to use the IGifRender interface, you can still use the class
'   and your own timer(s). Simply call this class's RenderFrame function each time
'   your VB timer event fires.
' 5. The SetAnimationBkg & SetAnimationBkgColor routines aren't usable if
'   you don't use the interface and are only functional after successfully loading a GIF
' 6. Whether or not you use the IGifRender interface, it must be included in your
'   project or you will need to edit this class to remove any references to it.

    Dim newProps As GIFcoreProperties
    Dim nrItems As Long
    
    Call Me.UnloadGIF        ' kill any timer, release/reset some memory if needed
    c_UserID = ViewerID      ' cache user's reference
    c_gifProps = newProps    ' start a fresh UDT
    
    ' setup a "2^(0-8)" lookup table used by ParseGIF & BuildDIBstrip
    ReDim c_aPOT(0 To 8)
    c_aPOT(0) = 1
    For nrItems = 1 To 8
        c_aPOT(nrItems) = c_aPOT(nrItems - 1) * 2
    Next
    
    ' scan for key properites from the GIF file
    nrItems = ParseGIF(gifFileName)

    On Error GoTo EH
    If nrItems = 0 Then ' couldn't parse the file
        Call Me.UnloadGIF   ' bug out
        Exit Function
    Else
        ' see if user wants class timer & interface or not?
        If IsWindow(timerHwnd) = 0 Then timerHwnd = 0
        If (GIFRenderInterface Is Nothing) Or (timerHwnd = 0) Then
            nrItems = -nrItems '  user must do their own timers
        Else
            Set c_IFace = GIFRenderInterface
            ' yep, let's ensure we can create a class timer
            c_tmrAddressOf = cb_AddressOf(1, 4)
            If Not c_tmrAddressOf = 0 Then   ' so far so good
                c_tmrOwner = timerHwnd
                ' test the timer function
                If SetTimer(c_tmrOwner, ObjPtr(Me), 60000, c_tmrAddressOf) = 0 Then
                    ' timer failure, user must do their own timers
                    nrItems = -nrItems
                    c_tmrAddressOf = 0
                    Set c_IFace = Nothing   ' prevents automation
                Else ' success, kill the test timer
                    KillTimer c_tmrOwner, ObjPtr(Me)
                End If
            End If
            ' clear array if timer failed or single frame
            If c_tmrAddressOf = 0 Or nrItems = 1 Then Erase z_Cb()
        End If
        ' transfer GIF to a DIB strip (all frames in one bitmap)
        If Not c_IFace Is Nothing Then ' send 100% progress to user
            ' send notification to user we finished parsing and are begining creating GIF frames
            ' This is the point where the user can call SetAnimationBkg which then allows the
            ' first frame to be immediately rendered while the other frames are still being processed
            c_IFace.Rendered c_UserID, 1, msgProgress, 0
        End If
        If BuildDIBstrip = False Then
            UnloadGIF               ' error; bug out
            Exit Function
        End If
    End If
    Close c_FileNum
    c_FileNum = 0
    Erase c_aPOT()
    Erase c_DataLen()
    
    If Not c_IFace Is Nothing Then ' send 100% progress to user
        If Not nrItems = 0 Then c_IFace.Rendered c_UserID, Abs(nrItems), msgProgress, 100
    End If
    
    LoadGIF = nrItems

EH:
    If Err Then
        Err.Clear
        UnloadGIF
    End If
End Function

Public Sub UnloadGIF()
    ' frees memory, releases any timer subclassing
    ' IF YOU USE END IN YOUR APP, CALL THIS BEFORE END IS EXECUTED
    On Error Resume Next
    If Not IsArrayEmpty(Not z_Cb) Then  ' remove any timer
        KillTimer c_tmrOwner, ObjPtr(Me)
        cb_Terminate
        Erase z_Cb()
    End If
    c_tmrAddressOf = 0      ' also serves as flag indicating no class timer
    Set c_IFace = Nothing   ' remove any interface
    Erase c_Frames()        ' clear all GIF frame info
    c_aniState = gfaStopped  ' current animation state
    c_curFrame = 0          ' reset frame index
    ' clean up GDI memory objects
    With c_UserDCdata
        If Not .bmp = 0 Then  ' delete any back buffer bitmap
            DeleteObject SelectObject(.DC, .bmp)
            .bmp = 0
        End If
        If Not .DC = 0 Then   ' delete any back buffer DC
            DeleteDC .DC
            .DC = 0
        End If
        If Not .Brush = 0 Then    ' delete any colored brush
            DeleteObject .Brush
            .Brush = 0
        End If
    End With
    If Not c_DIB = 0 Then
        DeleteObject SelectObject(c_DC, c_DIB) ' clear DIB strip & DC
        DeleteDC c_DC
        c_DIB = 0
        c_DC = 0
    End If
    ' finally clear any used arrays & close GIF file if needed
    Erase c_ColorTables.Tables()    ' global & local palettes
    Erase c_maskTable()             ' mask palette
    Erase c_aPOT()                  ' power of two lookup (should already be cleared)
    Erase c_DataLen()               ' gif image positions within file (should already be cleared)
    c_ColorTables.Index = 0         ' reset for next gif file
    If Not c_FileNum = 0 Then       ' ensure we don't keep the gif file open
        Close c_FileNum
        c_FileNum = 0
    End If
    If Err Then Err.Clear
End Sub

Public Function RenderFrame(ByVal Index As Long, ByVal hdc As Long, _
            Optional ByVal hWnd2Refresh As Long = 0, Optional ByVal AutoRedraw As Boolean = False, _
            Optional ByVal destX As Long = 0, Optional ByVal destY As Long = 0, _
            Optional ByVal destCx As Long = 0, Optional ByVal destCY As Long = 0, _
            Optional ByVal StretchMode As ScaleGIFConstants = gfsShrinkScaleToFit) As Boolean

 ' Function is basically a StretchBlt type of call with a few other parameters
 ' This exposes the class's real render function to the user should the user
 '      want to render the frame without using the class timer
 
 ' Parameters:
 ' Index : frame index to render
 ' hDC : destination DC. If the DC is unknown (i.e., VB Frame controls), you can
 '      pass the .hWnd instead, the routine will get the DC from it.
 ' hWnd2Refresh : after rendering, class will refresh passed hwnd (optional)
 ' AutoRedraw : if you want function to refresh the passed hDC, identify whether
 '      or not it has .AutoRedraw set to True. Depending on how that property is
 '      set, different APIs are used
 ' destX,destY : DC's X,Y location to begin rendering (non-negative values only)
 ' destCx,destCy : the width,height to be rendered to. If 0, then frame's width/height used
 ' StretchMode : if destCx,destCy are not the same as the frame's width,height
 '      then this determines how the image will fit the destination rectangle

    If IsArrayEmpty(Not c_Frames) Then Exit Function
    If Index < 1 Or Index > UBound(c_Frames) Then Exit Function
    
    If destCx < 1 Then destCx = c_Frames(Index).Dimensions.Right
    If destCY < 1 Then destCY = c_Frames(Index).Dimensions.Bottom
    
    ' pass off to the worker function
    RenderFrame = RenderFrame_Timer(Index, hdc, hWnd2Refresh, AutoRedraw, False, _
        False, destX, destY, destCx, destCY, StretchMode)
        
End Function

Public Property Get FrameProperty(ByVal Index As Long, ByVal propID As GifFrameProps) As Long
    
    On Error GoTo EH
    ' Only those properties that you would need to render your own frames are exposed.
    ' All properties are read-only; this is not a gif editor
    If IsArrayEmpty(Not c_Frames) Then Exit Property
    If Index < 1 Or Index > UBound(c_Frames) Then Exit Property
    
    Select Case propID
        Case gfpDelay:
            If c_Frames(Index).Delay < c_MinDelay Then
                FrameProperty = c_MinDelay
            Else
                FrameProperty = c_Frames(Index).Delay
            End If
        Case gfpLeft: FrameProperty = c_Frames(Index).Dimensions.Left
        Case gfpTop: FrameProperty = c_Frames(Index).Dimensions.Top
        Case gfpWidth: FrameProperty = c_Frames(Index).Dimensions.Right
        Case gfpHeight: FrameProperty = c_Frames(Index).Dimensions.Bottom
        Case gfpDisposal: FrameProperty = c_Frames(Index).Disposal
        Case gfpFrameCount
            If Not IsArrayEmpty(Not c_Frames) Then FrameProperty = UBound(c_Frames)
    End Select
EH:
End Property

Public Property Get AnimationIndex() As Long    ' READ ONLY
    AnimationIndex = c_curFrame ' return most recent frame rendered
    '^^ not reliable if animating because after this is returned it will change
    '   when the next frame of an animated GIF is rendered. Is reliable if the
    '   animation is stopped, paused or stepped.
End Property
Public Property Get AnimationState() As AnimationActions
    AnimationState = c_aniState  ' returns current state
End Property
Public Property Let AnimationState(ByVal Action As AnimationActions)
    
  ' This property allows the user to stop, start or pause the animation displayed
  ' by this class. If no gif is loaded or, when the gif was loaded, the user did
  ' not pass IGifRender interface or, a timer was not created and not a single-frame
  ' GIF, then this property will abort (See LoadGif)
  
    On Error GoTo EH
    If c_IFace Is Nothing Then Exit Property
    If Action < gfaStopped Or Action > gfaRefresh Then Action = gfaStopped
    
    Select Case Action
    Case gfaStopped, gfaPaused
        If UBound(c_Frames) = 1 Or c_aniState = gfaPaused Then
            c_aniState = gfaStopped
        Else ' only stop/pause if playing
            If c_aniState = gfaPlaying Then KillTimer c_tmrOwner, ObjPtr(Me)
            c_aniState = Action
        End If
        
    Case gfaPlaying
        If c_aniState = gfaPlaying Then Exit Property
        If c_aniState = gfaStopped Then ' reset loop counter if needed
            c_AniLoops = 0
            c_curFrame = 1
        Else ' was paused; increment frame and restart timer
            c_curFrame = c_curFrame + 1
        End If
        c_aniState = gfaPlaying
        AnimationState = gfaRefresh  ' show frame w/o timer first; then start timer
        '^^ By doing it this way, we ensure the AnimationIndex property is correct when
        '   this routine returns; otherwise, the timer may fire before user calls their
        '   next line of code and the AnimationIndex will have changed
        If UBound(c_Frames) > 1 Then
            SetTimer c_tmrOwner, ObjPtr(Me), FrameProperty(c_curFrame, gfpDelay), c_tmrAddressOf
        Else
            c_aniState = gfaStopped
            c_IFace.Rendered c_UserID, 1, msgLoopsEnded, 1
        End If
        
    Case gfaStepping ' stepping forward a frame
        If UBound(c_Frames) = 1 Then
            c_aniState = gfaStopped ' set current state for single-frame gifs
        Else
            If c_aniState = gfaPlaying Then    ' was animating; stop timer
                KillTimer c_tmrOwner, ObjPtr(Me)
                c_AniLoops = 0
            Else
                c_curFrame = c_curFrame + 1 ' move to next frame
            End If
            c_aniState = Action
        End If
        AnimationState = gfaRefresh  ' show frame
        
    Case gfaRefresh
        Dim tDC As Long, tHwnd As Long, bAutoRedraw As Boolean, bNotify As Boolean
        ' validate low/high frame index
        If c_curFrame < 1 Or c_curFrame > UBound(c_Frames) Then c_curFrame = 1
        ' get drawing DC and hWnd (if necessary) from client
        Call c_IFace.GetRenderDC(c_UserID, c_curFrame, tDC, tHwnd, bAutoRedraw, bNotify)
        If Not tDC = 0 Then ' draw the frame
            RenderFrame_Timer c_curFrame, tDC, tHwnd, bAutoRedraw, bNotify, True
        End If
    End Select
EH:
End Property

Public Function SetAnimationBkgColor(ByVal bkgColor As Long) As Boolean
    ' function can be useful for animation backgrounds that are solid colors
    ' This will change the color on the fly and is only available after the
    ' SetAnimationBkg function has been called and only if it was established
    ' with a solid color background
    
    Dim tRect As RECT, bRestart As Boolean
    If c_UserDCdata.Brush = 0 Then Exit Function
    ' we do this by pausing animation, updating the offscreen DC and brush
    If Me.AnimationState = gfaPlaying Then
        bRestart = True
        Me.AnimationState = gfaPaused
    End If
    ' destroy old & create new
    DeleteObject c_UserDCdata.Brush
    If bkgColor < 0 Then bkgColor = GetSysColor(bkgColor And &HFF&)
    c_UserDCdata.Brush = CreateSolidBrush(bkgColor)
    ' fill our back buffer and restart animation if needed
    SetRect tRect, 0, 0, c_UserDCdata.Bounds.Right + 1, c_UserDCdata.Bounds.Bottom + 1
    FillRect c_UserDCdata.DC, tRect, c_UserDCdata.Brush
    If bRestart Then Me.AnimationState = gfaPlaying
    
End Function

Public Function SetAnimationBkg(ByVal bkgMethod As BackgroundTypes, ByVal bkgDC_Color As Long, _
            Optional ByVal destX As Long, Optional ByVal destY As Long = 0, _
            Optional ByVal destCx As Long, Optional ByVal destCY As Long = 0, _
            Optional ByVal StretchMode As ScaleGIFConstants = gfsShrinkScaleToFit, _
            Optional ByVal bkgDC_X As Long = -1, Optional bkgDC_Y As Long = -1) As Boolean

  ' Function will use parameters to create an offscreen rendering background bitmap/DC
  ' Must be called after a gif is loaded because it uses the overall gif size
  ' to help determine default values and scaling.
  
  ' False return value indicates failure; invalid parameter(s)
  
  ' Parameters:
  ' bkgMethod: either gfdSolidColor or gfdBkgFromDC
  ' bkgDC_Color: (DC or Color used for the background of a GIF)
  '     if bkgMethod=gfdSolidColor then this is a color value for the target DC
  '     if bkgMethod=gfdBkgFromDC then a copy of the DC area will be used
  ' destX,Y : The X,Y coordinates where rendering will take place on target DC
  ' destCx,Cy : The width and height of the rendering (zeros use gif's height/width)
  ' bkgDC_X,Y
  '     if bkgMethod=gfdSolidColor then bkgDC_X,Y are ignored
  '     if bkgMethod=gfdBkgFromDC then they are X,Y coordinates from the DC to copy
  '         background from. The width,height copied from that DC is the destCx,Y
  '         values. If bkgDC_X,Y are < 0, then destX and destY will be used
  ' This option allows you to use a DC/bitmap, other than the target destination
  ' for the GIF's background (i.e., an offscreen DC)
  
    ' used only when class timers are animating
    If c_tmrAddressOf = 0 Then Exit Function
  
    Dim tDC As Long, bmpMultiplier As Long, bRestart As Boolean
    Dim hWnd2release As Long, tBMP As Long, tRect As RECT
    Dim f As Long
    Dim oldBuffSettings As UserDCInfo
    
    If c_aniState = gfaPlaying Then    ' if animating, pause til we finish
        Me.AnimationState = gfaPaused
        bRestart = True
    End If
    oldBuffSettings = c_UserDCdata
    
    ' validate some parameters
    If destCx < 1 Then destCx = Me.Width    ' use gif overall width
    If destCY < 1 Then destCY = Me.Height   ' use gif overall height
    
    ' scale as needed, can reduce size of back buffer
    With c_UserDCdata
        SetRect .Bounds, destX, destY, destCx, destCY
        .Mode = StretchMode
    End With
    
    ' check for one exception. If all frames are the same height/width and all frames
    ' use disposal code other than 2/3, and no frames use transparency, then we
    ' don't need a background buffer, regardless how many frames exist
    If UBound(c_Frames) = 1 Then
        ' for single frame GIFs, the only reason for a backbuffer would be transparency
        If c_Frames(1).IsTransparent = 0 Then f = 2 ' no buffer else f=0 & buffer needed
    Else
        For f = 1 To UBound(c_Frames)
            If c_Frames(f).IsTransparent = 1 Then Exit For
            If c_Frames(f).Disposal = 2 Then Exit For ' erase target
            If c_Frames(f).Disposal = 3 Then Exit For ' image ontop of image code
            If Not c_Frames(f).Dimensions.Right = c_gifProps.Width Then Exit For
            If Not c_Frames(f).Dimensions.Bottom = c_gifProps.Height Then Exit For
        Next
    End If
    With c_UserDCdata
        If Not .Brush = 0 Then    ' always destroy any previous brush
            DeleteObject .Brush
            .Brush = 0
        End If
        If f > UBound(c_Frames) Then    ' don't need a back buffer
            If Not .bmp = 0 Then        ' clean up any pre-existing memory objects
                DeleteObject SelectObject(.DC, .bmp)
                DeleteDC .DC
                .DC = 0
                .bmp = 0
            End If
            SetAnimationBkg = True  ' return True, drawing will be directly onto
            Exit Function           ' the target vs off-screen drawing
        End If
    End With
        
    ' the internal offscreen bitmap will be destCx wide, but height is variable:
    '   if any frame in the GIF uses disposal method 3, then height=height + (another height)
    '   if bkgMethod=gfdBkgFromDC, then height=height + (another height)
    '   if bkgMethod=gfdSolidColor, then height=height - (a height)
    ' determine initial multiplier whether or not disposal 3 is used
    If c_gifProps.Disposal3 = 1 Then bmpMultiplier = 3 Else bmpMultiplier = 2
    
    If bkgMethod = gfdBkgFromDC Then        ' ensure passed DC is a dc
        Select Case GetObjectType(bkgDC_Color)
        Case OBJ_DC, OBJ_MEMDC
        Case Else   ' nope, but like VB.Frames, we'll get dc from hWnd
            If IsWindow(bkgDC_Color) = 0 Then
                ' can't use what was passed
                c_UserDCdata = oldBuffSettings
                If bRestart Then Me.AnimationState = gfaPlaying
                Exit Function ' abort
            End If
            ' user passed a window handle; get DC from it
            hWnd2release = bkgDC_Color
            bkgDC_Color = GetDC(hWnd2release)
        End Select
    Else
        bmpMultiplier = bmpMultiplier - 1 ' solid color bkgs don't use extra bmp space
    End If
    
    On Error GoTo EH
    tDC = GetDC(0&)
    With c_UserDCdata
        ' scale GIF dimensions to target dimensions based on passed StretchMode
        ScaleToDestination 0&, .Bounds.Left, .Bounds.Top, .Bounds.Right, .Bounds.Bottom, _
            StretchMode, .ScaleCx, .ScaleCy
        If .DC = 0 Then   ' create temp DC unless already created
            .DC = CreateCompatibleDC(tDC)
            If Not .DC = 0 Then
                SetStretchBltMode .DC, HALFTONE   ' better StretchBlt results
                ' create back buffer bitmap (flicker-free drawing)
                tBMP = CreateCompatibleBitmap(tDC, .Bounds.Right, .Bounds.Bottom * bmpMultiplier)
            End If
        Else   ' delete previous bitmap if needed & recreate it
            If destCx > .Bounds.Right Or destCY > .Bounds.Bottom Then
                If Not .bmp = 0 Then DeleteObject SelectObject(.DC, .bmp)
                .bmp = 0
                ' create new bitmap and select it into the temp DC
                tBMP = CreateCompatibleBitmap(tDC, .Bounds.Right, .Bounds.Bottom * bmpMultiplier)
            End If
        End If
        ReleaseDC 0&, tDC
        If tBMP = 0 Then  ' error. can't create offscreen bitmap
            GoTo EH
        Else 'select our back buffer bitmap into our DC
            .bmp = SelectObject(.DC, tBMP)
        End If
    End With
    
    ' here we are initializing the back buffer bitmap
    If bkgMethod = gfdBkgFromDC Then
        'copy the passed DC to our temp DC in up to 2 places (stacked on each other)
        If bkgDC_X < 0 Then bkgDC_X = c_UserDCdata.Bounds.Left
        If bkgDC_Y < 0 Then bkgDC_Y = c_UserDCdata.Bounds.Top
        With c_UserDCdata.Bounds
            ' copy to cached section of bitmap
            BitBlt c_UserDCdata.DC, 0, .Bottom, .Right, .Bottom, bkgDC_Color, bkgDC_X, bkgDC_Y, vbSrcCopy
            ' copy to drawing section of bitmap
            BitBlt c_UserDCdata.DC, 0, 0, .Right, .Bottom, c_UserDCdata.DC, 0, .Bottom, vbSrcCopy
        End With
    Else    ' solid color bkg, create brush so we don't have to create it for every frame
        If bkgDC_Color < 0 Then bkgDC_Color = GetSysColor(bkgDC_Color And &HFF&)
        c_UserDCdata.Brush = CreateSolidBrush(bkgDC_Color)
        ' fill drawing portion of the bitmap
        SetRect tRect, 0, 0, c_UserDCdata.Bounds.Right, c_UserDCdata.Bounds.Bottom
        FillRect c_UserDCdata.DC, tRect, c_UserDCdata.Brush
    End If
    
    If Not hWnd2release = 0 Then ReleaseDC hWnd2release, bkgDC_Color
    If bRestart Then Me.AnimationState = gfaPlaying
    SetAnimationBkg = True
    Exit Function
    
EH: 'clean up
    With c_UserDCdata
        If Not .bmp Then DeleteObject SelectClipRgn(.DC, .bmp)
        If Not .DC = 0 Then DeleteDC .DC  ' if user ignores return value, any drawing will be
        .DC = 0                           ' done directly on target. This won't look good!
        .bmp = 0
        SetRect .Bounds, destX, destY, destCx, destCY
    End With
    If Not hWnd2release = 0 Then ReleaseDC hWnd2release, bkgDC_Color

End Function

Public Sub ScaleToDestination(ByVal Index As Long, ByRef X As Long, ByRef Y As Long, _
                ByRef CX As Long, ByRef CY As Long, _
                ByVal Mode As ScaleGIFConstants, _
                ByRef xRatio As Single, ByRef yRatio As Single)

    ' function scales an image to the target destination based on the
    ' stretch mode setting, image dimensions, and target dimensions
    ' Originally a private routine; but modified and made public so coders
    '   don't need to write their own scaling functions
    
    ' If Index = 0 then the overall GIF scaling information is returned; otherwise
    ' the scaling info is only for the passed frame.
    
    ' If Index = 0:
    '   - the xRatio, yRatio can be used for every frame so this routine isn't called for each frame.
    '   - the X,Y,CX,CY parameters may be changed if Mode <> gfsStretch
    '   - the goal is to back buffer the least amount of target DC possible
    '   - How this class uses the modified parameters for any frame to be drawn:
    '     FrameLeft = X + Int(ActualLeft * xRatio) ( X may be changed if centering)
    '     FrameTop = Y + Int(ActualTop * yRatio)   ( Y may be changed if centering)
    '     FrameWidth = Int(ActualWidth * xRatio)
    '     FrameHeight = Int(ActualHeight * yRatio)
    
    On Error GoTo EH
    If IsArrayEmpty(Not c_Frames) Then Exit Sub  ' no gif loaded yet
    
    If Not Index = 0 Then ' scaling is for the overall GIF else for individual frames
        If Index < 1 Or Index > UBound(c_Frames) Then Index = 0
    End If
    
    If (Mode And Not gfsCentered) = gfsClip Then
        xRatio = Me.Width / Me.Width    ' funky? Well this prevents a bogus error saying
        yRatio = Me.Height / Me.Height  ' "expression too complex" which I get when setting
        If Mode = gfsCentered Then ' centering needs more calcs
            If Me.Width < CX Then   ' find where left of overall gif would start
                X = Int((CX - Me.Width) \ 2) + X
                CX = Me.Width
            End If
            If Me.Height < CY Then  ' find where top of overall gif would start
                Y = Int((CY - Me.Height) \ 2) + Y
                CY = Me.Height
            End If
        End If
        If Index = 0 Then
            If Me.Width < CX Then CX = Me.Width
            If Me.Height < CY Then CY = Me.Height
        Else ' for the individual frame; not called by this class
             ' ensure frame's right and bottom edge stay within bounds
            With c_Frames(Index).Dimensions
                If X + .Right + .Left > CX Then CX = CX - .Left Else CX = .Right
                X = X + .Left    ' frame's new left position
                If Y + .Bottom + .Top > CY Then CY = CY - .Top Else CY = .Bottom
                Y = Y + .Top     ' frame's new top position
            End With
        End If
    
    Else ' scaling in one way or another....
        xRatio = CX / Me.Width
        yRatio = CY / Me.Height
        Select Case (Mode And Not gfsCentered)
        Case gfsStretch
            ' nothing to calculate when stretching
        Case gfsShrinkScaleToFit
            If xRatio > 1 And yRatio > 1 Then
                xRatio = 1  ' image will fit without scaling; use 1:1 scaling
                yRatio = 1
            Else            ' image must be scaled; use same ratio for width/height
                If xRatio > yRatio Then xRatio = yRatio Else yRatio = xRatio
            End If
        Case Else ' always scale to target dimensions; use same ratio for width/height
            If xRatio > yRatio Then xRatio = yRatio Else yRatio = xRatio
        End Select
    
        ' no centering when stretching 'cause image is stretched to target area
        If Not ((Mode And Not gfsCentered) = gfsStretch) Then
        
            If (Mode And gfsCentered) = gfsCentered Then  ' centering needs more calcs
                X = (CX - Int(Me.Width * xRatio)) \ 2 + X
                Y = (CY - Int(Me.Height * yRatio)) \ 2 + Y
                If Index = 0 Then
                    CX = Me.Width * xRatio
                    CY = Me.Height * yRatio
                Else    ' calculate centered position for this frame
                    X = X + Int(c_Frames(Index).Dimensions.Left * xRatio)
                    Y = Y + Int(c_Frames(Index).Dimensions.Top * yRatio)
                End If
            ElseIf Index = 0 Then
                CX = Me.Width * xRatio
                CY = Me.Height * yRatio
            End If
            If Not Index = 0 Then   ' calculate this frame's scaled width/height
                CX = Int(c_Frames(Index).Dimensions.Right * xRatio)
                CY = Int(c_Frames(Index).Dimensions.Bottom * yRatio)
            End If
        End If
    End If
EH:
End Sub

Public Property Get WindowBkgColor() As Long    ' READ ONLY
    ' This class does not use this information. Should you want it to set
    ' your DC background color, here it is.
    If Not IsArrayEmpty(Not c_ColorTables.Tables) Then ' have a global color table?
        Dim lRtnColor As Long
        ' convert BGR to RGB
        lRtnColor = c_ColorTables.Tables(c_gifProps.BkgIndex, 0)
        lRtnColor = (lRtnColor And &HFF) * &H10000 Or _
            ((lRtnColor \ &H100&) And &HFF) * &H100& Or _
                ((lRtnColor \ &H10000) And &HFF)
        WindowBkgColor = lRtnColor
    End If
End Property
Public Property Get Height() As Long    ' READ ONLY
    Height = c_gifProps.Height
End Property
Public Property Get Width() As Long     ' READ ONLY
    Width = c_gifProps.Width
End Property
Public Property Let LoopCount(ByVal newCount As Long)
    ' how many loops to perform? 0 or -1 = infinite
    If newCount = -1 Then newCount = 0
    c_gifProps.Loops = (newCount And &HFFFF&)
End Property
Public Property Get LoopCount() As Long
    LoopCount = c_gifProps.Loops
End Property
Public Property Let MinimumDelay(ByVal newMinimum As Long)
    ' Some GIFs have zero for their delay. Zero is not valid
    ' What should the absolute minimum delay be before a next frame is displayed?
    ' Per MSDN, with NT, setting a timer less than 10 ms will default to 10 ms.
    c_MinDelay = (newMinimum And &HFFFF&)
    If c_MinDelay < 10 Then c_MinDelay = 10
End Property
Public Property Get MinimumDelay() As Long
    MinimumDelay = c_MinDelay
End Property

Public Function FramePalette(ByVal Index As Long, Palette() As Long) As Boolean
    ' Returns the palette as a 256 color BGR array. Not used by this class.
    ' If function returns false, the Palette() is all black. Invalid Index or no GIF loaded
    ' Convert to RGB if needed after array is returned. Example:
    '   For x = 0 to 255
    '       myArray(x) = ((myArray(x) \ &H10000) And &HFF) Or _
    '                    ((myArray(x) \ &H100&) And &HFF) * &H100& Or _
    '                    (myArray(x) And &HFF) * &H10000
    '   Next x
    ' ---------------------------------------------------------------------
    ReDim Palette(0 To 255) As Long
    '^^ return something to prevent errors if user calls UBound()
    If Not IsArrayEmpty(Not c_Frames) Then
        If Index > 0 And Index < UBound(c_Frames) + 1 Then
            ' return the requested palette
            CopyMemory Palette(0), c_ColorTables.Tables(0, c_Frames(Index).TblIndex), PALETTECOUNT * 4
        End If
    End If
End Function

Public Function FrameImage(ByVal Index As Long, ByVal PicType As StdPictureType) As StdPicture
    
    ' function will return a GIF frame as a stdPicture either in GIF or BMP format.
    ' NOTE: Bitmap format should be returned when wanting to edit the image
    '           or wanting the image without transparency applied.
    '       GIF format should be returned to display the image with transparency
    '       -- the GIF format is not guaranteed to retain the actual palette indexes
    '          but the bitmap format will always maintain them.
    '       see BuildDIBstrip for more discussion on this.
    
    ' The returned image is always 8 bit, 256 colors
    
    If IsArrayEmpty(Not c_Frames) Then Exit Function
    If Index < 0 Or Index > UBound(c_Frames) Then Exit Function
    
    If PicType = gfiPicBMP Then ' return as bitmap stdPicture; else default to GIF stdPicture
    
        Set FrameImage = ConvertStripToBMP(Index)
    
    Else
    
        Set FrameImage = ConvertStripToGIF(Index)
    
    End If
    
End Function


' PRIVATE AND HELPER ROUTINES FOLLOW....

Private Function ConvertStripToBMP(Index As Long) As IPicture

    Dim tBMPI As BITMAPINFO, dibDC As Long, hDib As Long, hOldBmp As Long
    Dim lpPictDesc As PICTDESC, aGuid(0 To 3) As Long, bRestart As Boolean
    
    On Error GoTo EH
    With c_Frames(Index)
        With tBMPI.bmiHeader        ' fill in our DIB structure
            .biBitCount = 8
            .biClrUsed = PALETTECOUNT
            .biHeight = c_Frames(Index).Dimensions.Bottom
            .biWidth = c_Frames(Index).Dimensions.Right
            .biPlanes = 1
            .biSize = Len(tBMPI.bmiHeader)
        End With                    ' copy the palette to new DIB
        CopyMemory tBMPI.bmiPalette(0), c_ColorTables.Tables(0, .TblIndex), PALETTECOUNT * 4
        ' create DIB and exit out if errors encountered
        hDib = CreateDIBSection(c_DC, tBMPI, 0&, ByVal 0&, 0&, 0&)
        If hDib = 0 Then Exit Function
        dibDC = CreateCompatibleDC(c_DC)
        If dibDC = 0 Then
            DeleteObject hDib
            Exit Function
        End If
        hOldBmp = SelectObject(dibDC, hDib)
        If c_aniState = gfaPlaying Then
            bRestart = True ' pause animation until we copy the image
            Me.AnimationState = gfaPaused
        End If
        ' ensure we have original/full color palette in our DC; will never be by default
        SetDIBColorTable c_DC, 0, PALETTECOUNT, c_ColorTables.Tables(0, .TblIndex)
        c_ColorTables.Index = -1
        BitBlt dibDC, 0, 0, .Dimensions.Right, .Dimensions.Bottom, c_DC, .imgOffset, 0, vbSrcCopy
        If bRestart = True Then Me.AnimationState = gfaPlaying
    End With
    SelectObject dibDC, hOldBmp
    DeleteDC dibDC
    ' good to go; create stdPicture from the bitmap handle
    With lpPictDesc
        .Size = Len(lpPictDesc)
        .Type = vbPicTypeBitmap
        .hHandle = hDib
        .hPal = 0
    End With
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    aGuid(0) = &H7BF80980
    aGuid(1) = &H101ABF32
    aGuid(2) = &HAA00BB8B
    aGuid(3) = &HAB0C3000
    ' create stdPicture
    Call OleCreatePictureIndirect(lpPictDesc, aGuid(0), True, ConvertStripToBMP)
EH:
End Function
   
Private Function ConvertStripToGIF(Index As Long) As StdPicture

'--- Return as GIF; need to LZW compress bitmap bytes into GIF format
    
' HEAVILY MODIFIED AUTHOR'S ORIGINAL CODE (by Carles P.V.)
' - Reorganized the original routine; removed GOTOs and flow is smoother in hash routine
' - Pulled 11 global declarations into this routine.
' - The following original routines were combined into this one:
'   pvClearBlock, pvClearTable, pvCharInit & InitMasks
' - Overall changes might be a hair slower for small gifs but faster for bigger gifs
'   example using Win98's Clouds.BMP (8bit, 640x480); in IDE: 160 ms faster on my pc
' - code is not portable; customized for this class only; the portable code can
'   be found in my GIF editor project which will convert 1,2,4,8 bit bitmaps to GIF

  Const MAX_BITS                    As Long = 12  ' Per GIF docs, 12 is the Max
  Const MAX_BITSHIFT                As Long = 2 ^ MAX_BITS
  Const MAX_CODE                    As Long = 2 ^ MAX_BITS ' Should NEVER generate this code
  Const TABLE_SIZE                  As Long = 5003         ' 80% occupancy (hash)
  
  Dim LIdx     As Long      ' hash table index
  Dim lFCode   As Long      ' pixel pattern (hashed)
  Dim lC       As Long      ' most recent pixel
  Dim lEnt     As Long      ' previous known pattern
  Dim lDisp    As Long
  Dim m_lShift As Long
  
  Dim m_lCodeCount As Long ' count of codes used
  Dim m_lMaxCode As Long   ' Maximum code, given m_lBits
  '-- Block compression parameters.
  Dim m_lInitBits      As Long  ' baseline LZW compression size
  Dim m_lClearCode     As Long  ' clear code table & start again
  Dim m_lHashTable(0 To TABLE_SIZE - 1) As Long
  Dim m_lCodeTable(0 To TABLE_SIZE - 1) As Long

  ' Added by LaVolpe to read from the custom DIB strip
  Dim TSA As SAFEARRAY2D, dibBytes() As Byte    ' DMA DIB processing
  Dim bExistingPattern As Boolean                       ' indicates hash match found
  Dim dibRow As Long, dibCol As Long, aPtr As Long ' DIB looping parameters
  Dim tBMP As BITMAP  ' used to get DIB strip pointer & dimensions (not cached in class)
  Dim bRestart As Boolean
  
    'Added by LaVolpe -- gif frame custom build; byte at a time
    On Error GoTo EH
    ' pause animation if needed
    If Me.AnimationState = gfaPlaying Then
        Me.AnimationState = gfaPaused
        bRestart = True     ' flag indicating to restart animation
    End If
    ' the actual dib handle isn't cached, unselect it for the handle
    aPtr = SelectObject(c_DC, c_DIB)
    GetGDIObject aPtr, Len(tBMP), tBMP  ' use API to get the DIB bits pointer
    SelectObject c_DC, aPtr             ' reselect it into DC
    If tBMP.bmBits = 0 Then             ' should never happen since I control it
        If bRestart Then Me.AnimationState = gfaPlaying ' continue animation if needed
        Exit Function
    End If
    With TSA
        .rgSABound(0).cElements = tBMP.bmHeight
        .rgSABound(1).cElements = ByteAlignOnWord(8, tBMP.bmWidth)
        .cbElements = 1
        .cDims = 2
        .pvData = tBMP.bmBits   ' dib pointer
    End With
    CopyMemory ByVal VarPtrArray(dibBytes()), VarPtr(TSA), &H4
    
    ' oversize/guesstimate compressed GIF data & include GIF block information
    ReDim c_DIBarray(0 To 800 + (c_Frames(Index).Dimensions.Bottom * c_Frames(Index).Dimensions.Right))
    '^ 800= 13 hdr + 10 img descrip + 8 img ctrl + 768 palette + 1 EOF flag
    With c_Frames(Index)
        ' start building the GIF frame by hand, a byte at a time
        CopyMemory c_DIBarray(0), &H38464947, 4  ' add the 6byte GIF89a signature
        CopyMemory c_DIBarray(4), &H6139&, 2
        CopyMemory c_DIBarray(6), .Dimensions.Right, &H2   ' now the width
        CopyMemory c_DIBarray(8), .Dimensions.Bottom, &H2  ' & height
        ' add the color table flag and the table itself
        c_DIBarray(10) = 135 ' global color table @ 8 bits (128 or 7)
        ' can skip next 2 bytes: bkg window palette index & pixel aspect ratio
        aPtr = 13
        For LIdx = 0 To PALETTECOUNT - 1 ' convert BGR palette to RGB
            c_DIBarray(aPtr) = (c_ColorTables.Tables(LIdx, .TblIndex) \ &H10000) And &HFF
            c_DIBarray(aPtr + 1) = (c_ColorTables.Tables(LIdx, .TblIndex) \ &H100&) And &HFF
            c_DIBarray(aPtr + 2) = c_ColorTables.Tables(LIdx, .TblIndex) And &HFF
            aPtr = aPtr + 3
        Next
        ' add the image control block if needed
        If .IsTransparent = 1 Then
            CopyMemory c_DIBarray(aPtr), &H4F921, 3  ' Introducer(33);BlockID(249);BlockLen(4)
            c_DIBarray(aPtr + 3) = 1                 ' transparency flag
            ' skip next 2 bytes which is the Delay time; not needed for single frame
            c_DIBarray(aPtr + 6) = .TransIndex
            ' skip next byte which is the block terminator - zero
            aPtr = aPtr + 8 ' next position in array
        End If
        ' add the image description block
        c_DIBarray(aPtr) = 44 ' BlockID
        ' Left and Top coords are unnecessary for single frame; leave zeros
        CopyMemory c_DIBarray(aPtr + 5), c_DIBarray(6), &H4 ' width & height
        ' the packed byte would be next; but we'll leave it as zero
        ' since frame is not interlaced and not using a local color table.
        c_DIBarray(aPtr + 10) = 8 ' LZW compression size; 1st byte of compressed image
        aPtr = aPtr + 11 '10=img description block size + LZW compression byte
        ' next comes compressing DIB into LZW sub blocks, then finishing off stream
    End With
    
    'Initialize Masks -- Init LUT for fast 2 ^ x - 1 (was InitMasks routine)
    ReDim c_aPOT(0 To 16)
    c_aPOT(0) = 0
    For LIdx = 1 To 16
        c_aPOT(LIdx) = 2 * (c_aPOT(LIdx - 1) + 1) - 1
    Next LIdx
    
    '-- Reset output buffer values
    c_BitBucketBuff = 0
    c_BitBucketCount = 0
    ReDim c_aBuff(0 To 254) ' (was pvCharInit)

    '-- Set up the necessary startup values
    m_lInitBits = tBMP.bmBitsPixel + 1 '(DIB bit depth + 1)
    c_lCurrentBits = m_lInitBits
    m_lMaxCode = c_aPOT(c_lCurrentBits)
    m_lClearCode = PALETTECOUNT     ' 2^(m_lInitBits - 1)
    m_lCodeCount = PALETTECOUNT + 2 ' 2^(m_lInitBits - 1) + 2
    
    '-- Set hash code range bound for shifting
    lFCode = TABLE_SIZE
    Do While lFCode < 65536
        m_lShift = m_lShift + 1
        lFCode = lFCode + lFCode
    Loop
    m_lShift = 1 + c_aPOT(8 - m_lShift)
    'Added by LaVolpe -- quick erase, setting all table entries to -1 (was pvClearTable)
    FillMemory m_lHashTable(0), TABLE_SIZE * 4, 255  ' clear hash table
    
    '-- Start...
    c_lSubBlockSize = 1 ' position for 1st byte in data sub block
    Call pvOutputCode(m_lClearCode, aPtr) ' all images begin with a clear table flag
    
    ' start LZW patterns & also start looping on 2nd pixel
    lEnt = dibBytes(c_Frames(Index).imgOffset, tBMP.bmHeight - 1)
    dibCol = c_Frames(Index).imgOffset + 1
    
    'Added by LaVolpe -- looping is my enhancement;
    ' hash algorithm reorganized, comments added, modified very little
    For dibRow = tBMP.bmHeight - 1 To tBMP.bmHeight - c_Frames(Index).Dimensions.Bottom Step -1
    
        ' process each pixel in line of image
        For dibCol = dibCol To c_Frames(Index).imgOffset + c_Frames(Index).Dimensions.Right - 1
        
            lC = dibBytes(dibCol, dibRow)
            
            lFCode = lC * MAX_BITSHIFT + lEnt   ' add to existing pattern
            LIdx = (lC * m_lShift) Xor lEnt     ' XOR hashing
    
            If LIdx >= TABLE_SIZE Then LIdx = 0 ' added by LaVolpe (sanity check)
            
            If (m_lHashTable(LIdx) = lFCode) Then   ' found existing pattern
                lEnt = m_lCodeTable(LIdx)
            Else
                If (m_lHashTable(LIdx) > -1) Then ' else Empty slot
                    
                    lDisp = TABLE_SIZE - LIdx     ' Secondary hash (after G. Knott)
                    If (LIdx = 0) Then lDisp = 1
                    
                    Do  ' Hash Probing
                        LIdx = LIdx - lDisp
                        If (LIdx < 0) Then LIdx = LIdx + TABLE_SIZE
        
                        If (m_lHashTable(LIdx) = lFCode) Then
                            lEnt = m_lCodeTable(LIdx)
                            bExistingPattern = True
                            Exit Do
                        End If
        
                    Loop While (m_lHashTable(LIdx) > 0) ' Continue probing
                End If
                
                If bExistingPattern = True Then
                    bExistingPattern = False    ' reset flag
                Else
                    Call pvOutputCode(lEnt, aPtr)    ' write previous pattern
                    lEnt = lC  ' set current palette index as previous pattern
                    m_lCodeTable(LIdx) = m_lCodeCount ' store code & hash index
                    m_lHashTable(LIdx) = lFCode
                    
                    ' check for LZW compression increments
                    If m_lCodeCount > m_lMaxCode Then
                        ' ran out of codes for current compression size
                        If (m_lCodeCount = MAX_CODE) Then
                            ' add clear code to output stream (was pvClearBlock)
                            Call pvOutputCode(m_lClearCode, aPtr)
                            FillMemory m_lHashTable(0), TABLE_SIZE * 4, 255  ' clear hash table
                            m_lCodeCount = PALETTECOUNT + 1 ' reset code counter to 1 less 'cause it is incremented right away
                            c_lCurrentBits = m_lInitBits    ' reset baseline LZW compression size
                        Else
                            c_lCurrentBits = c_lCurrentBits + 1 ' increment compression size
                        End If
                        m_lMaxCode = c_aPOT(c_lCurrentBits) ' new max count for current compression size
                    End If
                    m_lCodeCount = m_lCodeCount + 1  ' increment the number of patterns
                End If
            End If
        Next dibCol
        dibCol = c_Frames(Index).imgOffset  ' reset to 1st pixel of line
    Next dibRow

    '--  Put out the final code & image data termination code
    Call pvOutputCode(lEnt, aPtr)
    ' finish off the stream
    Call pvOutputCode(PALETTECOUNT + 1, aPtr) ' add end of image data flag
    Call pvOutputCode(-1, aPtr)  ' flush remaining bytes in output bitbucket
    ' the next byte would be a sub block terminator which is zero; ignore & go on
    c_DIBarray(aPtr + 1) = 59   ' add end of file flag
        
    Set ConvertStripToGIF = PictureFromByteStream(c_DIBarray(), aPtr + 2)

EH:
    If Err Then
        Stop        ' for testing if error occurs
        Err.Clear
        Resume      ' for testing if error occurs
    End If
    
    ' clear global arrays, no longer used
    If TSA.cbElements = 1 Then CopyMemory ByVal VarPtrArray(dibBytes()), 0&, &H4
    Erase c_DIBarray()
    Erase c_aPOT()
    Erase c_aBuff()
    If bRestart Then Me.AnimationState = gfaPlaying ' continue animation if needed
End Function

Private Sub pvOutputCode(ByVal lCode As Long, ByRef arrayPtr As Long)

' MODIFIED/REORGANIZED AUTHOR'S ORIGINAL CODE (by Carles P.V.)
' - Also combined original pvCharOut and pvFlushChar routines herein

    If lCode < 0 Then
    
        If Not c_BitBucketCount = 0 Then ' add last bits of the output buffer to array
            c_aBuff(c_lSubBlockSize) = c_BitBucketBuff
            c_lSubBlockSize = c_lSubBlockSize + 1
        End If
        If Not c_lSubBlockSize = 1 Then ' still have bytes to write
            c_aBuff(0) = c_lSubBlockSize
            ' ensure array large enough for this block + 5 extra bytes needed to finish off stream
            If UBound(c_DIBarray) < arrayPtr + c_lSubBlockSize + 6 Then
                ReDim Preserve c_DIBarray(0 To arrayPtr + c_lSubBlockSize + 6)
            End If
            CopyMemory c_DIBarray(arrayPtr), c_aBuff(0), c_lSubBlockSize + 1
            arrayPtr = arrayPtr + c_lSubBlockSize + 1
        End If
    
    Else
        ' add latest code to the bitbucket & track total bit count
        c_BitBucketBuff = c_BitBucketBuff Or (lCode * (c_aPOT(c_BitBucketCount) + 1))
        c_BitBucketCount = c_BitBucketCount + c_lCurrentBits
    
        Do Until c_BitBucketCount < 8
            ' remove 8 bits at a time and place in output buffer (byte array)
            c_aBuff(c_lSubBlockSize) = (c_BitBucketBuff And &HFF&)
            c_BitBucketBuff = c_BitBucketBuff \ &H100&
            c_BitBucketCount = c_BitBucketCount - 8
            If (c_lSubBlockSize = 254) Then
                ' max allowable subblock data size is 255 byte blocks; write it
                c_aBuff(0) = c_lSubBlockSize
                If UBound(c_DIBarray) < arrayPtr + 260 Then
                    ' sanity check; unless the image is not compressible we
                    ' shouldn't trigger this Redim. ^^ 260=255+5 end of file/image bytes
                    ReDim Preserve c_DIBarray(0 To arrayPtr + 512)
                End If
                CopyMemory c_DIBarray(arrayPtr), c_aBuff(0), c_lSubBlockSize + 1
                arrayPtr = arrayPtr + c_lSubBlockSize + 1
                c_lSubBlockSize = 0
            End If
            ' keep track of next byte position to write to
            c_lSubBlockSize = c_lSubBlockSize + 1
        Loop
    End If

End Sub

Private Function PictureFromByteStream(aStream() As Byte, nrBytes As Long) As IPicture
  
  ' function creates a stdPicture (IPicture interface) from a byte array
  ' NOTE: This is a very unforgiving function. If the stream is not in the proper format
  '       the OleLoadPicture API will most likely lock up the application (GPF)
  ' -- Don't modify this routine or the 2 routines that call this function
  
  Dim IID_IPicture(15)
  Dim istm As IUnknown
  Dim hMem As Long, hPtr As Long
  
    hMem = GlobalAlloc(&H2, nrBytes)    ' reserve some system memory
    If Not hMem = 0 Then
        hPtr = GlobalLock(hMem)         ' get pointer to that memory
        If Not hPtr = 0 Then            ' copy local array to global array
            CopyMemory ByVal hPtr, aStream(LBound(aStream)), nrBytes
            Call GlobalUnlock(hMem)
            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                ' get the IPicture GUID
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                    ' if byte array contains invalid picture format, following locks up
                  Call OleLoadPicture(ByVal ObjPtr(istm), nrBytes, 0, IID_IPicture(0), PictureFromByteStream)
                End If
            End If
        End If
    End If

End Function

Private Function ParseGIF(ByVal FileName As String) As Long

    On Error Resume Next
    ' a modified routine from some of my other GIF postings
    ' This one is scaled back and skips many GIF blocks not needed for displaying.
    ' It also tracks a few file positions for later use in the BuildDIBstrip routine
    
    If Len(Dir$(FileName, vbArchive Or vbHidden Or vbReadOnly Or vbSystem)) = 0 Then Exit Function
    If FileLen(FileName) < 20 Then Exit Function
    If Err Then
        Err.Clear
        Exit Function
    End If
    
    Dim gByte As Byte                       ' general purpose Byte
    Dim gLong As Long                       ' general purpose Long
    Dim gString As String                   ' general purpose String
    Dim lFrameCount As Integer              ' nr frames in GIF
    Dim bGlobalColorTable As Boolean        ' does a global exist; if not, then possibly corrupt gif
    Dim cFrame As GIFframeProperties        ' frame data
    Dim emptyFrame As GIFframeProperties    ' overall GIF data
    
    On Error GoTo ExitReadRoutine
    
    ReDim c_DataLen(1 To 10)  ' used to track start position of image description & size of compressed data
    
    With c_gifProps
    
        c_FileNum = FreeFile()
        Open FileName For Binary Access Read As c_FileNum
        c_EOF = LOF(c_FileNum)
    
        ' read signature
        Call ReadGifFile_Variable(6)
        gString = LCase(StrConv(c_aBuff, vbUnicode))
        If Not (gString = "gif89a" Or gString = "gif87a") Then Exit Function
        
        ' skip to the global color table information
        Seek c_FileNum, 11
        gByte = ReadGifFile_Byte()             ' the packed byte
        ' GIF Logical Screen Descriptor packed byte per specs
            ' bit pos = 0: nr of bits = 1 ' global color table used
            ' bit pos = 1: nr of bits = 3 ' original resolution
            ' bit pos = 4: nr of bits = 1 ' table sorted
            ' bit pos = 5: nr of bits = 3 ' palette bit depth (can be 0 thru 7 inclusive)
        c_gifProps.BkgIndex = ReadGifFile_Byte() ' cache window bkg color index
        
        Seek c_FileNum, 14   ' move ahead to next byte after header
        If (gByte And 128) = 128 Then ' color table used? If so, read it
            c_ColorTables.Index = 0
            Call ReadGifFile_ColorTable(0, (gByte And &H7) + 1)
            bGlobalColorTable = True
        'Else no global color table; probably uses local color tables
        End If
    
    End With
    
    ' scan thru the entire file to find all the images & other key data
    With c_gifProps
    
        Do
            Select Case ReadGifFile_Byte()    ' read a single byte
            Case 0  ' block terminators
            Case 33 'Extension Introducer
                Select Case ReadGifFile_Byte() ' read the extension type
                
                Case 255    ' application extension, look for a loop count
                    ' the Netscape extension should always be before any images
                    ' Get the length of extension: will always be 11
                    gByte = ReadGifFile_Byte()
                    ' read next 11 bytes & see if it is a netscape extension
                    Call ReadGifFile_Variable(gByte)
                    gString = UCase(StrConv(c_aBuff, vbUnicode))
                    If UCase$(gString) = "NETSCAPE2.0" Then
                        ' within the data, we can extract the number of loops
                        ' an animated gif is suppose to run.
                        gByte = ReadGifFile_Byte()
                        If gByte = 3 Then   ' valid netscape extension
                            ' move ahead one byte & the next two is the loop count
                            Seek c_FileNum, Seek(c_FileNum) + 1
                            .Loops = (ReadGifFile_Integer And &HFFFF&) ' convert unsigned Integer to Long
                        Else
                            Seek c_FileNum, Seek(c_FileNum) - 1
                        End If
                    End If
                    SkipGifBlock
                
                Case 249    ' Graphic Control Label
                            ' (description of frame & is an optional block) 8 bytes
                    ' Graphic Control Extension (packed byte)
                        'bit pos = 0: nr of bits = 3 ' reserved
                        'bit pos = 3: nr of bits = 3 ' disposal method
                        'bit pos = 6: nr of bits = 1 ' user input flag
                        'bit pos = 7: nr of bits = 1 ' transparency
                    
                    cFrame = emptyFrame ' begin a new frame structure
                    With cFrame
                        .imgOffset = Seek(c_FileNum) - 2 ' image starts here in the file
                        gByte = ReadGifFile_Byte() ' skip static byte (block size of 4)
                        ' get next byte which contains several properties
                        gByte = ReadGifFile_Byte()
                        
                        ' how is frame cleared after it is shown
                        .Disposal = ((gByte \ &H4) And &H3)
                        If .Disposal = 3 Then c_gifProps.Disposal3 = 1
                        
                        ' how long does frame stay before being disposed & make 1/1000sec vs 1/100sec
                        .Delay = (ReadGifFile_Integer And &HFFFF&) * 10 ' convert unsigned Integer to Long
                        
                        .IsTransparent = (gByte And &H1) ' has transparency?
                        If .IsTransparent = 1 Then
                            .TransIndex = ReadGifFile_Byte  ' transparent index
                        Else
                            Seek c_FileNum, .imgOffset + 7  ' skip transindex byte
                        End If
                        ' skip rest of block
                        Call SkipGifBlock
                    End With
                    
                Case Else   ' Comment block, plain text extension, or Unknown extension
                    Call SkipGifBlock
                End Select
                
            Case 44 ' Image Descriptor (image dimensions & color table)
                ' location of image within logical window
                ' GIF Image Descriptor (packed byte)
                    'bit pos = 0: nr of bits = 1 ' local color table used
                    'bit pos = 1: nr of bits = 1 ' interlaced
                    'bit pos = 2: nr of bits = 1 ' table sorted
                    'bit pos = 3: nr of bits = 2 ' reserved
                    'bit pos = 5: nr of bits = 3 ' palette bit depth
                With cFrame
                    ' mark position where image description starts
                    c_DataLen(lFrameCount + 1).Y = Seek(c_FileNum) - 1
                    ' does image start here or did it start with Block 249 above?
                    If .imgOffset = 0 Then .imgOffset = c_DataLen(lFrameCount + 1).Y
                    
                    ' convert unsigned integers to long
                    .Dimensions.Left = (ReadGifFile_Integer And &HFFFF&)
                    .Dimensions.Top = (ReadGifFile_Integer And &HFFFF&)
                    .Dimensions.Right = (ReadGifFile_Integer And &HFFFF&)
                    .Dimensions.Bottom = (ReadGifFile_Integer And &HFFFF&)
                    
                    ' next byte indicates if local color table used
                    gByte = ReadGifFile_Byte()
                    If (gByte And 128) = 128 Then   ' local color table used?
                        .TblIndex = c_ColorTables.Index + 1 ' update ref & read table
                        Call ReadGifFile_ColorTable(.TblIndex, (gByte And &H7) + 1)
                        c_ColorTables.Index = .TblIndex     ' update table count
                    Else
                        If bGlobalColorTable = False Then Exit Function ' corrupted gif
                        ' frame says to use a global table, but no global color table exists
                    End If
                    ' skip the LZW byte & move to 1st byte of image
                    gByte = ReadGifFile_Byte()
                    SkipGifBlock ' done with block
                End With
                
                ' calculate image size in compressed bytes (includes local table if used)
                lFrameCount = lFrameCount + 1
                c_DataLen(lFrameCount).X = Seek(c_FileNum) - c_DataLen(lFrameCount).Y
                If c_DataLen(lFrameCount).X < 3 Then ' then invalid image data
                    lFrameCount = lFrameCount - 1 ' roll back the frame
                Else
                    ReDim Preserve c_Frames(1 To lFrameCount)
                    If lFrameCount = UBound(c_DataLen) Then
                        ReDim Preserve c_DataLen(1 To lFrameCount + 5)
                    End If
                    c_Frames(lFrameCount) = cFrame
                    ' ensure logical window large enough for every frame
                    With c_Frames(lFrameCount).Dimensions
                        If .Left + .Right > c_gifProps.Width Then c_gifProps.Width = .Left + .Right
                        If .Top + .Bottom > c_gifProps.Height Then c_gifProps.Height = .Top + .Bottom
                    End With
                End If
                cFrame = emptyFrame ' always start with a new frame
                
            Case 59 ' Trailer (end of images)
                ' Although more images may exist, this flag tells us not to use any others
                ' finalize the last frame's end offset
                Exit Do
            Case Else
                ' shouldn't happen; abort with what we have
                Exit Do
            End Select
        Loop
          
    End With
    ' got this far? good to go
          
ExitReadRoutine:
    If Err Then Err.Clear
    If Not lFrameCount = 0 Then
        If lFrameCount > UBound(c_Frames) Then ReDim Preserve c_Frames(1 To lFrameCount)
    End If
    Erase c_aBuff()
    ParseGIF = lFrameCount
    
End Function

'/==================================================================================
'   Read thru bytes until a zero-byte Block Terminator is found
'/==================================================================================
Private Sub SkipGifBlock()
    Dim curByte As Byte
    curByte = ReadGifFile_Byte()         ' current byte value
    Do While curByte > 0
        Seek c_FileNum, Seek(c_FileNum) + curByte
        curByte = ReadGifFile_Byte()
    Loop
End Sub
'/==================================================================================
'   Read one byte from an open file
'/==================================================================================
Private Function ReadGifFile_Byte() As Byte

    If Seek(c_FileNum) > c_EOF Then
        Err.Raise 53, "ReadGifFile", "End of File"
        Exit Function
    End If
    Get c_FileNum, , ReadGifFile_Byte
End Function
'/==================================================================================
'   Read an Integer (2 byte) value from an open file
'/==================================================================================
Private Function ReadGifFile_Integer() As Integer
    If Seek(c_FileNum) >= c_EOF Then
        Err.Raise 53, "ReadGifFile", "End of File"
        Exit Function
    End If
    Get c_FileNum, , ReadGifFile_Integer
End Function
'/==================================================================================
'   Read one or more bytes from the open gif file
'/==================================================================================
Private Sub ReadGifFile_Variable(ByVal nrBytes As Long)

    ReDim c_aBuff(0 To nrBytes - 1)
    If Seek(c_FileNum) + nrBytes - 1 > c_EOF Then GoTo EH
    Get c_FileNum, , c_aBuff
EH:
End Sub
'/==================================================================================
'   Reads color tables inside GIF file and updates class collection
'/==================================================================================
Private Sub ReadGifFile_ColorTable(ByVal TableSlot As Long, ByVal BitDepth As Long)
    Dim dBytes(0 To 2) As Byte, c As Long
    ReDim Preserve c_ColorTables.Tables(0 To PALETTECOUNT - 1, 0 To TableSlot)
    If Seek(c_FileNum) + c_aPOT(BitDepth) * 3 > c_EOF Then
        Err.Raise 53, "ReadGifFile", "End of File"
        Exit Sub
    End If
    For c = 0 To c_aPOT(BitDepth) - 1 ' convert RGB to BGR along the way
        Get c_FileNum, , dBytes()
        c_ColorTables.Tables(c, TableSlot) = (dBytes(0) * &H10000) Or (dBytes(1) * &H100&) Or dBytes(2)
    Next
End Sub

Private Function RenderFrame_Timer(ByVal Index As Long, ByVal hdc As Long, _
            ByVal hWnd2Refresh As Long, AutoRedraw As Boolean, _
            PostNotify As Boolean, bInternal As Boolean, _
            Optional ByVal destX As Long, Optional ByVal destY As Long, _
            Optional ByVal destCx As Long, Optional ByVal destCY As Long, _
            Optional ByVal StretchMode As ScaleGIFConstants) As Boolean

 ' Function renders a frame as a result of the class timer firing or as a
 ' result from the user calling the public RenderFrame function
 
 ' If called internally (from class timer), then
 '      The c_UserDCdata data (see SetAnimationBkg) is used to render frames
 ' Else the optional parameters are used and are passed from RenderFrame
 

    Dim eraseRect As RECT                   ' frame's bounding rect
    Dim refreshRect As RECT                 ' target bounding rect
    Dim gPIC As StdPicture                  ' current frame's image
    Dim hWnd2release As Long                ' used if getting DC from hWnd
    Dim bmpCacheOffset As Long              ' location of bkg image if any
    Dim tDC As Long                         ' target DC for rendering
    ' used if called from RenderFunction
    Dim xRatio As Single, yRatio As Single

    Select Case GetObjectType(hdc)          ' ensure hDC passed is a DC
    Case OBJ_DC, OBJ_MEMDC
    Case Else                               ' otherwise get it from the passed hWnd
        If IsWindow(hdc) = 0 Then Exit Function ' passed value isn't an hWnd either
        hWnd2release = hdc
        hdc = GetDC(hWnd2release)
    End Select
    
    On Error GoTo EH
    If bInternal = True Then ' safety check coming; shouldn't happen but...
        If c_UserDCdata.DC = 0 Then
        ' when user called SetAnimationBkg and if it failed but user ignored return
        ' value, we can get here. Also, in other cases, if every frame of the GIF
        ' (including single frames) are the same height/width, don't use transparency,
        ' and require no special Disposal handling, then the class will force this
        ' option vs creating and maintaining an unneeded offscreen DC & bitmap.
            bInternal = False
            destX = c_UserDCdata.Bounds.Left: destY = c_UserDCdata.Bounds.Top
            destCx = c_UserDCdata.Bounds.Right: destCY = c_UserDCdata.Bounds.Bottom
            StretchMode = c_UserDCdata.Mode
        End If
    End If
    
    If bInternal Then
        StretchMode = c_UserDCdata.Mode
        ' location and size of scaled frame within our memory DC
        With c_Frames(Index).Dimensions
            If (StretchMode And Not gfsCentered) = 0 Then
                ' clipping vs scaling; build a clipping rectangle
                eraseRect = c_Frames(Index).Dimensions
                If .Right + .Left > c_UserDCdata.Bounds.Right Then
                    eraseRect.Right = c_UserDCdata.Bounds.Right - .Left
                End If
                If .Bottom + .Top > c_UserDCdata.Bounds.Bottom Then
                    eraseRect.Bottom = c_UserDCdata.Bounds.Bottom - .Top
                End If
            Else ' build a scaled rectangle
                SetRect eraseRect, Int(.Left * c_UserDCdata.ScaleCx), Int(.Top * c_UserDCdata.ScaleCy), _
                    Int(.Right * c_UserDCdata.ScaleCx), Int(.Bottom * c_UserDCdata.ScaleCy)
            End If
        End With
        ' if Disposal=3, this frame must replace current DC contents when it is removed
        If c_Frames(Index).Disposal = 3 Then
            ' set offset in offscreen bmp where dest bkg image can be found
            If c_UserDCdata.Brush = 0 Then
                bmpCacheOffset = c_UserDCdata.Bounds.Bottom * 2
            Else
                bmpCacheOffset = c_UserDCdata.Bounds.Bottom
            End If
            With eraseRect
                BitBlt c_UserDCdata.DC, 0&, bmpCacheOffset, .Right, .Bottom, c_UserDCdata.DC, .Left, .Top, vbSrcCopy
            End With
        End If
        tDC = c_UserDCdata.DC
    Else            ' called from RenderFunction, not the class timer
        tDC = hdc
        With c_Frames(Index).Dimensions
            ScaleToDestination Index, destX, destY, destCx, destCY, StretchMode, xRatio, yRatio
            SetRect eraseRect, destX, destY, destCx, destCY
        End With
    End If
        
    ' if bInternal=True, then frame drawn to memory DC else drawn to target DC
    TransferFrame Index, tDC, (Not (StretchMode And Not gfsCentered) = 0), eraseRect, bInternal
    If bInternal = True Then
        With c_UserDCdata ' draw frame from memory DC to target DC (flicker free)
            BitBlt hdc, .Bounds.Left, .Bounds.Top, .Bounds.Right, .Bounds.Bottom, .DC, 0, 0, vbSrcCopy
        End With
    End If
        
    If Not hWnd2release = 0 Then ReleaseDC hWnd2release, hdc ' release dc if needed
        
    If bInternal = True Then
        ' handle disposal method: disposal codes per GIF documentation:
        ' 0 = no action required
        ' 1 = do not modify the graphic
        ' 2 = erase the graphic with background (current frame's dimensions)
        ' 3 = replace with previous graphic (current frame's dimensions)
        ' 4-7 = reserved; treated as disposal code zero by this routine
        Select Case c_Frames(Index).Disposal
        Case 2     ' remove the frame we just drew; done in back buffer; not target
            If c_UserDCdata.Brush = 0 Then
                With eraseRect
                    BitBlt c_UserDCdata.DC, .Left, .Top, .Right, .Bottom, c_UserDCdata.DC, .Left, c_UserDCdata.Bounds.Bottom + .Top, vbSrcCopy
                End With
            Else    ' solid bkg color used
                With c_Frames(Index).Dimensions
                    eraseRect.Right = Int((.Left + .Right) * c_UserDCdata.ScaleCx)
                    eraseRect.Bottom = Int((.Top + .Bottom) * c_UserDCdata.ScaleCy)
                End With
                FillRect c_UserDCdata.DC, eraseRect, c_UserDCdata.Brush
            End If
        Case 3     ' replace previous DC contents from back buffer
            With eraseRect
                BitBlt c_UserDCdata.DC, .Left, .Top, .Right, .Bottom, c_UserDCdata.DC, 0&, bmpCacheOffset, vbSrcCopy
            End With
        Case Else:  'no action needed or image must remain
        End Select
    End If
    
    If Not hWnd2Refresh = 0 Then ' refresh updated image if requested
        ' 1st identify the area to be refreshed
        If bInternal Then eraseRect = c_UserDCdata.Bounds
        With eraseRect
            SetRect refreshRect, .Left, .Top, .Left + .Right, .Bottom + .Top
        End With
        If AutoRedraw = True Then
            ' calling Me.Refresh on a DC with AutoRedraw=True, replaces the DC with a
            ' new one. This API will refresh the DC w/o that happening; faster overall.
            InvalidateRect hWnd2Refresh, refreshRect, 0&
        Else
            RedrawWindow hWnd2Refresh, refreshRect, 0&, &H108 ' update now & validate
        End If
    End If
    
    If PostNotify = True Then
        If Not c_IFace Is Nothing Then Call c_IFace.Rendered(c_UserID, Index, msgRendered, hdc)
    End If
    
    RenderFrame_Timer = True
EH:
End Function

Private Function IsArrayEmpty(ByVal lArrayPointer As Long) As Boolean
  ' test to see if an array has been initialized
  IsArrayEmpty = (lArrayPointer = -1)
End Function

Private Function BuildDIBstrip() As Boolean
    
    ' Function creates one DIB (a strip of frames as a single bitmap):
    
    ' The process I use creates a data stream (bytes) containing the GIF format
    ' for each frame of the GIF. That stream is sent to an API to convert stream to
    ' a stdPicture for each frame then render the stdPicture to the DIB strip.
    ' This is up to 3 times faster than decompressing the GIF frames by hand but has
    ' a drawback which is addressed. What's the drawback?
    ' Pixel palette indexes can change during stdPicture creation:
    
    ' ok; thru trial & error, I found that the stdPicture will rewrite the
    ' GIF in certain cases. This happens when a palette color is repeated.
    ' In that case, the GIF will be re-written to use the 1st palette index that
    ' references that repeated color. When transparency is used, this prevents me
    ' from using masks on my DIB strip because the transparency index I parsed from
    ' the GIF file may no longer be valid if that index was NOT the only instance of
    ' that color. So a work around is to ensure every palette index is unique. This,
    ' of course, is not an issue if each frame was stored as a stdPicture because the
    ' IPicture interface (stdPicture) must keep that info for its use. But I want
    ' to use one GDI resource vs potentially dozens upon dozens of stdPictures.
    
    Dim maxDataLen As Long, maxHeight As Long, totalWidth As Long
    Dim bTransparency As Boolean
    Dim f As Long, tDC As Long, aPtr As Long
    Dim tPic As StdPicture, stripBMP As BITMAPINFO
    ' used to ensure palette index uniqueness
    Dim uniquePal(0 To 767) As Byte
    ' used to upload GIF file data into the byte stream
    Dim fileData() As Byte, imgStream() As Byte
    
    On Error GoTo EH
    ' minimal image size optimizing used: assuming most GIF frames will be of similar
    ' height and width, calculate the maximum height of all frames and the
    ' total width of all frames. Our DIB strip will be horizontal
    For f = 1 To UBound(c_Frames)
        With c_Frames(f)
            totalWidth = totalWidth + .Dimensions.Right ' total width
            If .Dimensions.Bottom > maxHeight Then
                maxHeight = .Dimensions.Bottom          ' calc max height
            End If
            If .IsTransparent Then bTransparency = True ' used to initialize mask array
            ' guesstimate size of byte stream needed to create stdPic GIFs
            If c_DataLen(f).X > maxDataLen Then maxDataLen = c_DataLen(f).X
        End With
    Next
    
    With stripBMP.bmiHeader            '  do the color dib strip
        .biBitCount = 8
        .biClrUsed = PALETTECOUNT
        .biHeight = maxHeight
        .biWidth = totalWidth
        .biPlanes = 1
        .biSize = Len(stripBMP.bmiHeader)
    End With
    With stripBMP
        For f = 1 To PALETTECOUNT - 1   ' ensure each palette entry is used only once
            uniquePal(f * 3 + 2) = f    ' RGB (byte) - used by GIF
            .bmiPalette(f) = f          ' BGR (long) version - used by DIB
        Next
    End With
    tDC = GetDC(0&)                    ' get DC to create GIFs from
    c_DIB = CreateDIBSection(tDC, stripBMP, 0&, ByVal 0&, 0&, 0&) ' create the DIB
    If c_DIB = 0 Then ' Houston; we have a problem! Can't display the GIF with this class
        ReleaseDC 0&, tDC
        Exit Function
    End If

    ' create a DC & select our strip into it
    c_DC = CreateCompatibleDC(tDC)
    ReleaseDC 0&, tDC
    c_DIB = SelectObject(c_DC, c_DIB)
    c_ColorTables.Index = -1 ' will force 1st rendering in TransferFrame to update palette
    
    ' Populate our DIB strip... using stdPictures because it is up to 3 times faster than
    ' manually decompressing the GIF using VB algorithms alone. Just requires some
    ' extra work because of a quirk with how stdPicture can rewrite a GIF frame.
    
    ReDim imgStream(1 To maxDataLen + 790) ' oversize array to handle any frame
    '^^ 790 = 768 for global palette;13 for header;8 for image control block;1 for EOF
    ' the array size IS EXACTLY correct when an 8-bit, transparent GIF file uses
    ' both global & local tables -- DO NOT reduce its size at all
    ' FYI: GIF header Layout?
    '  1-6   Signature (i.e., GIF87a or GIF89a)
    '  7-10  Logical Window - overall dimensions (width/height, 2 bytes each)
    '  11    Packed byte describing global color table
    '  12    Background Window color index
    '  13    Pixel aspect ratio
    '  14-xx Global color table if it exists
    ReDim fileData(1 To 13)                     ' get logical window info & signature
    Get c_FileNum, 1, fileData()
    CopyMemory imgStream(1), fileData(1), 13    ' copy logical window into stream
    CopyMemory imgStream(14), uniquePal(0), PALETTECOUNT * 3 ' copy our unique palette into the global
    imgStream(11) = imgStream(11) Or 135        ' 135=our global exists & is 256 colors
    
    ' create an all-black mask array
    If bTransparency = True Then ReDim c_maskTable(0 To PALETTECOUNT - 1)
    
    For f = 1 To UBound(c_Frames)
        aPtr = 782          ' next position in stream
        With c_Frames(f)
            ' the c_DataLen() array elements are set in the ParseGIF routine
            If Not c_DataLen(f).Y = c_Frames(f).imgOffset Then
                ' get the image control block (always 8 bytes) & copy to stream
                ' FYI: GIF control block layout? (n/a for v87a)
                ' 1     Extension Introducer (fixed at 33)
                ' 2     Block ID (fixed at 249)
                ' 3     Remainng bytes in block after this byte (fixed at 4)
                ' 4     Transparency & other v89a options (packed byte)
                ' 5-6   Delay time in hundredths of seconds
                ' 7     Transparency Index
                ' 8     Block terminator (fixed at 0)
                ReDim fileData(1 To 8)
                Get c_FileNum, c_Frames(f).imgOffset, fileData()
                CopyMemory imgStream(aPtr), fileData(1), 8
                imgStream(aPtr + 3) = (imgStream(aPtr + 3) And Not 1)
                '^^ turn transparnecy off; otherwise, stdPicture will not write
                '   the correct index when it is rendered to my DIB strip
                aPtr = aPtr + 8
            End If
            ' get the image description block (always 10 bytes), and the
            '   local color table (0 to 768 bytes) & compressed image (variable length)
            ' FYI: GIF image description block
            ' 1     Block ID (fixed at 44)
            ' 2-5   Frame left/top offsets, 2 bytes each
            ' 6-9   Frame width/height, 2 bytes each
            ' 10    Packed byte describing local color table
            ' 11-xx Local color table, if used
            ' xxxxx compressed image data where its 1st byte is the LZW compression size
            ReDim fileData(1 To c_DataLen(f).X)
            Get c_FileNum, c_DataLen(f).Y, fileData()
            CopyMemory imgStream(aPtr), fileData(1), c_DataLen(f).X
            ' ^^ includes EOI (end of image flag) & block terminator
            '       ensure logical window is at least as big as this frame (some GIFs
            '       are corrupted this way). First zero out Top/Left - not needed
            CopyMemory imgStream(aPtr + 1), 0&, 4
            ' now copy frame's width/height to logical window width/height
            CopyMemory imgStream(7), imgStream(aPtr + 5), 4
            If Not .TblIndex = 0 Then
                ' the local table starts 10 bytes after block starts & the byte before that
                ' one tells how many colors in the table; replace with our unique palette
                CopyMemory imgStream(aPtr + 10), uniquePal(0), _
                            c_aPOT((imgStream(aPtr + 9) And &H7) + 1) * 3
            End If
            aPtr = aPtr + c_DataLen(f).X    ' calc total bytes in this frame
            imgStream(aPtr) = 59            ' add an EOF flag
            
            ' identify where in the strip this frame will begin
            Set tPic = PictureFromByteStream(imgStream(), aPtr)
            If f = 1 Then
                .imgOffset = 0
            Else
                .imgOffset = c_Frames(f - 1).imgOffset + c_Frames(f - 1).Dimensions.Right
            End If
            ' let's create the stdPicture preserving all index values; actual colors
            ' are not important at this step, but palette indexes are very important.
            If Not tPic Is Nothing Then
                ' blt the stdPicture to our strip at the correct position
                tPic.Render c_DC + 0&, .imgOffset + 0&, 0&, .Dimensions.Right + 0&, .Dimensions.Bottom + 0&, 0&, tPic.Height, tPic.Width, -tPic.Height, ByVal 0&
                ' ^^ FYI: the "+ 0&" above are required else errors
            End If
        End With
        If f = 1 Then
            ' when first frame is finished, we can display it if IGifRender interface has been established
            If Not c_IFace Is Nothing Then
                Call TimerProc(0, 0, 0, 0)
                SetDIBColorTable c_DC, 0, 256, stripBMP.bmiPalette(0)
                c_ColorTables.Index = -1 ' will force 1st rendering in TransferFrame to update palette
            End If
        End If
        ' send a progress message
        If Not c_IFace Is Nothing Then
            c_IFace.Rendered c_UserID, f, msgProgress, Int((f / UBound(c_Frames)) * 100)
        End If
    Next
    
    BuildDIBstrip = True
      
EH:
    If Err Then Err.Clear    'failure; calling routine will ensure memory objects are cleared
End Function

Private Sub TransferFrame(ByVal fNr As Long, ByVal hdc As Long, _
                ByVal bStretch As Boolean, destR As RECT, bInternal As Boolean)

    ' function Blts the gif frame onto the passed DC while scaling
    ' It also creates a mask on the fly using a niffty color table tweak.
    ' Majority of code is keeping track of color tables during the task
    
    ' when clipping, a frame can be completely outside of the DC; simply abort
    If destR.Bottom < 1 Or destR.Right < 1 Then Exit Sub
    
    Dim mRop As Long
    Dim tOldBmp As Long
    Dim lTransColor As Long
    Dim oldBrEx As POINTAPI, oldStretchMode As Long
        
    With c_Frames(fNr)
        If bInternal = False Then ' set target DC stretch mode & cache original setting
            oldStretchMode = SetStretchBltMode(hdc, HALFTONE)
            SetBrushOrgEx hdc, 0, 0, oldBrEx
        End If
        
        If .IsTransparent = 1 Then ' color table hack to create masks for paletted images
            ' The c_maskTable is all black, no other colors
            c_maskTable(.TransIndex) = vbWhite    ' set transparent index to white
            SetDIBColorTable c_DC, 0, PALETTECOUNT, c_maskTable(0) ' put the table to the DIB
            c_maskTable(.TransIndex) = vbBlack ' set the color back to black
            ' now draw the mask to the destination DC, note the ROP used >>
            If bStretch = True Then
                StretchBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, _
                    c_DC, .imgOffset, 0&, .Dimensions.Right, .Dimensions.Bottom, vbSrcAnd
            Else    ' clipping only
                BitBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, c_DC, .imgOffset, 0&, vbSrcAnd
            End If
            mRop = vbSrcPaint       ' set ROP for the color portion
        Else
            mRop = vbSrcCopy        ' no transparency used; default ROP
        End If
        
        ' do we need to select a different table for the current frame?
        If .IsTransparent = 1 Or Not .TblIndex = c_ColorTables.Index Then
            If .IsTransparent = 1 Then
                ' cache original transparency color & change it to black
                lTransColor = c_ColorTables.Tables(.TransIndex, .TblIndex)
                c_ColorTables.Tables(.TransIndex, .TblIndex) = vbBlack
            End If
            c_ColorTables.Index = .TblIndex  ' update current table ref
            ' change the dib colors & replace transparency color
            SetDIBColorTable c_DC, 0, PALETTECOUNT, c_ColorTables.Tables(0, .TblIndex)
            If .IsTransparent = 1 Then c_ColorTables.Tables(.TransIndex, .TblIndex) = lTransColor
        End If
        
        ' blt the color image and done
        If bStretch = True Then
            StretchBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, _
                c_DC, .imgOffset, 0&, .Dimensions.Right, .Dimensions.Bottom, mRop
        Else
            BitBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, c_DC, .imgOffset, 0&, mRop
        End If
        
        If bInternal = False Then
            ' reset to user DC's original settings
            SetStretchBltMode hdc, oldStretchMode
            SetBrushOrgEx hdc, oldBrEx.X, oldBrEx.Y, oldBrEx
        End If
    End With
    
End Sub

Private Function ByteAlignOnWord(ByVal BitDepth As Byte, ByVal Width As Long) As Long
    ' generic function to align byte range on dWord boundaries for any bit depth
    ByteAlignOnWord = (((Width * BitDepth) + &H1F) And Not &H1F&) \ &H8
End Function

'*************************************************************************************************
'* cCallback - Class generic callback template
'*
'* Note:
'*  The callback declarations and code are exactly the same for a Class, Form or UserControl.
'*  The callback declarations and code can co-exist with subclassing declarations and code.
'*    With both types of code in a single file,..
'*      delete the duplicated declarations and code, Ctrl+F5 will find them for you
'*      pay careful attention to the nOrdinal parameter to zAddressOf
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 The original..................................................................... 20060408
'* v1.1 Added multi-thunk support........................................................ 20060409
'* v1.2 Added optional IDE protection.................................................... 20060411
'* v1.3 Added an optional callback target object......................................... 20060413
'*************************************************************************************************

'-Callback code-----------------------------------------------------------------------------------
Private Function cb_AddressOf(ByVal nOrdinal As Long, _
                              ByVal nParamCount As Long, _
                     Optional ByVal nThunkNo As Long = 0, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
'*************************************************************************************************
'* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
'* nParamCount  - The number of parameters that will callback
'* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
'* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety   - Optional, set to false to disable IDE protection.
'*************************************************************************************************
Const MAX_FUNKS   As Long = 1                                               'Number of simultaneous thunks, adjust to taste
Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
  Dim nAddr       As Long
  
  If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
    MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
  If nAddr = 0 Then
    MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If z_CbMem = 0 Then                                                       'If memory hasn't been allocated
    ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
    z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
  End If
  
  If z_Cb(0, nThunkNo) = 0 Then                                             'If this ThunkNo hasn't been initialized...
    z_Cb(3, nThunkNo) = _
              GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
    z_Cb(4, nThunkNo) = &HBB60E089
    z_Cb(5, nThunkNo) = VarPtr(z_Cb(0, nThunkNo))                           'Set the data address
    z_Cb(6, nThunkNo) = &H73FFC589: z_Cb(7, nThunkNo) = &HC53FF04: z_Cb(8, nThunkNo) = &H7B831F75: z_Cb(9, nThunkNo) = &H20750008: z_Cb(10, nThunkNo) = &HE883E889: z_Cb(11, nThunkNo) = &HB9905004: z_Cb(13, nThunkNo) = &H74FF06E3: z_Cb(14, nThunkNo) = &HFAE2008D: z_Cb(15, nThunkNo) = &H53FF33FF: z_Cb(16, nThunkNo) = &HC2906104: z_Cb(18, nThunkNo) = &H830853FF: z_Cb(19, nThunkNo) = &HD87401F8: z_Cb(20, nThunkNo) = &H4589C031: z_Cb(21, nThunkNo) = &HEAEBFC
  End If
  
  z_Cb(0, nThunkNo) = ObjPtr(oCallback)                                     'Set the Owner
  z_Cb(1, nThunkNo) = nAddr                                                 'Set the callback address
  
  If bIdeSafety Then                                                        'If the user wants IDE protection
    z_Cb(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
  End If
    
  z_Cb(12, nThunkNo) = nParamCount                                          'Set the parameter count
  z_Cb(17, nThunkNo) = nParamCount * 4                                      'Set the number of stck bytes to release on thunk return
  
  nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                   'Calculate where in the allocated memory to copy the thunk
  RtlMoveMemory nAddr, VarPtr(z_Cb(0, nThunkNo)), FUNK_LEN                  'Copy thunk code to executable memory
  cb_AddressOf = nAddr + 16                                                 'Thunk code start address
End Function

'Terminate the callback thunks
Private Sub cb_Terminate()
Const MEM_RELEASE As Long = &H8000&                                         'Release allocated memory flag

  If z_CbMem <> 0 Then                                                      'If memory allocated
    If VirtualFree(z_CbMem, 0, MEM_RELEASE) <> 0 Then                       'Release
      z_CbMem = 0                                                           'Indicate memory released
    End If
  End If
End Sub

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim I     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, I, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, I, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, I, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  I = I + 4                                                                 'Bump to the next entry
  J = I + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While I < J
    RtlMoveMemory VarPtr(nAddr), I, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    I = I + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

'*************************************************************************************************
'* Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
'*************************************************************************************************

'Callback ordinal 1
Private Function TimerProc(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
    
    ' Note. When hWnd is zero, this this routine is being called from LoadGIF to quick-display 1st frame
    If Not hWnd = 0 Then KillTimer c_tmrOwner, ObjPtr(Me)  ' stop current timer
    
    Dim tDC As Long, tHwnd As Long, bRestart As Boolean
    Dim bAutoRedraw As Boolean, bPostNotify As Boolean, bLoopComplete As Boolean
    
    ' determine next frame in the animation order
    c_curFrame = c_curFrame + 1
    If c_curFrame > UBound(c_Frames) Then
        c_curFrame = 1
        bLoopComplete = True
    End If
    ' get drawing DC and hWnd (if necessary) from client
    On Error GoTo EH
    Call c_IFace.GetRenderDC(c_UserID, c_curFrame, tDC, tHwnd, bAutoRedraw, bPostNotify)
    If Not tDC = 0 Then ' draw the frame
        RenderFrame_Timer c_curFrame, tDC, tHwnd, bAutoRedraw, bPostNotify, True
    End If
        
    If UBound(c_Frames) = 1 Then  ' single frame gif
        c_AniLoops = 1
        c_aniState = gfaStopped
    Else
        ' determine if timer should continue
        bRestart = True
        If bLoopComplete = True Then ' another loop finished
            If Not c_gifProps.Loops = 0 Then  ' has specified number of animation loops
                c_AniLoops = c_AniLoops + 1 ' increment number of loops completed
                If c_AniLoops = c_gifProps.Loops Then ' end of desired loops, no more animation
                    c_aniState = gfaStopped
                    bRestart = False
                End If
            End If
        End If
    End If
    
    If Not hWnd = 0 Then
        If bRestart = True Then ' set timer for current frame
            SetTimer hWnd, TimerID, FrameProperty(c_curFrame, gfpDelay), c_tmrAddressOf ' set timer
        Else                    ' send loop termination message
            Call c_IFace.Rendered(c_UserID, c_curFrame, msgLoopsEnded, c_AniLoops)
        End If
    End If

EH:
' CAUTION: DO NOT ADD ANY ADDITIONAL CODE OR COMMENTS PAST THE "END FUNCTION"
'          STATEMENT BELOW. Paul Caton's zProbe routine will read it as a start
'          of a new function/sub and the class timer's will fail every time.
End Function



