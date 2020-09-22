VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Truecolor frame delta compression
'Update 1

Dim Running&

Dim ScreenWidth&
Dim ScreenHeight&

Dim FPS_Max As Byte

Dim ScreenGrab As AnimSurf2D
Dim DeltaSurf As AnimSurf2D

Dim RawScaledBytesCount&
Dim BytesThisFrame&

Dim X%
Dim Y%

Private Type X_RGB_Values
 X_POS      As Integer
 BytRed    As Byte
 BytGreen  As Byte
 BytBlue   As Byte
End Type

Private Type VStackType
 Actual_Y As Integer
 H_Width  As Integer
End Type

Dim AbacusBits() As X_RGB_Values
Dim VStack() As VStackType

Dim DeltaScanLines&

Dim DeskDC&

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Translation_X() As Long
Private Translation_Y() As Long

Private Sub Form_Load()
Dim Tick&
Dim NextTick&

 'Server does everything below here
 ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
 ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
 
 MakeSurf ScreenGrab, ScreenWidth, ScreenHeight
 
 ScaleMode = vbPixels
 
 DeskDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
 
 Show
 DoEvents
 
 FPS_Max = 10
 
 Do While DoEvents
  Tick = GetTickCount
  If Tick > NextTick Then
   Encode
   DecodeAndRender
   If FPS_Max > 0 Then
    NextTick = Tick + 1000 / FPS_Max
   Else
    NextTick = Tick
   End If
  End If
 Loop
 
 DeleteDC DeskDC
 
 Erase Translation_X
 Erase Translation_Y
 
 ClearSurface ScreenGrab
 ClearSurface DeltaSurf
 
 Unload Me
 
End Sub
Private Sub Encode()
Dim WidthStackPointer% '% =  As Integer
Dim TmpXptr&
Dim Y_SRC&
Dim BGRA&
Dim singleX_SRC!
Dim i_singleX!
 
 BitBlt ScreenGrab.mem_hDC, 0, 0, _
        ScreenGrab.Dims.Wide, ScreenGrab.Dims.High, _
        DeskDC, 0, 0, SRCCOPY
        
 'each frame reserves 4 bytes to tell how many bytes
 'are used to create the frame
 BytesThisFrame = 4
 
 DeltaScanLines = -1
 
 For Y = 0 To DeltaSurf.Dims.HighM1
 
  'scaled y for desktop to window
  Y_SRC = Translation_Y(Y) 'Int(Y * ScreenGrab.Dims.Height / DeltaSurf.Dims.Height)
  
  'set number of pixels that have changed on this scanline
  WidthStackPointer = -1
  
  'this section updates 'server' scaled buffer and begins
  'the encode process along scanline
  For X = 0 To DeltaSurf.Dims.WideM1
   BGRA = ScreenGrab.LDib(Translation_X(X), Y_SRC)
   If BGRA <> DeltaSurf.LDib(X, Y) Then
    WidthStackPointer = WidthStackPointer + 1
    With AbacusBits(WidthStackPointer, Y)
     'send 2 bytes: where along scanline did change occur
     .X_POS = X
     'send 3 bytes: red, green and blue
     .BytRed = (BGRA And &HFF0000) / 65536
     .BytGreen = (BGRA And &HFF00&) / 256
     .BytBlue = BGRA And &HFF
    End With
   'keeping track
    BytesThisFrame = BytesThisFrame + 5
   End If
  Next X
  
  'this section 'sends' data if any pixels change on
  'this scanline
  If WidthStackPointer > 0 Then
   DeltaScanLines = DeltaScanLines + 1
   'going to send 2 bytes for which scanline has
   'new pixels, and 2 bytes for how many pixels
   'along scanline are new
   VStack(DeltaScanLines).Actual_Y = Y
   VStack(DeltaScanLines).H_Width = WidthStackPointer
   BytesThisFrame = BytesThisFrame + 4
  End If
  
  If RawScaledBytesCount < BytesThisFrame Then
   'encoding so far has produced a bigger result
   'than sending raw, so abort
   Exit For
  End If
  
 Next Y
 
 
 If RawScaledBytesCount < BytesThisFrame Then
 
  'delta formatting has exceeded 'raw', so switching
  'to 'raw'
   
  For Y = 0 To DeltaSurf.Dims.HighM1
   Y_SRC = Translation_Y(Y) 'Int(Y * ScreenGrab.Dims.Height / DeltaSurf.Dims.Height)
   For X = 0 To DeltaSurf.Dims.WideM1
    TmpXptr = Translation_X(X) 'Int(x * ScreenGrab.Dims.width / DeltaSurf.Dims.width)
    With AbacusBits(X, Y)
     .BytRed = ScreenGrab.Dib(TmpXptr, Y_SRC).Red
     .BytGreen = ScreenGrab.Dib(TmpXptr, Y_SRC).Green
     .BytBlue = ScreenGrab.Dib(TmpXptr, Y_SRC).Blue
    End With
   Next X
  Next Y
  
  Caption = RawScaledBytesCount & " bytes"
  
 Else
   
  Caption = BytesThisFrame & " bytes"
 
 End If
 
End Sub


Private Sub DecodeAndRender()
Dim X%
Dim Y%
Dim X_Track&
Dim Y_Track&
 
      '4 bytes
 If BytesThisFrame <= RawScaledBytesCount Then
 
  For Y_Track = 0 To DeltaScanLines '2 bytes
   Y = VStack(Y_Track).Actual_Y '2 bytes
   For X_Track = 1 To VStack(Y_Track).H_Width '2 bytes
    With AbacusBits(X_Track, Y)
    '2 bytes for .X_Pos
    '3 bytes for .BytBlue, .BytGreen, .BytRed
     DeltaSurf.LDib(.X_POS, Y) = .BytRed * 65536 Or .BytGreen * 256& Or .BytBlue
    End With
   Next X_Track
  Next Y_Track
  
 Else 'BytesThisFrame > RawScaledBytesCount .. render full
 
  For Y = 0 To DeltaSurf.Dims.HighM1
   For X = 0 To DeltaSurf.Dims.WideM1
    With AbacusBits(X, Y)
    '3 bytes
     DeltaSurf.LDib(X, Y) = RGB(.BytBlue, .BytGreen, .BytRed)
    End With
   Next X
  Next Y
  
 End If
 
 BlitToDC hdc, DeltaSurf
 
End Sub

Private Sub Form_KeyDown(IntKey As Integer, Shift As Integer)
 Select Case IntKey
 Case vbKeyEscape
  Running = False
 End Select
End Sub

Private Sub Form_Resize()
 
 RawScaledBytesCount = ScaleWidth * ScaleHeight * 3 + 4
 
 If RawScaledBytesCount > 4 Then
 
  MakeSurf DeltaSurf, ScaleWidth, ScaleHeight
  ReDim AbacusBits(DeltaSurf.Dims.WideM1, DeltaSurf.Dims.HighM1)
  ReDim VStack(DeltaSurf.Dims.HighM1)
  
  Erase Translation_X
  Erase Translation_Y
 
  ReDim Translation_X(DeltaSurf.Dims.WideM1)
  ReDim Translation_Y(DeltaSurf.Dims.HighM1)
 
  If DeltaSurf.Dims.HighM1 > 0 Then
  For Y = 0 To DeltaSurf.Dims.HighM1
   Translation_Y(Y) = Int(Y * ScreenGrab.Dims.HighM1 / DeltaSurf.Dims.HighM1)
  Next
  End If
  If DeltaSurf.Dims.WideM1 > 0 Then
  For X = 0 To DeltaSurf.Dims.WideM1
   Translation_X(X) = Int(X * ScreenGrab.Dims.WideM1 / DeltaSurf.Dims.WideM1)
  Next
  End If
 
 End If
 
End Sub


