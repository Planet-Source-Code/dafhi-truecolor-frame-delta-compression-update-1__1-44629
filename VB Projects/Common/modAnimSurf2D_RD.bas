Attribute VB_Name = "modAnimSurf2D_RD"
Option Explicit

'+------------------+---------------------------------+
'| modAnimSurf2D_RD | developed in Visual Basic 6.0   |
'+---------+--------+---------------------------------+
'| Release | 1.01                                     |
'+---------+-------+----------------------------------+
'| Original author | dafhi                            |
'+-------------+---+----------------------------------+
'| Description | Research & Development of meaningful |
'+-------------+ and useful 32-bit Dib-handling subs, |
'| functions, types, and variable naming conventions. |
'+----------------------------------------------------+

Public Type RGBQUAD
 Blue  As Byte
 Green As Byte
 Red   As Byte
 Alpha As Byte
End Type

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Type BITMAPINFOHEADER '40 bytes
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
Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Type PicBmp
    Size As Long
    Type As PictureTypeConstants
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs

Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLbound As Long
End Type

Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PointAPI
 X As Long
 Y As Long
End Type

Private Type DimsAPI
 Wide As Long
 High As Long
 WideM1 As Long
 HighM1 As Long
 bSRGB As Boolean
 bHSV As Boolean
End Type

Private Type PrecisionRGBA
 sRed As Single
 sGrn As Single
 sBlu As Single
 aLph As Single
End Type

Public Type HSVTYPE
 h_hue As Single
 s_saturation As Single
 v_value As Single
End Type

Private Type StartAndLenM1
 LStart As Long
 LFinal As Long
End Type

Private Type LayerConstruct
 UB_Lengths As Long
 LB_Lengths As Long
 Piece() As StartAndLenM1
 LocPtr As Long
 Alpha As Long
 AlphaOrigin As Long
End Type

Private Type Terraces2
 TerraceRef() As Long
 UB_Tr As Integer
 LB_Tr As Integer
 Terrace() As LayerConstruct
 Encoded As Boolean
End Type

Private Type RunLengthHeightField
 Reds As Terraces2
 Greens As Terraces2
 Blues As Terraces2
 Hues As Terraces2
 Sats As Terraces2
 Vals As Terraces2
 WhichAreEncoded As Long
End Type

Public Type AnimSurf2D
 Dims As DimsAPI            ' Width, Height, Width - 1, Height - 1
 wDiv2 As Single            ' Width / 2
 hDiv2 As Single            ' Height / 2
 TotalPixels As Long        ' Width * Height
 UB1D As Long               ' TotalPixels - 1
 TopLeft As Long
 Index As Long
 SafeAry2D As SAFEARRAY2D
 SafeAry2D_L As SAFEARRAY2D
 SafeAry1D As SAFEARRAY1D
 SafeAry1D_L As SAFEARRAY1D
 SafeAry1D_sRGB As SAFEARRAY1D
 mem_hDIb As Long
 mem_hBmpPrev As Long
 mem_hDC As Long
 BMinfo As BITMAPINFO
 Dib() As RGBQUAD           ' access 2D image as Bytes
 Dib1D() As RGBQUAD         ' 1D
 LDib() As Long             ' 2D Longs
 LDib1D() As Long           ' 1D
 sRGB() As PrecisionRGBA    ' high-depth color processing
 sRGB1D() As PrecisionRGBA  '
 HSV1D() As HSVTYPE
 HorizRLenSegs() As Integer   ' advanced custom blit architecture
 RLSeg() As Integer           ' advanced custom blit architecture
 BackColor As Long
 ClsnType As Long
 Clsn_Pattern() As Long
 EraseDib() As Long
 EraseStack() As Long
 EraseStackPtr As Long
 IA As RunLengthHeightField
End Type

Public Type PrecisionPointAPI
 sX As Single
 sY As Single
End Type

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle&, IPic As IPicture) As Long
Private LMem_hBmpPrev&

Public BackBuffer  As AnimSurf2D

Private SurfAuto(5000) As AnimSurf2D
Private SurfAutoUsed As Long

Public StrFileFolder$

Public Const BLIT_TYPE_RGB As Long = 0
Public Const BLIT_TYPE_ALPHA As Long = 1

Public Const RedLayer As Long = 1
Public Const GreenLayer As Long = 2
Public Const BlueLayer As Long = 3

Private Const LB1 As Long = -32768
Private Const UB1 As Long = 32767

Private Const LB1M1 As Long = LB1 - 1

Dim Layer&

Dim DrawX&
Dim DrawY&

Dim DrawRight&

Public BGRed&
Public BGGrn&
Public BGBlu&

Dim FGGrn&
Dim FGBlu&

Dim BGColor&
Dim FGColor&

Public hPub!
Public sPub!
Public vPub!

Public minim!
Public maxim!

Public sR!
Public sG!
Public sB!

Public subt!
Dim SHF_intensity!

Public rgb_shift_intensity!
Dim sngMaxMin_diff!

Dim i6i!
Dim i1s!
Dim i2s!
Dim i3s!
Dim i4s!

Public Const GrayScaleRGB As Long = 1 + 256 + 65536

Private HSV_Private As HSVTYPE
Private value1!

Dim sat1!
Dim diff1!
Dim Lng1&

Public RedsL(255) As Long
Public GreensL(255) As Long
Public BluesL(255) As Long
Public HuesL(1529) As Long
Public SatsL(255) As Long
Public ValsL(255) As Long

Public Function WuSample(BufSrc As AnimSurf2D, pel_src_x!, pel_src_y!)
Dim wu_edge_left!
Dim wu_edge_right!
Dim wu_edge_top!
Dim wu_edge_bottom!
Dim alpha_left!
Dim alpha_right!
Dim alpha_top!
Dim alpha_bottom!
Dim pel_center_x!
Dim pel_center_y!
Dim PelA&
Dim PelB&
Dim PelC&
Dim PelD&
Dim RedFour&
Dim GreenFour&
Dim BlueFour&
Dim alpha_topl!
Dim alpha_topr!
Dim alpha_botl!
Dim alpha_botr!
Dim LeftColumn&
Dim RightColumn&
Dim TopRow&
Dim BottomRow&
Dim RedSum&
Dim GreenSum&
Dim BlueSum&
Dim RGBSum&
Dim pel_division_x!
Dim pel_division_y!

 wu_edge_left = pel_src_x - 0.5
 wu_edge_right = wu_edge_left + 1
 wu_edge_bottom = pel_src_y - 0.5
 wu_edge_top = wu_edge_bottom + 1
 
 LeftColumn = Int(pel_src_x)
 BottomRow = Int(pel_src_y)
 
 pel_division_x = LeftColumn + 0.5
 pel_division_y = BottomRow + 0.5
 
 If pel_division_x <> wu_edge_right Then
 
  RightColumn = LeftColumn + 1
  alpha_left = pel_division_x - wu_edge_left
  alpha_right = 1 - alpha_left
  
  If pel_division_y <> wu_edge_top Then
  
   TopRow = BottomRow + 1
   alpha_bottom = pel_division_y - wu_edge_bottom
   alpha_top = 1 - alpha_bottom
   
   alpha_botl = alpha_bottom * alpha_left
   alpha_botr = alpha_bottom * alpha_right
   alpha_topl = alpha_left * alpha_top
   alpha_topr = alpha_right * alpha_top
   
   PelA = BufSrc.LDib(LeftColumn, TopRow)
   PelB = BufSrc.LDib(RightColumn, TopRow)
   PelC = BufSrc.LDib(LeftColumn, BottomRow)
   PelD = BufSrc.LDib(RightColumn, BottomRow)
 
   BlueSum = (PelA And &HFF&) * alpha_topl + _
             (PelB And &HFF&) * alpha_topr + _
             (PelC And &HFF&) * alpha_botl + _
             (PelD And &HFF&) * alpha_botr
 
   GreenSum = alpha_topl * (PelA And &HFF00&) / 256& _
            + alpha_topr * (PelB And &HFF00&) / 256& _
            + alpha_botl * (PelC And &HFF00&) / 256& _
            + alpha_botr * (PelD And &HFF00&) / 256&
 
   RedSum = alpha_topl * (PelA And &HFF0000) / 65536 _
          + alpha_topr * (PelB And &HFF0000) / 65536 _
          + alpha_botl * (PelC And &HFF0000) / 65536 _
          + alpha_botr * (PelD And &HFF0000) / 65536
 
  Else 'pel_division_y = wu_edge_top
  
   PelC = BufSrc.LDib(LeftColumn, BottomRow)
   PelD = BufSrc.LDib(RightColumn, BottomRow)
   BlueSum = (PelC And &HFF&) * alpha_left + _
             (PelD And &HFF&) * alpha_right
   GreenSum = alpha_left * (PelC And &HFF00&) / 256& _
            + alpha_right * (PelD And &HFF00&) / 256&
   RedSum = alpha_left * (PelC And &HFF0000) / 65536 _
          + alpha_right * (PelD And &HFF0000) / 65536
  
  End If 'pel_division_y <> wu_edge_top
  WuSample = RedSum * 65536 Or GreenSum * 256& Or BlueSum
 
 Else 'pel_division_x = wu_edge_right
 
  If pel_division_y <> wu_edge_top Then
   TopRow = BottomRow + 1
   alpha_bottom = pel_division_y - wu_edge_bottom
   alpha_top = 1 - alpha_bottom
   PelA = BufSrc.LDib(LeftColumn, TopRow)
   PelC = BufSrc.LDib(LeftColumn, BottomRow)
   BlueSum = (PelA And &HFF&) * alpha_top + _
             (PelC And &HFF&) * alpha_bottom
   GreenSum = alpha_top * (PelA And &HFF00&) / 256& _
            + alpha_bottom * (PelC And &HFF00&) / 256&
   RedSum = alpha_top * (PelA And &HFF0000) / 65536 _
          + alpha_bottom * (PelC And &HFF0000) / 65536
   WuSample = RedSum * 65536 Or GreenSum * 256& Or BlueSum
  Else 'pel_division_y = wu_edge_top
   WuSample = BufSrc.LDib(LeftColumn, BottomRow)
  End If 'pel_division_y <> wu_edge_top
  
 End If 'pel_division_x <> wu_edge_right
 
End Function
Public Function WuSampleAPI(BufSrc As AnimSurf2D, PrecisionPoint1 As PrecisionPointAPI)
Dim wu_edge_left!
Dim wu_edge_right!
Dim wu_edge_top!
Dim wu_edge_bottom!
Dim alpha_left!
Dim alpha_right!
Dim alpha_top!
Dim alpha_bottom!
Dim pel_center_x!
Dim pel_center_y!
Dim PelA&
Dim PelB&
Dim PelC&
Dim PelD&
Dim RedFour&
Dim GreenFour&
Dim BlueFour&
Dim alpha_topl!
Dim alpha_topr!
Dim alpha_botl!
Dim alpha_botr!
Dim LeftColumn&
Dim RightColumn&
Dim TopRow&
Dim BottomRow&
Dim RedSum&
Dim GreenSum&
Dim BlueSum&
Dim RGBSum&
Dim pel_division_x!
Dim pel_division_y!

 wu_edge_left = PrecisionPoint1.sX - 0.5
 wu_edge_right = wu_edge_left + 1
 wu_edge_bottom = PrecisionPoint1.sY - 0.5
 wu_edge_top = wu_edge_bottom + 1
 
 LeftColumn = Int(PrecisionPoint1.sX)
 BottomRow = Int(PrecisionPoint1.sY)
 
 pel_division_x = LeftColumn + 0.5
 pel_division_y = BottomRow + 0.5
 
 If pel_division_x <> wu_edge_right Then
 
  RightColumn = LeftColumn + 1
  alpha_left = pel_division_x - wu_edge_left
  alpha_right = 1 - alpha_left
  
  If pel_division_y <> wu_edge_top Then
  
   TopRow = BottomRow + 1
   alpha_bottom = pel_division_y - wu_edge_bottom
   alpha_top = 1 - alpha_bottom
   
   alpha_botl = alpha_bottom * alpha_left
   alpha_botr = alpha_bottom * alpha_right
   alpha_topl = alpha_left * alpha_top
   alpha_topr = alpha_right * alpha_top
   
   PelA = BufSrc.LDib(LeftColumn, TopRow)
   PelB = BufSrc.LDib(RightColumn, TopRow)
   PelC = BufSrc.LDib(LeftColumn, BottomRow)
   PelD = BufSrc.LDib(RightColumn, BottomRow)
 
   BlueSum = (PelA And &HFF&) * alpha_topl + _
             (PelB And &HFF&) * alpha_topr + _
             (PelC And &HFF&) * alpha_botl + _
             (PelD And &HFF&) * alpha_botr
 
   GreenSum = alpha_topl * (PelA And &HFF00&) / 256& _
            + alpha_topr * (PelB And &HFF00&) / 256& _
            + alpha_botl * (PelC And &HFF00&) / 256& _
            + alpha_botr * (PelD And &HFF00&) / 256&
 
   RedSum = alpha_topl * (PelA And &HFF0000) / 65536 _
          + alpha_topr * (PelB And &HFF0000) / 65536 _
          + alpha_botl * (PelC And &HFF0000) / 65536 _
          + alpha_botr * (PelD And &HFF0000) / 65536
 
  Else 'pel_division_y = wu_edge_top
  
   PelC = BufSrc.LDib(LeftColumn, BottomRow)
   PelD = BufSrc.LDib(RightColumn, BottomRow)
   BlueSum = (PelC And &HFF&) * alpha_left + _
             (PelD And &HFF&) * alpha_right
   GreenSum = alpha_left * (PelC And &HFF00&) / 256& _
            + alpha_right * (PelD And &HFF00&) / 256&
   RedSum = alpha_left * (PelC And &HFF0000) / 65536 _
          + alpha_right * (PelD And &HFF0000) / 65536
  
  End If 'pel_division_y <> wu_edge_top
  WuSampleAPI = RedSum * 65536 Or GreenSum * 256& Or BlueSum
 
 Else 'pel_division_x = wu_edge_right
 
  If pel_division_y <> wu_edge_top Then
   TopRow = BottomRow + 1
   alpha_bottom = pel_division_y - wu_edge_bottom
   alpha_top = 1 - alpha_bottom
   PelA = BufSrc.LDib(LeftColumn, TopRow)
   PelC = BufSrc.LDib(LeftColumn, BottomRow)
   BlueSum = (PelA And &HFF&) * alpha_top + _
             (PelC And &HFF&) * alpha_bottom
   GreenSum = alpha_top * (PelA And &HFF00&) / 256& _
            + alpha_bottom * (PelC And &HFF00&) / 256&
   RedSum = alpha_top * (PelA And &HFF0000) / 65536 _
          + alpha_bottom * (PelC And &HFF0000) / 65536
   WuSampleAPI = RedSum * 65536 Or GreenSum * 256& Or BlueSum
  Else 'pel_division_y = wu_edge_top
   WuSampleAPI = BufSrc.LDib(LeftColumn, BottomRow)
  End If 'pel_division_y <> wu_edge_top
  
 End If 'pel_division_x <> wu_edge_right
 
End Function
Public Sub ColorShift()

 If sR < sB Then
  If sR < sG Then
   If sG < sB Then
    sngMaxMin_diff = sB - sR
    sG = sG - sngMaxMin_diff * rgb_shift_intensity
    If sG < sR Then
     subt = sR - sG
     sG = sR
     sR = sR + subt
    End If
   Else
    sngMaxMin_diff = sG - sR
    sB = sB + sngMaxMin_diff * rgb_shift_intensity
    If sB > sG Then
     subt = sB - sG
     sB = sG
     sG = sG - subt
    End If
   End If
  Else
   sngMaxMin_diff = sB - sG
   sR = sR + sngMaxMin_diff * rgb_shift_intensity
   If sR > sB Then
    subt = sR - sB
    sR = sB
    sB = sB - subt
   End If
  End If
 ElseIf sR > sG Then
  If sB < sG Then
   sngMaxMin_diff = sR - sB
   sG = sG + sngMaxMin_diff * rgb_shift_intensity
   If sG > sR Then
    subt = sG - sR
    sG = sR
    sR = sR - subt
   End If
  Else
   sngMaxMin_diff = sR - sG
   sB = sB - sngMaxMin_diff * rgb_shift_intensity
   If sB < sG Then
    subt = sG - sB
    sB = sG
    sG = sG + subt
   End If
  End If
 Else
  sngMaxMin_diff = sG - sB
  sR = sR - sngMaxMin_diff * rgb_shift_intensity
  If sR < sB Then
   subt = sB - sR
   sR = sB
   sB = sB + subt
  End If
 End If
 
End Sub
Public Sub ColorShiftReverse()

 If sR < sB Then
  If sR < sG Then
   If sG < sB Then
    sngMaxMin_diff = sB - sR
    sG = sG + sngMaxMin_diff * rgb_shift_intensity
    If sG > sB Then
     subt = sG - sB
     sG = sB
     sB = sB - subt
    End If
   Else
    sngMaxMin_diff = sG - sR
    sB = sB - sngMaxMin_diff * rgb_shift_intensity
    If sB < sR Then
     subt = sR - sB
     sB = sR
     sR = sR + subt
    End If
   End If
  Else
   sngMaxMin_diff = sB - sG
   sR = sR - sngMaxMin_diff * rgb_shift_intensity
   If sR < sG Then
    subt = sG - sR
    sR = sG
    sG = sG + subt
   End If
  End If
 ElseIf sR > sG Then
  If sB < sG Then
   sngMaxMin_diff = sR - sB
   sG = sG - sngMaxMin_diff * rgb_shift_intensity
   If sG < sB Then
    subt = sB - sG
    sG = sB
    sB = sB + subt
   End If
  Else
   sngMaxMin_diff = sR - sG
   sB = sB + sngMaxMin_diff * rgb_shift_intensity
   If sB > sR Then
    subt = sB - sR
    sB = sR
    sR = sR - subt
   End If
  End If
 Else
  sngMaxMin_diff = sG - sB
  sR = sR + sngMaxMin_diff * rgb_shift_intensity
  If sR > sG Then
   subt = sR - sG
   sR = sG
   sG = sG - subt
  End If
 End If
 
End Sub
Public Sub NormaliseSHF(CurrentValue!, maximum!)

 SHF_intensity = Abs(CurrentValue) / (Abs(maximum) / 6)
 
 If maximum! < 0 Or CurrentValue < 0 Then
  SHF_intensity = 6 - SHF_intensity
 End If

 If SHF_intensity >= 1 Then
  If SHF_intensity < 2 Then
  i1s = SHF_intensity - 1
  ElseIf SHF_intensity < 3 Then
  i2s = SHF_intensity - 2
  ElseIf SHF_intensity < 4 Then
  i3s = SHF_intensity - 3
  ElseIf SHF_intensity < 5 Then
  i4s = SHF_intensity - 4
  ElseIf SHF_intensity < 6 Then
  i6i = 6 - SHF_intensity
  End If
 End If

End Sub
Public Sub HueShift()

 If sR < sB Then
  If sR < sG Then
   minim = sR
   If sG < sB Then
    maxim = sB
   Else
    maxim = sG
   End If
  Else
   maxim = sB
   minim = sG
  End If
 ElseIf sR > sG Then
  maxim = sR
  If sB < sG Then
   minim = sB
  Else
   minim = sG
  End If
 Else
  maxim = sG
  minim = sB
 End If
 sngMaxMin_diff = maxim - minim
 
 If sR = maxim Then
  If sB = minim Then 'g +
   If rgb_shift_intensity < 1 Then
    sG = sG + sngMaxMin_diff * rgb_shift_intensity
    If sG > maxim Then
     subt = sG - maxim
     sG = maxim
     sR = maxim - subt
    End If
   ElseIf rgb_shift_intensity < 2 Then
    sR = maxim + minim - sG - i1s * sngMaxMin_diff
    sG = maxim
    If sR < minim Then
     subt = minim - sR
     sR = minim
     sB = minim + subt
    Else
     sB = minim
    End If
   ElseIf rgb_shift_intensity < 3 Then
    sB = sG + i2s * sngMaxMin_diff
    sR = minim
    If sB > maxim Then
     subt = sB - maxim
     sB = maxim
     sG = maxim - subt
    Else
     sG = maxim
    End If
   ElseIf rgb_shift_intensity < 4 Then
    sG = maxim + minim - sG - i3s * sngMaxMin_diff
    sB = maxim
    If sG < minim Then
     subt = minim - sG
     sG = minim
     sR = minim + subt
    Else
     sR = minim
    End If
   ElseIf rgb_shift_intensity < 5 Then
    sR = sG + i4s * sngMaxMin_diff
    sG = minim
    If sR > maxim Then
     subt = sR - maxim
     sR = maxim
     sB = maxim - subt
    Else
     sB = maxim
    End If
   ElseIf rgb_shift_intensity < 6 Then
    sG = sG - i6i * sngMaxMin_diff
    If sG < minim Then
     subt = minim - sG
     sG = minim
     sB = minim + subt
    Else
     sB = minim
    End If
   End If
  Else 'r max, g min, blue -
   If rgb_shift_intensity < 1 Then
    sB = sB - sngMaxMin_diff * rgb_shift_intensity
    If sB < minim Then
     subt = minim - sB
     sB = minim
     sG = sG + subt
    End If
   ElseIf rgb_shift_intensity < 2 Then
    sG = maxim + minim - sB + sngMaxMin_diff * i1s
    sB = minim
    If sG > maxim Then
     subt = sG - maxim
     sG = maxim
     sR = maxim - subt
    Else
     sR = maxim
    End If
   ElseIf rgb_shift_intensity < 3 Then
    sR = sB - sngMaxMin_diff * i2s
    sG = maxim
    If sR < minim Then
     subt = minim - sR
     sR = minim
     sB = minim + subt
    Else
     sB = minim
    End If
   ElseIf rgb_shift_intensity < 4 Then
    sB = maxim + minim - sB + sngMaxMin_diff * i3s
    sR = minim
    If sB > maxim Then
     subt = sB - maxim
     sB = maxim
     sG = maxim - subt
    Else
     sG = maxim
    End If
   ElseIf rgb_shift_intensity < 5 Then
    sG = sB - sngMaxMin_diff * i4s
    sB = maxim
    If sG < minim Then
     subt = minim - sG
     sG = minim
     sR = minim + subt
    Else
     sR = minim
    End If
   ElseIf rgb_shift_intensity < 6 Then
    sB = sB + sngMaxMin_diff * i6i
    If sB > maxim Then
     subt = sB - maxim
     sB = maxim
     sR = maxim - subt
    Else
     sR = maxim
    End If
   End If
  End If
 ElseIf sR = minim Then
  If sG = maxim Then 'blue +
   If rgb_shift_intensity < 1 Then
    sB = sB + sngMaxMin_diff * rgb_shift_intensity
    If sB > maxim Then
     subt = sB - maxim
     sB = maxim
     sG = maxim - subt
    End If
   ElseIf rgb_shift_intensity < 2 Then
    sG = maxim + minim - sB - sngMaxMin_diff * i1s
    sB = maxim
    If sG < minim Then
     subt = minim - sG
     sG = minim
     sR = minim + subt
    Else
     sR = minim
    End If
   ElseIf rgb_shift_intensity < 3 Then
    sR = sB + sngMaxMin_diff * i2s
    sG = minim
    If sR > maxim Then
     subt = sR - maxim
     sR = maxim
     sB = maxim - subt
    Else
     sB = maxim
    End If
   ElseIf rgb_shift_intensity < 4 Then
    sB = maxim + minim - sB - sngMaxMin_diff * i3s
    sR = maxim
    If sB < minim Then
     subt = minim - sB
     sB = minim
     sG = minim + subt
    Else
     sG = minim
    End If
   ElseIf rgb_shift_intensity < 5 Then
    sG = sB + sngMaxMin_diff * i4s
    sB = minim
    If sG > maxim Then
     subt = sG - maxim
     sG = maxim
     sR = maxim - subt
    Else
     sR = maxim
    End If
   ElseIf rgb_shift_intensity < 6 Then
    sB = sB - sngMaxMin_diff * i6i
    If sB < minim Then
     subt = minim - sB
     sB = minim
     sR = minim + subt
    Else
     sR = minim
    End If
   End If
  Else 'blue max green -
   If rgb_shift_intensity < 1 Then
    sG = sG - sngMaxMin_diff * rgb_shift_intensity
    If sG < minim Then
     subt = minim - sG
     sG = minim
     sR = minim + subt
    End If
   ElseIf rgb_shift_intensity < 2 Then
    sR = maxim + minim - sG + sngMaxMin_diff * i1s
    sG = minim
    If sR > maxim Then
     subt = sR - maxim
     sR = maxim
     sB = maxim - subt
    Else
     sB = maxim
    End If
   ElseIf rgb_shift_intensity < 3 Then
    sB = sG - sngMaxMin_diff * i2s
    sR = maxim
    If sB < minim Then
     subt = minim - sB
     sB = minim
     sG = minim + subt
    Else
     sG = minim
    End If
   ElseIf rgb_shift_intensity < 4 Then
    sG = maxim + minim - sG + sngMaxMin_diff * i3s
    sB = minim
    If sG > maxim Then
     subt = sG - maxim
     sG = maxim
     sR = maxim - subt
    Else
     sR = maxim
    End If
   ElseIf rgb_shift_intensity < 5 Then
    sR = sG - sngMaxMin_diff * i4s
    sG = maxim
    If sR < minim Then
     subt = minim - sR
     sR = minim
     sB = minim + subt
    Else
     sB = minim
    End If
   ElseIf rgb_shift_intensity < 6 Then
    sG = sG + sngMaxMin_diff * i6i
    If sG > maxim Then
     subt = sG - maxim
     sG = maxim
     sB = maxim - subt
    Else
     sB = maxim
    End If
   End If
  End If
 ElseIf sB = maxim Then 'green min, red +
  If rgb_shift_intensity < 1 Then
   sR = sR + sngMaxMin_diff * rgb_shift_intensity
   If sR > maxim Then
    subt = sR - maxim
    sR = maxim
    sB = maxim - subt
   End If
  ElseIf rgb_shift_intensity < 2 Then
   sB = maxim + minim - sR - sngMaxMin_diff * i1s
   sR = maxim
   If sB < minim Then
    subt = minim - sB
    sB = minim
    sG = minim + subt
   Else
    sG = minim
   End If
  ElseIf rgb_shift_intensity < 3 Then
   sG = sR + sngMaxMin_diff * i2s
   sB = minim
   If sG > maxim Then
    subt = sG - maxim
    sG = maxim
    sR = maxim - subt
   Else
    sR = maxim
   End If
  ElseIf rgb_shift_intensity < 4 Then
   sR = maxim + minim - sR - sngMaxMin_diff * i3s
   sG = maxim
   If sR < minim Then
    subt = minim - sR
    sR = minim
    sB = minim + subt
   Else
    sB = minim
   End If
  ElseIf rgb_shift_intensity < 5 Then
   sB = sR + sngMaxMin_diff * i4s
   sR = minim
   If sB > maxim Then
    subt = sB - maxim
    sB = maxim
    sG = maxim - subt
   Else
    sG = maxim
   End If
  ElseIf rgb_shift_intensity < 6 Then
   sR = sR - sngMaxMin_diff * i6i
   If sR < minim Then
    subt = minim - sR
    sR = minim
    sG = minim + subt
   Else
    sG = minim
   End If
  End If
 Else 'blue min, green max, red -
  If rgb_shift_intensity < 1 Then
   sR = sR - sngMaxMin_diff * rgb_shift_intensity
   If sR < minim Then
    subt = minim - sR
    sR = minim
    sB = minim + subt
   End If
  ElseIf rgb_shift_intensity < 2 Then
   sB = minim - sR + maxim + sngMaxMin_diff * i1s
   sR = minim
   If sB > maxim Then
    subt = sB - maxim
    sB = maxim
    sG = maxim - subt
   Else
    sG = maxim
   End If
  ElseIf rgb_shift_intensity < 3 Then
   sG = sR - sngMaxMin_diff * i2s
   sB = maxim
   If sG < minim Then
    subt = minim - sG
    sG = minim
    sR = minim + subt
   Else
    sR = minim
   End If
  ElseIf rgb_shift_intensity < 4 Then
   sR = maxim + minim - sR + sngMaxMin_diff * i3s
   sG = minim
   If sR > maxim Then
    subt = sR - maxim
    sR = maxim
    sB = maxim - subt
   Else
    sB = maxim
   End If
  ElseIf rgb_shift_intensity < 5 Then
   sB = sR - sngMaxMin_diff * i4s
   sR = maxim
   If sB < minim Then
    subt = minim - sB
    sB = minim
    sG = minim + subt
   Else
    sG = minim
   End If
  ElseIf rgb_shift_intensity < 6 Then
   sR = sR + sngMaxMin_diff * i6i
   If sR > maxim Then
    subt = sR - maxim
    sR = maxim
    sG = maxim - subt
   Else
    sG = maxim
   End If
  End If
 End If
 
End Sub
Public Sub ShiftSpriteColor(Surf As AnimSurf2D, amount!)
Dim DrawX&
Dim DrawY&
Dim H_Segs&
Dim SegsL&
Dim SLX&
Dim SLE&
Dim TSL&
Dim SrcX&, SrcY&

 If Surf.Dims.bSRGB Then
 
  rgb_shift_intensity = amount - 6 * Int(amount / 6)
 
  If rgb_shift_intensity <> 0 Then
  
   If rgb_shift_intensity <= 1! Then
   
    For DrawY = 0 To Surf.TopLeft Step Surf.Dims.Wide
     H_Segs = Surf.HorizRLenSegs(SrcY)
     TSL = 0
     For SegsL = 0 To H_Segs Step 2
      SLX = TSL + Surf.RLSeg(SegsL, SrcY)
      SLE = SLX + Surf.RLSeg(SegsL + 1, SrcY)
      TSL = SLE + 1
      For SrcX = SLX To SLE
       sR = Surf.sRGB(SrcX, SrcY).sRed
       sG = Surf.sRGB(SrcX, SrcY).sGrn
       sB = Surf.sRGB(SrcX, SrcY).sBlu
       ColorShift
       Surf.Dib(SrcX, SrcY).Red = sR
       Surf.Dib(SrcX, SrcY).Green = sG
       Surf.Dib(SrcX, SrcY).Blue = sB
       Surf.sRGB(SrcX, SrcY).sRed = sR
       Surf.sRGB(SrcX, SrcY).sGrn = sG
       Surf.sRGB(SrcX, SrcY).sBlu = sB
      Next SrcX
     Next SegsL
     SrcY = SrcY + 1
    Next DrawY
    
   ElseIf rgb_shift_intensity >= 5 Then
 
    rgb_shift_intensity = 6! - rgb_shift_intensity
  
    For DrawY = 0 To Surf.TopLeft Step Surf.Dims.Wide
     H_Segs = Surf.HorizRLenSegs(SrcY)
     TSL = 0
     For SegsL = 0 To H_Segs Step 2
      SLX = TSL + Surf.RLSeg(SegsL, SrcY)
      SLE = SLX + Surf.RLSeg(SegsL + 1, SrcY)
      TSL = SLE + 1
      For SrcX = SLX To SLE
       sR = Surf.sRGB(SrcX, SrcY).sRed
       sG = Surf.sRGB(SrcX, SrcY).sGrn
       sB = Surf.sRGB(SrcX, SrcY).sBlu
       ColorShiftReverse
       Surf.Dib(SrcX, SrcY).Red = sR
       Surf.Dib(SrcX, SrcY).Green = sG
       Surf.Dib(SrcX, SrcY).Blue = sB
       Surf.sRGB(SrcX, SrcY).sRed = sR
       Surf.sRGB(SrcX, SrcY).sGrn = sG
       Surf.sRGB(SrcX, SrcY).sBlu = sB
      Next SrcX
     Next SegsL
     SrcY = SrcY + 1
    Next DrawY
   
   Else '1 < rgb_shift_intensity > 5
 
    NormaliseSHF rgb_shift_intensity, 6!
    For DrawY = 0 To Surf.TopLeft Step Surf.Dims.Wide
     H_Segs = Surf.HorizRLenSegs(SrcY)
     TSL = 0
     For SegsL = 0 To H_Segs Step 2
      SLX = TSL + Surf.RLSeg(SegsL, SrcY)
      SLE = SLX + Surf.RLSeg(SegsL + 1, SrcY)
      TSL = SLE + 1
      For SrcX = SLX To SLE
       sR = Surf.sRGB(SrcX, SrcY).sRed
       sG = Surf.sRGB(SrcX, SrcY).sGrn
       sB = Surf.sRGB(SrcX, SrcY).sBlu
       HueShift
       Surf.Dib(SrcX, SrcY).Red = sR
       Surf.Dib(SrcX, SrcY).Green = sG
       Surf.Dib(SrcX, SrcY).Blue = sB
       Surf.sRGB(SrcX, SrcY).sRed = sR
       Surf.sRGB(SrcX, SrcY).sGrn = sG
       Surf.sRGB(SrcX, SrcY).sBlu = sB
      Next SrcX
     Next SegsL
     SrcY = SrcY + 1
    Next DrawY
  
   End If
  
  End If
 
 End If

End Sub


Public Sub ZeroAlphaChannel(Surf As AnimSurf2D)

 For DrawY = 0 To Surf.Dims.HighM1
 For DrawX = 0 To Surf.Dims.WideM1
  Surf.Dib(DrawX, DrawY).Alpha = 0
  'Lng1 = Surf.Dib(DrawX, DrawY).Blue
  'Lng1 = Surf.Dib(DrawX, DrawY).Green
  'Lng1 = Surf.Dib(DrawX, DrawY).Red
 Next
 Next
 
End Sub
Private Sub RotateSourceOntoDest(SurfDest As AnimSurf2D, SurfSrc As AnimSurf2D, angBotLeft!, angRelHeading!, sLen!, MaskColor&, CollisionObjectType As Long)
 RotateSourceOntoDestUnWrap SurfAuto(SurfDest.Index), SurfAuto(SurfSrc.Index), angBotLeft, angRelHeading, sLen, MaskColor, CollisionObjectType
End Sub
Private Sub RotateSourceOntoDestUnWrap(SurfDest As AnimSurf2D, SurfSrc As AnimSurf2D, angBotLeft!, angRelHeading!, sLen!, ByVal MaskColor&, CollisionObjectType As Long)
Dim proj_x!, proj_y!, ix!, iy!
Dim x_left!, y_left!
Dim LngX&, LngY&
Dim SrcX&, SrcY&
Dim MR_X&
Dim MR_Blt&
Dim cBlt&
Dim RLSeg_Ptr&
Dim sAlpha!

    If MaskColor = -1 Then MaskColor = 0
    
    ReDim SurfDest.HorizRLenSegs(SurfDest.Dims.HighM1)
    ReDim SurfDest.RLSeg(SurfDest.Dims.Wide, SurfDest.Dims.HighM1)
    
    iy = -Cos(angRelHeading) 'source increment for x
    ix = Sin(angRelHeading) 'source increment for y
    x_left = SurfSrc.Dims.WideM1 / 2 + sLen * Cos(angBotLeft)
    y_left = SurfSrc.Dims.HighM1 / 2 + sLen * Sin(angBotLeft)
    For LngY = 0 To SurfSrc.Dims.HighM1
     proj_x = x_left
     proj_y = y_left
     RLSeg_Ptr = -1 '3 run-length encode variables
     MR_Blt = 0
     cBlt = -1
     For LngX = 0 To SurfSrc.Dims.WideM1
      SrcX = Int(proj_x + 0.5!) 'Round
      SrcY = Int(proj_y + 0.5!)
      If SrcX > -1 And SrcX < SurfSrc.Dims.Wide _
       And SrcY > -1 And SrcY < SurfSrc.Dims.High Then
       SurfDest.LDib(LngX, LngY) = SurfSrc.LDib(SrcX, SrcY)
       sAlpha = SurfSrc.sRGB(SrcX, SrcY).aLph
       SurfDest.sRGB(LngX, LngY).aLph = sAlpha
       SurfDest.sRGB(LngX, LngY).sRed = SurfSrc.sRGB(SrcX, SrcY).sRed
       SurfDest.sRGB(LngX, LngY).sGrn = SurfSrc.sRGB(SrcX, SrcY).sGrn
       SurfDest.sRGB(LngX, LngY).sBlu = SurfSrc.sRGB(SrcX, SrcY).sBlu
       If sAlpha < 0.002 Then
        If MR_Blt = 1 Then 'MR_Blt = run-length encode var
         RLSeg_Ptr = RLSeg_Ptr + 1 'run-length encode var
         SurfDest.RLSeg(RLSeg_Ptr, LngY) = cBlt
         cBlt = -1 'run-length encode var
         MR_Blt = 0
        End If
       Else 'sAlpha >= .002
        If MR_Blt = 0 Then
         RLSeg_Ptr = RLSeg_Ptr + 1
         SurfDest.RLSeg(RLSeg_Ptr, LngY) = cBlt + 1
         cBlt = -1
         MR_Blt = 1
        End If
       End If
      Else
       SurfDest.LDib(LngX, LngY) = MaskColor
       If MR_Blt = 1 Then
        RLSeg_Ptr = RLSeg_Ptr + 1
        SurfDest.RLSeg(RLSeg_Ptr, LngY) = cBlt
        cBlt = -1
        MR_Blt = 0
       End If
      End If
      proj_x = proj_x + ix
      proj_y = proj_y + iy
      cBlt = cBlt + 1
     Next LngX
     SurfDest.HorizRLenSegs(LngY) = RLSeg_Ptr
     If MR_Blt = 1 Then
      RLSeg_Ptr = RLSeg_Ptr + 1
      SurfDest.RLSeg(RLSeg_Ptr, LngY) = cBlt
     End If
     x_left = x_left - iy
     y_left = y_left + ix
    Next LngY
    
End Sub
Public Sub MakeSurf(Surf As AnimSurf2D, ByVal LWidth&, ByVal LHeight&, Optional ByVal LBX% = 0, Optional ByVal LBY = 0, Optional ByVal DoCreatePrecisionRGBA As Boolean = False, Optional ByVal DoCreateMapHSV As Boolean = False, Optional CreateErase As Boolean = False)
Dim MemBits&

 SetDims Surf, LWidth, LHeight

 If Surf.TotalPixels > 0 Then
 
    ClearDC Surf
  
    With Surf.BMinfo.bmiHeader
        .biSize = Len(Surf.BMinfo.bmiHeader)
        .biWidth = Surf.Dims.Wide
        .biHeight = Surf.Dims.High
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = 4 * Surf.TotalPixels
    End With
    
    Surf.mem_hDC = CreateCompatibleDC(0)
    If (Surf.mem_hDC <> 0) Then
    Surf.mem_hDIb = CreateDIBSection(Surf.mem_hDC, Surf.BMinfo, _
            DIB_RGB_COLORS, _
            MemBits, _
            0, 0)
        If Surf.mem_hDIb <> 0 Then
            Surf.mem_hBmpPrev = SelectObject(Surf.mem_hDC, Surf.mem_hDIb)
        Else
            DeleteObject Surf.mem_hDC
            Surf.mem_hDC = 0
        End If
    End If
    
    Surf.UB1D = Surf.TotalPixels - 1
 
    If DoCreatePrecisionRGBA Then
        ReDim Surf.sRGB(Surf.Dims.WideM1, Surf.Dims.HighM1)
        Surf.Dims.bSRGB = True
    End If
    
    If DoCreateMapHSV Then
        ReDim Surf.HSV1D(Surf.UB1D)
        Surf.Dims.bHSV = True
    End If
  
    SetSafeArrays Surf, MemBits, , DoCreatePrecisionRGBA, LBX, LBY
    
    If CreateErase Then
        ReDim Surf.EraseDib(Surf.UB1D)
        ReDim Surf.EraseStack(Surf.UB1D)
    End If
 
 End If
 
End Sub
Public Sub SetSafeArrays(Surf As AnimSurf2D, MemBits As Long, Optional BytesPixel As Byte = 4, Optional DoCreatePRGBA As Boolean = True, Optional ByVal LBX& = 0, Optional ByVal LBY& = 0)

    ClearSafeArrays Surf
    
    With Surf.SafeAry2D
    .cbElements = BytesPixel
    .cDims = 2
    .Bounds(0).lLbound = LBY
    .Bounds(1).lLbound = LBX
    .Bounds(0).cElements = Surf.Dims.High
    .Bounds(1).cElements = Surf.Dims.Wide
    .pvData = MemBits
    End With
    CopyMemory ByVal VarPtrArray(Surf.Dib), VarPtr(Surf.SafeAry2D), 4
 
    With Surf.SafeAry2D_L
    .cbElements = BytesPixel
    .cDims = 2
    .Bounds(0).lLbound = LBY
    .Bounds(1).lLbound = LBX
    .Bounds(0).cElements = Surf.Dims.High
    .Bounds(1).cElements = Surf.Dims.Wide
    .pvData = MemBits
    End With
    CopyMemory ByVal VarPtrArray(Surf.LDib), VarPtr(Surf.SafeAry2D_L), 4
 
    With Surf.SafeAry1D
    .cbElements = BytesPixel
    .cDims = 1
    .lLbound = 0
    .cElements = Surf.TotalPixels
    .pvData = MemBits
    End With
    CopyMemory ByVal VarPtrArray(Surf.Dib1D), VarPtr(Surf.SafeAry1D), 4
  
    With Surf.SafeAry1D_L
    .cbElements = BytesPixel
    .cDims = 1
    .lLbound = 0
    .cElements = Surf.TotalPixels
    .pvData = MemBits
    End With
    CopyMemory ByVal VarPtrArray(Surf.LDib1D), VarPtr(Surf.SafeAry1D_L), 4
    
    If Surf.Dims.bSRGB Then
    With Surf.SafeAry1D_sRGB
    .cbElements = 16
    .cDims = 1
    .lLbound = 0
    .cElements = Surf.TotalPixels
    .pvData = VarPtr(Surf.sRGB(0, 0).sRed)
    End With
    CopyMemory ByVal VarPtrArray(Surf.sRGB1D), VarPtr(Surf.SafeAry1D_sRGB), 4
    End If
    
End Sub

Public Sub CreateSurfaceFromFile(Surf As AnimSurf2D, strFileName$, Optional DoMask As Boolean = True, Optional ByVal MaskColor = -1, Optional ByVal sRed! = 1!, Optional ByVal sGreen! = 1!, Optional ByVal sBlue! = 1!, Optional sColorShift!, Optional DoColorize As Boolean, Optional ByVal StrFolder$ = "", Optional DoCreatePrecisionRGBA As Boolean = False, Optional DoMapHSV As Boolean = False)
Dim tBM As BITMAP, sPic As StdPicture
Dim CDC&, Loops&

On Local Error GoTo OOPS

    If StrFolder <> "" Then
     StrFileFolder = StrFolder
    End If
    
    ClearSurface Surf
    
    If Right$(StrFileFolder, 1) = "\" Or StrFileFolder = "" Then
        Set sPic = LoadPicture(StrFileFolder & strFileName)
    Else
        Set sPic = LoadPicture(StrFileFolder & "\" & strFileName)
    End If
    
    If sPic <> 0 Then
    
    CDC = CreateCompatibleDC(0)           ' Temporary device
    DeleteObject SelectObject(CDC, sPic)  ' Converted bitmap
    
    GetObjectAPI sPic, Len(tBM), tBM
    
    MakeSurf Surf, tBM.bmWidth, tBM.bmHeight, , , DoCreatePrecisionRGBA, DoMapHSV
    
    If Surf.TotalPixels > 0 Then
 
      BitBlt Surf.mem_hDC, 0, 0, tBM.bmWidth, tBM.bmHeight, _
             CDC, 0, 0, vbSrcCopy
                   
      ZeroAlphaChannel Surf
      
      CreateMaskStructure Surf, DoMask, MaskColor, sRed, sGreen, sBlue, , , , DoColorize
      
      ShiftSpriteColor Surf, sColorShift
      
      SetSurfMapHSV Surf
             
    End If
    
    End If 'Surf.TotalPixels > 0
 
OOPS:
 
    DeleteDC CDC
 
End Sub
Private Sub SetSurfMapHSV(Surf As AnimSurf2D)
 If Surf.Dims.bHSV Then
  For Lng1 = 0 To Surf.UB1D
   HSVTYPE_From_RGBQUAD_P Surf.HSV1D(Lng1), Surf.Dib1D(Lng1)
  Next
 End If
End Sub
Function CreatePicture(ByVal nWidth&, ByVal nHeight&, ByVal BitDepth&) As Picture
Dim Pic As PicBmp, IID_IDispatch As GUID
Dim BMI As BITMAPINFO
With BMI.bmiHeader
.biSize = Len(BMI.bmiHeader)
.biWidth = nWidth
.biHeight = nHeight
.biPlanes = 1
.biBitCount = BitDepth
End With
Pic.hBmp = CreateDIBSection(0, BMI, 0, 0, 0, 0)
IID_IDispatch.Data1 = &H20400: IID_IDispatch.Data4(0) = &HC0: IID_IDispatch.Data4(7) = &H46
Pic.Size = Len(Pic)
Pic.Type = vbPicTypeBitmap
OleCreatePictureIndirect Pic, IID_IDispatch, 1, CreatePicture
If CreatePicture = 0 Then Set CreatePicture = Nothing
End Function
Public Sub CreateFileFromSurface(strFileName$, Surf As AnimSurf2D, Optional CreateAlphaFileAlso As Boolean, Optional ByVal StrFolder$ = "")
Dim tBM As BITMAP, sPic As StdPicture
Dim CDC&, Loops&, Surf1 As AnimSurf2D
Dim LocalStrFile$

    If StrFolder <> "" Then
     StrFileFolder = StrFolder
    End If
    
    Set sPic = CreatePicture(Surf.Dims.Wide, Surf.Dims.High, 24)
    
    CDC = CreateCompatibleDC(0)           ' Temporary device
    DeleteObject SelectObject(CDC, sPic)  ' Converted bitmap
    
    BitBlt CDC, 0, 0, Surf.Dims.Wide, Surf.Dims.High, _
           Surf.mem_hDC, 0, 0, vbSrcCopy
    
    If Right$(strFileName, 4) <> ".bmp" Then
     strFileName = strFileName & ".bmp"
    End If
 
    If Right$(StrFileFolder, 1) = "\" Or StrFileFolder = "" Then
      LocalStrFile = StrFileFolder & strFileName
    Else
      LocalStrFile = StrFileFolder & "\" & strFileName
    End If
    
    SavePicture sPic, LocalStrFile
    
    If CreateAlphaFileAlso Then
     
     MakeSurf Surf1, Surf.Dims.Wide, Surf.Dims.High
     
     If Surf.Dims.bSRGB Then
      For DrawX = 0 To Surf.UB1D
       Surf1.LDib1D(DrawX) = Int(255 * Surf.sRGB1D(DrawX).aLph + 0.5) * GrayScaleRGB
      Next
     Else
      For DrawX = 0 To Surf.UB1D
       Surf1.LDib1D(DrawX) = Surf.Dib1D(DrawX).Alpha * GrayScaleRGB
      Next
     End If
     
     BitBlt CDC, 0, 0, Surf.Dims.Wide, Surf.Dims.High, _
            Surf1.mem_hDC, 0, 0, vbSrcCopy
     
     SavePicture sPic, Left$(LocalStrFile, Len(LocalStrFile) - 4) & "A" & ".bmp"
     
     ClearSurface Surf1
     
    End If
 
    DeleteDC CDC
    
End Sub
Public Sub CreateMaskStructure(Surf As AnimSurf2D, Optional DoMask As Boolean = True, Optional ByVal MaskColor = -1, Optional ByVal sRed! = 1!, Optional ByVal sGreen! = 1!, Optional ByVal sBlue! = 1!, Optional ByVal Alpha As Byte = 255, Optional BlitType& = -1, Optional CollisionObjectType& = -1, Optional DoColorize As Boolean)
Dim tsRed!
Dim tsGrn!
Dim tsBlu!
Dim tsMax!
Dim sAlpha!
Dim MR_X&
Dim MR_Blt&
Dim cBlt&
Dim RLSeg_Ptr&

    ReDim Surf.HorizRLenSegs(Surf.Dims.HighM1)
    ReDim Surf.RLSeg(Surf.Dims.Wide, Surf.Dims.HighM1)
      
    If Alpha > 0 Then
     tsMax = (sRed + sGreen + sBlue) * 255 / Alpha
    Else
     tsMax = 2550
    End If
    tsRed = sRed / tsMax
    tsGrn = sGreen / tsMax
    tsBlu = sBlue / tsMax
        
    If DoColorize Or BlitType = BLIT_TYPE_RGB Then
     For DrawY = 0 To Surf.Dims.HighM1
     For DrawX = 0 To Surf.Dims.WideM1
      BGRed = Surf.Dib(DrawX, DrawY).Red * sRed
      BGGrn = Surf.Dib(DrawX, DrawY).Green * sGreen
      BGBlu = Surf.Dib(DrawX, DrawY).Blue * sBlue
      If BGRed > 255 Then BGRed = 255
      If BGGrn > 255 Then BGGrn = 255
      If BGBlu > 255 Then BGBlu = 255
      Surf.Dib(DrawX, DrawY).Red = BGRed
      Surf.Dib(DrawX, DrawY).Green = BGGrn
      Surf.Dib(DrawX, DrawY).Blue = BGBlu
     Next DrawX
     Next DrawY
    End If
    
    If DoMask Then
      
     If MaskColor = -1 Then
      
      For DrawY = 0 To Surf.Dims.HighM1
       RLSeg_Ptr = -1
       MR_Blt = 0
       cBlt = -1
       For DrawX = 0 To Surf.Dims.WideM1
        BGRed = Surf.Dib(DrawX, DrawY).Red
        BGGrn = Surf.Dib(DrawX, DrawY).Green
        BGBlu = Surf.Dib(DrawX, DrawY).Blue
        sAlpha = BGRed * tsRed + BGGrn * tsGrn + BGBlu * tsBlu
        sAlpha = sAlpha / 255
        If Surf.Dims.bSRGB Then
         Surf.sRGB(DrawX, DrawY).sRed = BGRed
         Surf.sRGB(DrawX, DrawY).sGrn = BGGrn
         Surf.sRGB(DrawX, DrawY).sBlu = BGBlu
         Surf.sRGB(DrawX, DrawY).aLph = sAlpha
        End If
        If sAlpha < 0.002 Then
         If MR_Blt = 1 Then
         RLSeg_Ptr = RLSeg_Ptr + 1
         Surf.RLSeg(RLSeg_Ptr, DrawY) = cBlt
         cBlt = -1
         MR_Blt = 0
         End If
        Else 'sAlpha >= .002
         If MR_Blt = 0 Then
          RLSeg_Ptr = RLSeg_Ptr + 1
          Surf.RLSeg(RLSeg_Ptr, DrawY) = cBlt + 1
          cBlt = -1
          MR_Blt = 1
         End If
        End If
        cBlt = cBlt + 1
       Next DrawX
       Surf.HorizRLenSegs(DrawY) = RLSeg_Ptr
       If MR_Blt = 1 Then
        RLSeg_Ptr = RLSeg_Ptr + 1
        Surf.RLSeg(RLSeg_Ptr, DrawY) = cBlt
       End If
      Next DrawY
      
     Else 'MaskColor <> -1
      
      BGBlu = (MaskColor And vbRed) * 65536
      BGGrn = MaskColor And vbGreen
      BGRed = (MaskColor And vbBlue) / 65536
      MaskColor = BGRed Or BGGrn Or BGBlu
        
      For DrawY = 0 To Surf.Dims.HighM1
      RLSeg_Ptr = -1
      MR_Blt = 0
      cBlt = -1
      For DrawX = 0 To Surf.Dims.WideM1
       BGRed = Surf.Dib(DrawX, DrawY).Red
       BGGrn = Surf.Dib(DrawX, DrawY).Green
       BGBlu = Surf.Dib(DrawX, DrawY).Blue
       sAlpha = BGRed * tsRed + BGGrn * tsGrn + BGBlu * tsBlu
       sAlpha = sAlpha / 255
       'precision color array
       If Surf.Dims.bSRGB Then
        Surf.sRGB(DrawX, DrawY).sRed = BGRed
        Surf.sRGB(DrawX, DrawY).sGrn = BGGrn
        Surf.sRGB(DrawX, DrawY).sBlu = BGBlu
        Surf.sRGB(DrawX, DrawY).aLph = sAlpha
       End If
       If Surf.LDib(DrawX, DrawY) = MaskColor Then
        If MR_Blt = 1 Then
         RLSeg_Ptr = RLSeg_Ptr + 1
         Surf.RLSeg(RLSeg_Ptr, DrawY) = cBlt
         cBlt = -1
         MR_Blt = 0
        End If
       Else
        If MR_Blt = 0 Then
         RLSeg_Ptr = RLSeg_Ptr + 1
         Surf.RLSeg(RLSeg_Ptr, DrawY) = cBlt + 1
         cBlt = -1
         MR_Blt = 1
        End If
       End If
       cBlt = cBlt + 1
      Next DrawX
      Surf.HorizRLenSegs(DrawY) = RLSeg_Ptr
      If MR_Blt = 1 Then
       RLSeg_Ptr = RLSeg_Ptr + 1
       Surf.RLSeg(RLSeg_Ptr, DrawY) = cBlt
      End If
      Next DrawY
      
     End If 'MaskColor = -1
       
    Else 'DoMask = False
      
     tsMax = sRed
     If sGreen > tsMax Then tsMax = sGreen
     If sBlue > tsMax Then tsMax = sBlue
     tsRed = sRed / tsMax
     tsGrn = sGreen / tsMax
     tsBlu = sBlue / tsMax
     sAlpha = Alpha / 255
     For DrawY = 0 To Surf.Dims.HighM1
     For DrawX = 0 To Surf.Dims.WideM1
      BGRed = Surf.Dib(DrawX, DrawY).Red
      BGGrn = Surf.Dib(DrawX, DrawY).Green
      BGBlu = Surf.Dib(DrawX, DrawY).Blue
      BGRed = BGRed * tsRed
      BGGrn = BGGrn * tsGrn
      BGBlu = BGBlu * tsBlu
      If Surf.Dims.bSRGB Then
       Surf.sRGB(DrawX, DrawY).aLph = sAlpha
       Surf.sRGB(DrawX, DrawY).sRed = BGRed
       Surf.sRGB(DrawX, DrawY).sGrn = BGGrn
       Surf.sRGB(DrawX, DrawY).sBlu = BGBlu
      End If
      Surf.Dib(DrawX, DrawY).Red = BGRed
      Surf.Dib(DrawX, DrawY).Green = BGGrn
      Surf.Dib(DrawX, DrawY).Blue = BGBlu
     Next DrawX
     Surf.HorizRLenSegs(DrawY) = 0
     Surf.RLSeg(1, DrawY) = Surf.Dims.WideM1
     Next DrawY
       
    End If
    
End Sub
Public Sub CreateSurfaceFromBitmapStructure(Surf As AnimSurf2D, BM As BITMAP)

 ClearDC Surf
 SetDims Surf, BM.bmWidth, BM.bmHeight
 SetSafeArrays Surf, BM.bmBits, (BM.bmBitsPixel / 8)
 CreateMaskStructure Surf

End Sub

'BitBlt wrapper
Public Sub BlitToDC(ByVal lHDC As Long, Surf As AnimSurf2D, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
        )

    If (lDestWidth < 0) Then lDestWidth = Surf.BMinfo.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = Surf.BMinfo.bmiHeader.biHeight
    
    BitBlt lHDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, Surf.mem_hDC, lSrcLeft, lSrcTop, eRop
    
End Sub
Public Sub Tile(SurfDest As AnimSurf2D, SurfSrc As AnimSurf2D, ByVal DestX&, ByVal DestY&)
Dim XLeft&
Dim XRight&
Dim YTop&
Dim YBot&

 XLeft = Int(DestX) Mod SurfSrc.Dims.Wide
 If XLeft > 0 Then XLeft = XLeft - SurfSrc.Dims.Wide
 
 YBot = Int(DestY) Mod SurfSrc.Dims.High
 If YBot > 0 Then YBot = YBot - SurfSrc.Dims.High
 
 XRight = SurfDest.Dims.WideM1
 YTop = SurfDest.Dims.HighM1
 
 For DrawY = YBot To YTop Step SurfSrc.Dims.High
  For DrawX = XLeft To XRight Step SurfSrc.Dims.Wide
   BitBlt SurfDest.mem_hDC, DrawX, DrawY, _
               SurfSrc.Dims.Wide, SurfSrc.Dims.High, _
                  SurfSrc.mem_hDC, 0, 0, vbSrcCopy
  Next
 Next

End Sub
Public Sub SolidColorFill(Surf As AnimSurf2D, Red As Byte, Green As Byte, Blue As Byte)
Dim BGR&

 BGR = RGB(Blue, Green, Red)
 With Surf
 If .TotalPixels > 0 Then
   For DrawY = 0 To .Dims.HighM1
    For DrawX = 0 To .Dims.WideM1
    
    ''LDib is the array that bitblt copies from and blits to
    .LDib(DrawX, DrawY) = BGR
    '.sRGB(DrawX, DrawY).sBlu = Blue
    '.sRGB(DrawX, DrawY).sGrn = Green
    '.sRGB(DrawX, DrawY).sRed = Red
    
    Next
   Next
 End If
 End With
 
End Sub
Public Sub ClearSurface(Surf As AnimSurf2D)
 ClearSafeArrays Surf
 ClearDC Surf
 Erase Surf.sRGB
 Erase Surf.HorizRLenSegs
 Erase Surf.RLSeg
 Erase Surf.Clsn_Pattern
 Erase Surf.HSV1D
 Surf.Dims.bSRGB = False
 Surf.Dims.bHSV = False
 Surf.IA.Blues.Encoded = False
 Surf.IA.Greens.Encoded = False
 Surf.IA.Sats.Encoded = False
 Surf.IA.Hues.Encoded = False
 Surf.IA.Sats.Encoded = False
 Surf.IA.Vals.Encoded = False
 Surf.Dims.Wide = 0
 Surf.Dims.High = 0
 Surf.UB1D = -1
 Surf.TotalPixels = 0
End Sub
Private Sub ClearDC(Surf As AnimSurf2D)
     If (Surf.mem_hDC <> 0) Then
         If (Surf.mem_hDIb <> 0) Then
             SelectObject Surf.mem_hDC, Surf.mem_hBmpPrev
             DeleteObject Surf.mem_hDIb
         End If
         DeleteObject Surf.mem_hDC
     End If
     DeleteDC Surf.mem_hDC
     Surf.mem_hDC = 0: Surf.mem_hDIb = 0: Surf.mem_hBmpPrev = 0 ': .mem_Bits = 0
End Sub

Public Sub AlphaBlitRaw(SurfDest As AnimSurf2D, SurfSrc As AnimSurf2D, px!, py!, Optional DoTile As Boolean)
Dim DrawLeft& 'Destination
Dim DrawTop&
Dim DrawBot&

Dim AddDrawWidth&
Dim SrcClipLeft&
Dim SrcClipBot&
Dim SrX&
Dim SrY&
Dim sngAlpha!

Dim StartX& 'Tile
Dim StartY&
Dim EndX&
Dim EndY&
Dim LngX&
Dim LngY&

Dim H_Segs& 'advanced mask architecture
Dim SegsL&
Dim SLX&
Dim SLE&
Dim TSL&
Dim RB&

 If SurfDest.TotalPixels > 0 Then
 
 If DoTile Then
  StartX = Int(px - SurfSrc.wDiv2) Mod SurfSrc.Dims.Wide - SurfSrc.Dims.Wide
  StartY = Int(py - SurfSrc.hDiv2) Mod SurfSrc.Dims.High - SurfSrc.Dims.High
  EndX = SurfDest.Dims.Wide
  EndY = SurfDest.Dims.High
 Else
  StartX = Int(px - SurfSrc.wDiv2)
  StartY = Int(py - SurfSrc.hDiv2)
  EndX = StartX
  EndY = StartY
 End If
 
 For LngY = StartY To EndY Step SurfSrc.Dims.High
 For LngX = StartX To EndX Step SurfSrc.Dims.Wide
 
 DrawLeft = LngX
 DrawRight = DrawLeft + SurfSrc.Dims.WideM1
 DrawBot = LngY
 DrawTop = DrawBot + SurfSrc.Dims.HighM1
 
 If DrawLeft < 0 Then
  SrcClipLeft = -DrawLeft
  DrawLeft = 0
 Else
  SrcClipLeft = 0
 End If
 If DrawRight > SurfDest.Dims.WideM1 Then
  DrawRight = SurfDest.Dims.WideM1
 End If
 
 If DrawBot < 0 Then
  SrcClipBot = -DrawBot
  DrawBot = DrawLeft
 Else
  SrcClipBot = 0
  DrawBot = SurfDest.Dims.Wide * DrawBot + DrawLeft
 End If
 If DrawTop > SurfDest.Dims.HighM1 Then
  DrawTop = SurfDest.Dims.HighM1
 End If
 
 DrawTop = SurfDest.Dims.Wide * DrawTop + DrawLeft
 
 AddDrawWidth = DrawRight - DrawLeft
 
 SrY = SrcClipBot
 
 RB = SrcClipLeft + AddDrawWidth
 
 For DrawY = DrawBot To DrawTop Step SurfDest.Dims.Wide
  H_Segs = SurfSrc.HorizRLenSegs(SrY)
  TSL = 0
  DrawX = DrawY
  For SegsL = 0 To H_Segs Step 2
   SLE = TSL
   SrX = SurfSrc.RLSeg(SegsL, SrY)
   SLX = SLE + SrX 'draw start
   If SLX > SrcClipLeft Then
    If SLE <= SrcClipLeft Then
     DrawX = DrawX + SLX - SrcClipLeft
    Else
     DrawX = DrawX + SrX
    End If
    SLE = SLX + SurfSrc.RLSeg(SegsL + 1, SrY)
   Else
    SLE = SLX + SurfSrc.RLSeg(SegsL + 1, SrY)
    SLX = SrcClipLeft
   End If
   TSL = SLE + 1
   If SLE > RB Then SLE = RB
   For SrX = SLX To SLE
    'These are used to speed up rendering time slightly.
    'The result is a close approximation to the real thing.
    FGColor = SurfSrc.LDib(SrX, SrY)
    BGColor = SurfDest.LDib1D(DrawX)
    BGGrn = BGColor And vbGreen
    FGGrn = FGColor And vbGreen
    BGBlu = BGColor And &HFF&
    FGBlu = FGColor And &HFF&
    sngAlpha = SurfSrc.sRGB(SrX, SrY).aLph
    SurfDest.LDib1D(DrawX) = _
        BGColor + sngAlpha * (FGColor - BGColor) And vbBlue Or _
        BGGrn + sngAlpha * (FGGrn - BGGrn) And vbGreen Or _
        BGBlu + sngAlpha * (FGBlu - BGBlu)
    DrawX = DrawX + 1
   Next SrX
  Next SegsL
  SrY = SrY + 1
 Next DrawY
 
 Next LngX
 Next LngY
  
 End If 'Surf.TotalPixels > 0
  
End Sub
Public Sub RGBBlitRaw(SurfDest As AnimSurf2D, SurfSrc As AnimSurf2D, ByVal px!, ByVal py!, Optional ByVal Red As Byte = 255, Optional ByVal Green As Byte = 255, Optional ByVal Blue As Byte = 255, Optional ByVal Alpha As Byte = 255, Optional DoTile As Boolean = False)
Dim DrawLeft& 'Destination
Dim DrawRight&
Dim DrawTop&
Dim DrawBot&

Dim SrcClipLeft& 'Source
Dim SrcClipBot&
Dim AddDrawWidth&
Dim SrcX&
Dim SrcY&
Dim sngAlpha!
Dim sngAlpha2!
Dim sFGRed!
Dim sFGGrn!
Dim sFGBlu!

'horizontally-segmented blit architecture
Dim H_Segs&
Dim SegsL&
Dim SLX&
Dim SLE&
Dim TSL&
Dim RB&

'Tile
Dim XL&
Dim YL&
Dim x1&, x2&
Dim y1&, y2&

 If DoTile Then
  x1 = Int(px - SurfSrc.wDiv2) Mod SurfSrc.Dims.Wide - SurfSrc.Dims.Wide
  y1 = Int(py - SurfSrc.hDiv2) Mod SurfSrc.Dims.High - SurfSrc.Dims.High
  x2 = SurfDest.Dims.Wide
  y2 = SurfDest.Dims.High
 Else
  x1 = Int(px - SurfSrc.wDiv2)
  y1 = Int(py - SurfSrc.hDiv2)
  x2 = x1
  y2 = y1
 End If
 
 For YL = y1 To y2 Step SurfSrc.Dims.High
 For XL = x1 To x2 Step SurfSrc.Dims.Wide
 
 DrawLeft = XL
 DrawRight = DrawLeft + SurfSrc.Dims.WideM1
 DrawBot = YL
 DrawTop = DrawBot + SurfSrc.Dims.HighM1
 
 If DrawLeft < 0 Then
  SrcClipLeft = -DrawLeft
  DrawLeft = 0
 Else
  SrcClipLeft = 0
 End If
 If DrawRight > SurfDest.Dims.WideM1 Then
  DrawRight = SurfDest.Dims.WideM1
 End If
 
 If DrawBot < 0 Then
  SrcClipBot = -DrawBot
  DrawBot = DrawLeft
 Else
  SrcClipBot = 0
  DrawBot = SurfDest.Dims.Wide * DrawBot + DrawLeft
 End If
 If DrawTop > SurfDest.Dims.HighM1 Then
  DrawTop = SurfDest.Dims.HighM1
 End If
 
 DrawTop = SurfDest.Dims.Wide * DrawTop + DrawLeft
 
 AddDrawWidth = DrawRight - DrawLeft
 
 SrcY = SrcClipBot
 
 RB = SrcClipLeft + AddDrawWidth
 
 If Red = 255 And Green = 255 And Blue = 255 And Alpha = 255 Then
  
  For DrawY = DrawBot To DrawTop Step SurfDest.Dims.Wide
   H_Segs = SurfSrc.HorizRLenSegs(SrcY)
   TSL = 0
   DrawX = DrawY
   For SegsL = 0 To H_Segs Step 2
    SLE = TSL
    SrcX = SurfSrc.RLSeg(SegsL, SrcY)
    SLX = SLE + SrcX 'draw start
    If SLX > SrcClipLeft Then
     If SLE <= SrcClipLeft Then
      DrawX = DrawX + SLX - SrcClipLeft
     Else
      DrawX = DrawX + SrcX
     End If
     SLE = SLX + SurfSrc.RLSeg(SegsL + 1, SrcY)
    Else
     SLE = SLX + SurfSrc.RLSeg(SegsL + 1, SrcY)
     SLX = SrcClipLeft
    End If
    TSL = SLE + 1
    If SLE > RB Then SLE = RB
    For SrcX = SLX To SLE
     SurfDest.LDib1D(DrawX) = SurfSrc.LDib(SrcX, SrcY)
     DrawX = DrawX + 1
    Next SrcX
   Next SegsL
   SrcY = SrcY + 1
  Next DrawY
 
 Else 'one of the alphas is < 255
  
  If Alpha = 255 Then
  
  sFGRed = Red / 255
  sFGGrn = Green / 255
  sFGBlu = Blue / 255
  
  For DrawY = DrawBot To DrawTop Step SurfDest.Dims.Wide
   H_Segs = SurfSrc.HorizRLenSegs(SrcY)
   TSL = 0
   DrawX = DrawY
   For SegsL = 0 To H_Segs Step 2
    SLE = TSL
    SrcX = SurfSrc.RLSeg(SegsL, SrcY)
    SLX = SLE + SrcX 'draw start
    If SLX > SrcClipLeft Then
     If SLE <= SrcClipLeft Then
      DrawX = DrawX + SLX - SrcClipLeft
     Else
      DrawX = DrawX + SrcX
     End If
     SLE = SLX + SurfSrc.RLSeg(SegsL + 1, SrcY)
    Else
     SLE = SLX + SurfSrc.RLSeg(SegsL + 1, SrcY)
     SLX = SrcClipLeft
    End If
    TSL = SLE + 1
    If SLE > RB Then SLE = RB
    For SrcX = SLX To SLE
     FGColor = SurfSrc.LDib(SrcX, SrcY)
     SurfDest.LDib1D(DrawX) = _
      (FGColor And &HFF&) * sFGBlu Or _
      (FGColor And &HFF00&) * sFGGrn And &HFF00& Or _
      (FGColor And vbBlue) * sFGRed And vbBlue
     DrawX = DrawX + 1
    Next SrcX
   Next SegsL
   SrcY = SrcY + 1
  Next DrawY
  
  Else
  
  sngAlpha = Alpha / 255
  sngAlpha2 = 1 - sngAlpha
  
  sFGRed = sngAlpha * Red / 255
  sFGGrn = sngAlpha * Green / 255
  sFGBlu = sngAlpha * Blue / 255
  
  For DrawY = DrawBot To DrawTop Step SurfDest.Dims.Wide
   H_Segs = SurfSrc.HorizRLenSegs(SrcY)
   TSL = 0
   DrawX = DrawY
   For SegsL = 0 To H_Segs Step 2
    SLE = TSL
    SrcX = SurfSrc.RLSeg(SegsL, SrcY)
    SLX = SLE + SrcX 'draw start
    If SLX > SrcClipLeft Then
     If SLE <= SrcClipLeft Then
      DrawX = DrawX + SLX - SrcClipLeft
     Else
      DrawX = DrawX + SrcX
     End If
     SLE = SLX + SurfSrc.RLSeg(SegsL + 1, SrcY)
    Else
     SLE = SLX + SurfSrc.RLSeg(SegsL + 1, SrcY)
     SLX = SrcClipLeft
    End If
    TSL = SLE + 1
    If SLE > RB Then SLE = RB
    For SrcX = SLX To SLE
     BGColor = SurfDest.LDib1D(DrawX)
     FGColor = SurfSrc.LDib(SrcX, SrcY)
     SurfDest.LDib1D(DrawX) = _
      (FGColor And &HFF&) * sFGBlu + (BGColor And &HFF&) * sngAlpha2 Or _
      (FGColor And &HFF00&) * sFGGrn + (BGColor And &HFF00&) * sngAlpha2 And &HFF00& Or _
      (FGColor And vbBlue) * sFGRed + (BGColor And vbBlue) * sngAlpha2 And vbBlue
     DrawX = DrawX + 1
    Next SrcX
   Next SegsL
   SrcY = SrcY + 1
  Next DrawY
  
  End If 'Alpha = 255
 
 End If
 
 Next XL
 Next YL
 
End Sub
Public Sub SetDims(Surf As AnimSurf2D, Wide&, High&)

 Surf.Dims.Wide = Wide
 Surf.Dims.High = High

 Surf.Dims.WideM1 = Surf.Dims.Wide - 1
 Surf.Dims.HighM1 = Surf.Dims.High - 1
 
 Surf.wDiv2 = Surf.Dims.Wide / 2&
 Surf.hDiv2 = Surf.Dims.High / 2&
 
 Surf.TotalPixels = Surf.Dims.Wide * Surf.Dims.High
 Surf.UB1D = Surf.TotalPixels - 1
 
 Surf.TopLeft = Surf.TotalPixels - Surf.Dims.Wide
 
End Sub
Public Sub ClearSurfaceAuto()
Dim I2&
 For I2 = 1 To SurfAutoUsed
  'With SurfAuto(i)
  'SelectObject .mem_hDC, .mem_hBmpPrev
  'End With
  ClearSurface SurfAuto(I2)
 Next I2
End Sub
Private Sub ClearSafeArrays(Surf As AnimSurf2D)
 CopyMemory ByVal VarPtrArray(Surf.LDib), 0&, 4
 CopyMemory ByVal VarPtrArray(Surf.Dib), 0&, 4
 CopyMemory ByVal VarPtrArray(Surf.LDib1D), 0&, 4
 CopyMemory ByVal VarPtrArray(Surf.Dib1D), 0&, 4
 CopyMemory ByVal VarPtrArray(Surf.sRGB1D), 0&, 4
End Sub


' Run-Length processing subs

Public Sub RLEncode(Surf As AnimSurf2D, Optional ByVal HSV_RGB_or_Both__0_To_2& = 0)
Dim ValueIsPresent1() As Boolean
Dim ValueIsPresent2() As Boolean
Dim ValueIsPresent3() As Boolean
Dim HSV1 As HSVTYPE

    If Surf.TotalPixels > 0 Then
    
        Surf.IA.WhichAreEncoded = HSV_RGB_or_Both__0_To_2

        If HSV_RGB_or_Both__0_To_2 > 0 Then

            ReDim ValueIsPresent1(255)
            ReDim ValueIsPresent2(255)
            ReDim ValueIsPresent3(255)
            
            'If Not Surf.Dims.bSRGB Then
            '    ReDim Surf.sRGB(Surf.Dims.WideM1, Surf.Dims.HighM1)
            '    Surf.Dims.bSRGB = True
            '    CopyMemory ByVal VarPtrArray(Surf.sRGB1D), 0&, 4
            '    With Surf.SafeAry1D_sRGB
            '        .cbElements = 16
            '        .cDims = 1
            '        .lLbound = 0
            '        .cElements = Surf.TotalPixels
            '        .pvData = VarPtr(Surf.sRGB(0, 0).sRed)
            '    End With
            '    CopyMemory ByVal VarPtrArray(Surf.sRGB1D), VarPtr(Surf.SafeAry1D_sRGB), 4
            'End If
            
            For DrawX = 0 To Surf.UB1D
                ValueIsPresent1(Surf.Dib1D(DrawX).Red) = True
                ValueIsPresent2(Surf.Dib1D(DrawX).Green) = True
                ValueIsPresent3(Surf.Dib1D(DrawX).Blue) = True
            Next DrawX
            
            BuildTerraces ValueIsPresent1, Surf, Surf.IA.Reds, 255
            BuildTerraces ValueIsPresent2, Surf, Surf.IA.Greens, 255
            BuildTerraces ValueIsPresent3, Surf, Surf.IA.Blues, 255
        
        End If
        
        If HSV_RGB_or_Both__0_To_2 <> 1 Then
        
            If Not Surf.Dims.bHSV Then
                Surf.Dims.bHSV = True
            End If
            ReDim Surf.HSV1D(Surf.UB1D)
        
            ReDim ValueIsPresent1(255 + 4 * 255 + 254)
            ReDim ValueIsPresent2(255)
            ReDim ValueIsPresent3(255)
            
            For DrawX = 0 To Surf.UB1D
                HSVTYPE_From_RGBQUAD Surf.HSV1D(DrawX), Surf.Dib1D(DrawX)
                HSV1 = Surf.HSV1D(DrawX)
                ValueIsPresent1(Int(HSV1.h_hue + 0.5)) = True
                ValueIsPresent2(Int(255 * HSV1.s_saturation + 0.5)) = True
                ValueIsPresent3(HSV1.v_value) = True
            Next DrawX
 
            BuildTerraces ValueIsPresent1, Surf, Surf.IA.Hues, 255 + 4 * 255 + 254
            BuildTerraces ValueIsPresent2, Surf, Surf.IA.Sats, 255
            BuildTerraces ValueIsPresent3, Surf, Surf.IA.Vals, 255
        
        End If

        EncodeComponents Surf, Surf.IA
    
    End If

End Sub
Private Sub BuildTerraces(Booleans() As Boolean, Surf As AnimSurf2D, Terraces As Terraces2, TerraceCount As Long)
Dim PrevPresent As Boolean
Dim CurrPresent As Boolean
Dim Vert_Start&
Dim Vert_Target&
Dim VLenM1&
Dim LayerRef&
Dim RLS&
Dim RLE&
Dim I2&, J2&, L1&, I4&
Dim Node1&, Node2&
Dim TCP1&

    ReDim Terraces.TerraceRef(TerraceCount)
    TCP1 = TerraceCount + 1
    Terraces.UB_Tr = -1
    
    Do While RLS < TCP1
        For I4 = Node2 To TerraceCount
            If Booleans(I4) Then
                Node1 = I4
                Exit For
            End If
        Next I4
        RLE = TerraceCount
        Terraces.UB_Tr = Terraces.UB_Tr + 1
        For J2 = Node1 + 1 To TerraceCount
            If Booleans(J2) Then
                Node2 = J2
                RLE = Int((Node1 + Node2) / 2)
                Exit For
            End If
        Next J2
        For L1 = RLS To RLE
            Terraces.TerraceRef(L1) = Terraces.UB_Tr
        Next L1
        RLS = RLE + 1
        Node1 = Node2
        RLE = RLS
    Loop
    
    ReDim Terraces.Terrace(Terraces.LB_Tr To Terraces.UB_Tr)
    Terraces.UB_Tr = -1
    
    For Layer = 0 To TerraceCount
        If Booleans(Layer) Then
            Terraces.UB_Tr = Terraces.UB_Tr + 1
            Terraces.Terrace(Terraces.UB_Tr).Alpha = Layer
            Terraces.Terrace(Terraces.UB_Tr).AlphaOrigin = Layer
        End If
    Next Layer
    
    For LayerRef = Terraces.LB_Tr To Terraces.UB_Tr
        Terraces.Terrace(LayerRef).UB_Lengths = -1
    Next LayerRef
    
    Terraces.Encoded = True

End Sub
Private Sub EncodeComponents(Surf As AnimSurf2D, RLHV1 As RunLengthHeightField)
Dim StartPos&
Dim CurrentVal1&
Dim PreviousVal1&
Dim CurrentVal2&
Dim PreviousVal2&
Dim CurrentVal3&
Dim PreviousVal3&
Dim HSV1 As HSVTYPE
 
    If Surf.IA.WhichAreEncoded > 0 Then
    
        PreviousVal1 = -1
        PreviousVal2 = -1
        PreviousVal3 = -1
    
        For DrawX = 0 To Surf.UB1D
            CurrentVal1 = Surf.Dib1D(DrawX).Red
            CurrentVal2 = Surf.Dib1D(DrawX).Green
            CurrentVal3 = Surf.Dib1D(DrawX).Blue
            CVA RLHV1.Reds, CurrentVal1, PreviousVal1
            CVA RLHV1.Greens, CurrentVal2, PreviousVal2
            CVA RLHV1.Blues, CurrentVal3, PreviousVal3
        Next DrawX
        
        CVT RLHV1.Reds
        CVT RLHV1.Greens
        CVT RLHV1.Blues
        
        PreviousVal1 = -1
        PreviousVal2 = -1
        PreviousVal3 = -1
        
        For DrawX = 0 To Surf.UB1D
            CurrentVal1 = Surf.Dib1D(DrawX).Red
            CurrentVal2 = Surf.Dib1D(DrawX).Green
            CurrentVal3 = Surf.Dib1D(DrawX).Blue
            CVB RLHV1.Reds, CurrentVal1, PreviousVal1
            CVB RLHV1.Greens, CurrentVal2, PreviousVal2
            CVB RLHV1.Blues, CurrentVal3, PreviousVal3
        Next DrawX
    End If
    
    If Surf.IA.WhichAreEncoded <> 1 Then
    
        PreviousVal1 = -1
        PreviousVal2 = -1
        PreviousVal3 = -1
        
        For DrawX = 0 To Surf.UB1D
            HSV1 = Surf.HSV1D(DrawX)
            CurrentVal1 = Int(HSV1.h_hue + 0.5)
            CurrentVal2 = Int(255 * HSV1.s_saturation + 0.5)
            CurrentVal3 = HSV1.v_value
            CVA RLHV1.Hues, CurrentVal1, PreviousVal1
            CVA RLHV1.Sats, CurrentVal2, PreviousVal2
            CVA RLHV1.Vals, CurrentVal3, PreviousVal3
        Next DrawX
        
        CVT RLHV1.Hues
        CVT RLHV1.Sats
        CVT RLHV1.Vals
        
        PreviousVal1 = -1
        PreviousVal2 = -1
        PreviousVal3 = -1
        
        For DrawX = 0 To Surf.UB1D
            'HSVTYPE_From_RGBQUAD HSV1, Surf.Dib1D(DrawX)
            HSV1 = Surf.HSV1D(DrawX)
            CurrentVal1 = Int(HSV1.h_hue + 0.5)
            CurrentVal2 = Int(255 * HSV1.s_saturation + 0.5)
            CurrentVal3 = HSV1.v_value
            CVB RLHV1.Hues, CurrentVal1, PreviousVal1
            CVB RLHV1.Sats, CurrentVal2, PreviousVal2
            CVB RLHV1.Vals, CurrentVal3, PreviousVal3
        Next DrawX
        
    End If
    
End Sub
Private Sub CVT(Terrace1 As Terraces2)
    For DrawY = Terrace1.LB_Tr To Terrace1.UB_Tr
        RedimLayer1 Terrace1.Terrace(DrawY)
    Next DrawY
End Sub
Private Sub CVA(Terrace1 As Terraces2, CV1&, PV1&)
    If CV1 <> PV1 Then
        IncUBLoc Terrace1.Terrace(Terrace1.TerraceRef(CV1))
        PV1 = CV1
    End If
End Sub
Private Sub CVB(Terrace1 As Terraces2, CV1&, PV1&)
Dim Ref1&
    Ref1 = Terrace1.TerraceRef(CV1)
    If CV1 = PV1 Then
        AddLength Terrace1.Terrace(Ref1)
    Else
        NewVal Terrace1.Terrace(Ref1), DrawX
    End If
    PV1 = CV1
End Sub
Private Sub RedimLayer1(Layer1 As LayerConstruct)
    ReDim Layer1.Piece(Layer1.LB_Lengths To Layer1.UB_Lengths)
    Layer1.UB_Lengths = -1
End Sub
Private Sub IncUBLoc(Layer1 As LayerConstruct)
    Layer1.UB_Lengths = Layer1.UB_Lengths + 1
End Sub
Private Sub NewVal(Layer1 As LayerConstruct, LocationStart&)
    Layer1.UB_Lengths = Layer1.UB_Lengths + 1
    Layer1.Piece(Layer1.UB_Lengths).LStart = LocationStart
    Layer1.Piece(Layer1.UB_Lengths).LFinal = LocationStart
End Sub
Private Sub AddLength(Layer1 As LayerConstruct)
    Layer1.Piece(Layer1.UB_Lengths).LFinal = Layer1.Piece(Layer1.UB_Lengths).LFinal + 1
End Sub
Public Sub RLDecode(Surf As AnimSurf2D, Optional AnimateRestoreOriginal As Boolean, Optional hToDC&, Optional RestorePicButKeepTerraceVals As Boolean = False)
Dim Seg1&
Dim Vert_Start&
Dim Vert_Target&
Dim LayerPtr&

    If Surf.IA.Reds.Encoded Then
    
        If AnimateRestoreOriginal Then
    
            For Seg1 = Surf.IA.Reds.LB_Tr To Surf.IA.Reds.UB_Tr
                Surf.IA.Reds.Terrace(Seg1).Alpha = Surf.IA.Reds.Terrace(Seg1).AlphaOrigin
                DecodeTerrace Surf, Surf.IA.Reds.Terrace(Seg1), RedLayer
                BlitToDC hToDC, Surf
            Next Seg1
            For Seg1 = Surf.IA.Greens.LB_Tr To Surf.IA.Greens.UB_Tr
                Surf.IA.Greens.Terrace(Seg1).Alpha = Surf.IA.Greens.Terrace(Seg1).AlphaOrigin
                DecodeTerrace Surf, Surf.IA.Greens.Terrace(Seg1), GreenLayer
                BlitToDC hToDC, Surf
            Next Seg1
            For Seg1 = Surf.IA.Blues.LB_Tr To Surf.IA.Blues.UB_Tr
                Surf.IA.Blues.Terrace(Seg1).Alpha = Surf.IA.Blues.Terrace(Seg1).AlphaOrigin
                DecodeTerrace Surf, Surf.IA.Blues.Terrace(Seg1), BlueLayer
                BlitToDC hToDC, Surf
            Next Seg1
    
        Else 'AnimateRestoreOriginal = False
        
            If RestorePicButKeepTerraceVals Then
            
                For Seg1 = Surf.IA.Reds.LB_Tr To Surf.IA.Reds.UB_Tr
                    DecodeTerraceOrigin Surf, Surf.IA.Reds.Terrace(Seg1), RedLayer
                Next Seg1
                For Seg1 = Surf.IA.Greens.LB_Tr To Surf.IA.Greens.UB_Tr
                    DecodeTerraceOrigin Surf, Surf.IA.Greens.Terrace(Seg1), GreenLayer
                Next Seg1
                For Seg1 = Surf.IA.Blues.LB_Tr To Surf.IA.Blues.UB_Tr
                    DecodeTerraceOrigin Surf, Surf.IA.Blues.Terrace(Seg1), BlueLayer
                Next Seg1
            
            Else

                For Seg1 = Surf.IA.Reds.LB_Tr To Surf.IA.Reds.UB_Tr
                    DecodeTerrace Surf, Surf.IA.Reds.Terrace(Seg1), RedLayer
                Next Seg1
                For Seg1 = Surf.IA.Greens.LB_Tr To Surf.IA.Greens.UB_Tr
                    DecodeTerrace Surf, Surf.IA.Greens.Terrace(Seg1), GreenLayer
                Next Seg1
                For Seg1 = Surf.IA.Blues.LB_Tr To Surf.IA.Blues.UB_Tr
                    DecodeTerrace Surf, Surf.IA.Blues.Terrace(Seg1), BlueLayer
                Next Seg1
            
            End If
            
        End If 'AnimateRestoreOriginal
    
    End If

End Sub
Private Sub DecodeTerrace(Surf As AnimSurf2D, Layer1 As LayerConstruct, Component_1Red_To_Blue3 As Long)
Dim Seg1&
Dim I2&

    If Surf.TotalPixels > 0 Then
    
    If Component_1Red_To_Blue3 = 1 Then
    
        If Surf.IA.Reds.Encoded Then
            For Seg1 = Layer1.LB_Lengths To Layer1.UB_Lengths
                DrawY = Layer1.Piece(Seg1).LStart
                DrawRight = Layer1.Piece(Seg1).LFinal
                For I2 = DrawY To DrawRight
                    Surf.Dib1D(I2).Red = Layer1.Alpha
                    'Surf.sRGB1D(I2).sRed = Layer1.Alpha
                    HSVTYPE_From_RGBQUAD_P Surf.HSV1D(I2), Surf.Dib1D(I2)
                Next I2
            Next Seg1
        End If
        
    ElseIf Component_1Red_To_Blue3 = 2 Then
    
        If Surf.IA.Greens.Encoded Then
            For Seg1 = Layer1.LB_Lengths To Layer1.UB_Lengths
                DrawY = Layer1.Piece(Seg1).LStart
                DrawRight = Layer1.Piece(Seg1).LFinal
                For I2 = DrawY To DrawRight
                    Surf.Dib1D(I2).Green = Layer1.Alpha
                    'Surf.sRGB1D(I2).sGrn = Layer1.Alpha
                    HSVTYPE_From_RGBQUAD_P Surf.HSV1D(I2), Surf.Dib1D(I2)
                Next I2
            Next Seg1
        End If
        
    Else
    
        If Surf.IA.Blues.Encoded Then
            For Seg1 = Layer1.LB_Lengths To Layer1.UB_Lengths
                DrawY = Layer1.Piece(Seg1).LStart
                DrawRight = Layer1.Piece(Seg1).LFinal
                For I2 = DrawY To DrawRight
                    Surf.Dib1D(I2).Blue = Layer1.Alpha
                    'Surf.sRGB1D(I2).sBlu = Layer1.Alpha
                    HSVTYPE_From_RGBQUAD_P Surf.HSV1D(I2), Surf.Dib1D(I2)
                Next I2
            Next Seg1
        End If
        
    End If
    
    End If

End Sub
Private Sub DecodeTerraceOrigin(Surf As AnimSurf2D, Layer1 As LayerConstruct, Component_1Red_To_Blue3 As Long)
Dim Seg1&
Dim I2&

    If Surf.TotalPixels > 0 Then
    
    If Component_1Red_To_Blue3 = 1 Then
    
        If Surf.IA.Reds.Encoded Then
            For Seg1 = Layer1.LB_Lengths To Layer1.UB_Lengths
                DrawY = Layer1.Piece(Seg1).LStart
                DrawRight = Layer1.Piece(Seg1).LFinal
                For I2 = DrawY To DrawRight
                    Surf.Dib1D(I2).Red = Layer1.AlphaOrigin
                    'Surf.sRGB1D(I2).sRed = Layer1.AlphaOrigin
                    HSVTYPE_From_RGBQUAD_P Surf.HSV1D(I2), Surf.Dib1D(I2)
                Next I2
            Next Seg1
        End If
        
    ElseIf Component_1Red_To_Blue3 = 2 Then
    
        If Surf.IA.Greens.Encoded Then
            For Seg1 = Layer1.LB_Lengths To Layer1.UB_Lengths
                DrawY = Layer1.Piece(Seg1).LStart
                DrawRight = Layer1.Piece(Seg1).LFinal
                For I2 = DrawY To DrawRight
                    Surf.Dib1D(I2).Green = Layer1.AlphaOrigin
                    'Surf.sRGB1D(I2).sGrn = Layer1.AlphaOrigin
                    HSVTYPE_From_RGBQUAD_P Surf.HSV1D(I2), Surf.Dib1D(I2)
                Next I2
            Next Seg1
        End If
        
    Else
    
        If Surf.IA.Blues.Encoded Then
            For Seg1 = Layer1.LB_Lengths To Layer1.UB_Lengths
                DrawY = Layer1.Piece(Seg1).LStart
                DrawRight = Layer1.Piece(Seg1).LFinal
                For I2 = DrawY To DrawRight
                    Surf.Dib1D(I2).Blue = Layer1.AlphaOrigin
                    'Surf.sRGB1D(I2).sBlu = Layer1.AlphaOrigin
                    HSVTYPE_From_RGBQUAD_P Surf.HSV1D(I2), Surf.Dib1D(I2)
                Next I2
            Next Seg1
        End If
        
    End If
    
    End If

End Sub

Public Sub SetTerraceVal(Surf1 As AnimSurf2D, ByVal layer_0_To_255 As Single, ByVal alpha_0_To_255 As Single, Lng_1Red_To_3Blue&)
Dim LngLayer&
Dim AlphaRef1&

    layer_0_To_255 = Int(layer_0_To_255 Mod 256 + 0.5)
    alpha_0_To_255 = Int(alpha_0_To_255 Mod 256 + 0.5)

    If Lng_1Red_To_3Blue = 1 Then
    
        If Surf1.IA.Reds.Encoded Then
            LngLayer = Surf1.IA.Reds.TerraceRef(layer_0_To_255)
            AlphaRef1 = alpha_0_To_255
            If AlphaRef1 <> Surf1.IA.Reds.Terrace(LngLayer).Alpha Then
                Surf1.IA.Reds.Terrace(LngLayer).Alpha = AlphaRef1
                DecodeTerrace Surf1, Surf1.IA.Reds.Terrace(LngLayer), 1
            End If
        End If
        
    ElseIf Lng_1Red_To_3Blue = 2 Then
    
        If Surf1.IA.Greens.Encoded Then
            LngLayer = Surf1.IA.Greens.TerraceRef(layer_0_To_255)
            AlphaRef1 = alpha_0_To_255
            If AlphaRef1 <> Surf1.IA.Greens.Terrace(LngLayer).Alpha Then
                Surf1.IA.Greens.Terrace(LngLayer).Alpha = AlphaRef1
                DecodeTerrace Surf1, Surf1.IA.Greens.Terrace(LngLayer), 2
            End If
        End If
        
    ElseIf Lng_1Red_To_3Blue = 3 Then
    
        If Surf1.IA.Blues.Encoded Then
            LngLayer = Surf1.IA.Blues.TerraceRef(layer_0_To_255)
            AlphaRef1 = alpha_0_To_255
            If AlphaRef1 <> Surf1.IA.Blues.Terrace(LngLayer).Alpha Then
                Surf1.IA.Blues.Terrace(LngLayer).Alpha = AlphaRef1
                DecodeTerrace Surf1, Surf1.IA.Blues.Terrace(LngLayer), 3
            End If
        End If
        
    End If

End Sub


Public Sub RLDecodeHSV(Surf As AnimSurf2D, Optional AnimateRestoreOriginal As Boolean, Optional hToDC&)
Dim Seg1&
Dim Vert_Start&
Dim Vert_Target&
Dim LayerPtr&

    If Surf.IA.Hues.Encoded Then
    
        If AnimateRestoreOriginal Then
    
            For Seg1 = Surf.IA.Hues.LB_Tr To Surf.IA.Hues.UB_Tr
                Surf.IA.Hues.Terrace(Seg1).Alpha = Surf.IA.Hues.Terrace(Seg1).AlphaOrigin
                DecodeTerraceHSV Surf, Surf.IA.Hues.Terrace(Seg1), RedLayer
                BlitToDC hToDC, Surf
                DoEvents
            Next Seg1
            For Seg1 = Surf.IA.Sats.LB_Tr To Surf.IA.Sats.UB_Tr
                Surf.IA.Sats.Terrace(Seg1).Alpha = Surf.IA.Sats.Terrace(Seg1).AlphaOrigin
                DecodeTerraceHSV Surf, Surf.IA.Sats.Terrace(Seg1), GreenLayer
                BlitToDC hToDC, Surf
                DoEvents
            Next Seg1
            For Seg1 = Surf.IA.Vals.LB_Tr To Surf.IA.Vals.UB_Tr
                Surf.IA.Vals.Terrace(Seg1).Alpha = Surf.IA.Vals.Terrace(Seg1).AlphaOrigin
                DecodeTerraceHSV Surf, Surf.IA.Vals.Terrace(Seg1), BlueLayer
                BlitToDC hToDC, Surf
                DoEvents
            Next Seg1
    
        Else 'AnimateRestoreOriginal = False

            For Seg1 = Surf.IA.Hues.LB_Tr To Surf.IA.Hues.UB_Tr
                DecodeTerraceHSV Surf, Surf.IA.Hues.Terrace(Seg1), RedLayer
            Next Seg1
            For Seg1 = Surf.IA.Sats.LB_Tr To Surf.IA.Sats.UB_Tr
                DecodeTerraceHSV Surf, Surf.IA.Sats.Terrace(Seg1), GreenLayer
            Next Seg1
            For Seg1 = Surf.IA.Vals.LB_Tr To Surf.IA.Vals.UB_Tr
                DecodeTerraceHSV Surf, Surf.IA.Vals.Terrace(Seg1), BlueLayer
            Next Seg1
            
        End If 'AnimateRestoreOriginal
    
    End If

End Sub
Private Sub DecodeTerraceHSV(Surf As AnimSurf2D, Layer1 As LayerConstruct, Component_1Hue_To_Sat3 As Long)
Dim Seg1&
Dim I2&
Dim HSV1 As HSVTYPE
Dim RGB1 As RGBQUAD

    If Surf.TotalPixels > 0 Then
    
    If Component_1Hue_To_Sat3 = 1 Then
    
        If Surf.IA.Hues.Encoded Then
            For Seg1 = Layer1.LB_Lengths To Layer1.UB_Lengths
                DrawY = Layer1.Piece(Seg1).LStart
                DrawRight = Layer1.Piece(Seg1).LFinal
                For I2 = DrawY To DrawRight
                    Surf.HSV1D(I2).h_hue = Layer1.Alpha
                    RGBQUAD_From_HSVTYPE_P Surf.Dib1D(I2), Surf.HSV1D(I2)
                Next I2
            Next Seg1
        End If
        
    ElseIf Component_1Hue_To_Sat3 = 2 Then
    
        If Surf.IA.Sats.Encoded Then
            For Seg1 = Layer1.LB_Lengths To Layer1.UB_Lengths
                DrawY = Layer1.Piece(Seg1).LStart
                DrawRight = Layer1.Piece(Seg1).LFinal
                For I2 = DrawY To DrawRight
                    Surf.HSV1D(I2).s_saturation = Layer1.Alpha / 255&
                    RGBQUAD_From_HSVTYPE_P Surf.Dib1D(I2), Surf.HSV1D(I2)
                Next I2
            Next Seg1
        End If
        
    Else
    
        If Surf.IA.Vals.Encoded Then
            For Seg1 = Layer1.LB_Lengths To Layer1.UB_Lengths
                DrawY = Layer1.Piece(Seg1).LStart
                DrawRight = Layer1.Piece(Seg1).LFinal
                For I2 = DrawY To DrawRight
                    Surf.HSV1D(I2).v_value = Layer1.Alpha
                    RGBQUAD_From_HSVTYPE_P Surf.Dib1D(I2), Surf.HSV1D(I2)
                Next I2
            Next Seg1
        End If
        
    End If
    
    End If

End Sub

Public Sub SetTerraceValHSV(Surf1 As AnimSurf2D, ByVal layer_0_To_255or1529 As Single, ByVal alpha_0_To_255or1529 As Single, Lng_1Hue_To_3Val&)
Dim LngLayer&
Dim AlphaRef1&

    If Lng_1Hue_To_3Val = 1 Then
        layer_0_To_255or1529 = layer_0_To_255or1529 Mod 1530
        alpha_0_To_255or1529 = alpha_0_To_255or1529 Mod 1530
    Else
        layer_0_To_255or1529 = layer_0_To_255or1529 Mod 256
        alpha_0_To_255or1529 = alpha_0_To_255or1529 Mod 256
    End If
    
    layer_0_To_255or1529 = Int(layer_0_To_255or1529 + 0.5)
    alpha_0_To_255or1529 = Int(alpha_0_To_255or1529 + 0.5)

    If Lng_1Hue_To_3Val = 1 Then
     
        If Surf1.IA.Hues.Encoded Then
            LngLayer = Surf1.IA.Hues.TerraceRef(layer_0_To_255or1529)
            AlphaRef1 = alpha_0_To_255or1529
            If AlphaRef1 <> Surf1.IA.Hues.Terrace(LngLayer).Alpha Then
                Surf1.IA.Hues.Terrace(LngLayer).Alpha = AlphaRef1
                DecodeTerraceHSV Surf1, Surf1.IA.Hues.Terrace(LngLayer), 1
            End If
        End If
        
    ElseIf Lng_1Hue_To_3Val = 2 Then
    
        If Surf1.IA.Sats.Encoded Then
            LngLayer = Surf1.IA.Sats.TerraceRef(layer_0_To_255or1529)
            AlphaRef1 = alpha_0_To_255or1529
            If AlphaRef1 <> Surf1.IA.Sats.Terrace(LngLayer).Alpha Then
                Surf1.IA.Sats.Terrace(LngLayer).Alpha = AlphaRef1
                DecodeTerraceHSV Surf1, Surf1.IA.Sats.Terrace(LngLayer), 2
            End If
        End If
        
    ElseIf Lng_1Hue_To_3Val = 3 Then
    
        If Surf1.IA.Vals.Encoded Then
            LngLayer = Surf1.IA.Vals.TerraceRef(layer_0_To_255or1529)
            AlphaRef1 = alpha_0_To_255or1529
            If AlphaRef1 <> Surf1.IA.Vals.Terrace(LngLayer).Alpha Then
                Surf1.IA.Vals.Terrace(LngLayer).Alpha = AlphaRef1
                DecodeTerraceHSV Surf1, Surf1.IA.Vals.Terrace(LngLayer), 3
            End If
        End If
        
    End If

End Sub

Private Sub HSVTYPE_From_RGBQUAD_P(HSV1 As HSVTYPE, RGB1 As RGBQUAD)
 If RGB1.Red >= RGB1.Green Then
  If RGB1.Green >= RGB1.Blue Then '1. G +
   HSV1.v_value = RGB1.Red
   HSV1.s_saturation = RGB1.Red - RGB1.Blue
   If HSV1.s_saturation > 0 Then
    HSV1.h_hue = 255& * (RGB1.Green - RGB1.Blue) / HSV1.s_saturation
   End If
  Else '  R -> B
   If RGB1.Red <= RGB1.Blue Then
    HSV1.v_value = RGB1.Blue
    HSV1.s_saturation = RGB1.Blue - RGB1.Green
    If HSV1.s_saturation > 0 Then
     HSV1.h_hue = 1020 + 255& * (RGB1.Red - RGB1.Green) / HSV1.s_saturation
    End If
   Else
    HSV1.v_value = RGB1.Red
    HSV1.s_saturation = RGB1.Red - RGB1.Green
    If HSV1.s_saturation > 0 Then
     HSV1.h_hue = 1275 + 255& * (RGB1.Red - RGB1.Blue) / HSV1.s_saturation
    End If
   End If
  End If 'G >= B
 Else 'R < G
  If RGB1.Red >= RGB1.Blue Then ' 2. r-
   HSV1.v_value = RGB1.Green
   HSV1.s_saturation = RGB1.Green - RGB1.Blue
   If HSV1.s_saturation > 0 Then
    HSV1.h_hue = 255 + 255& * (RGB1.Green - RGB1.Red) / HSV1.s_saturation
   End If
  Else
   If RGB1.Green >= RGB1.Blue Then '3. B+
    HSV1.v_value = RGB1.Green
    HSV1.s_saturation = RGB1.Green - RGB1.Red
    If HSV1.s_saturation > 0 Then
     HSV1.h_hue = 510 + 255& * (RGB1.Blue - RGB1.Red) / HSV1.s_saturation
    End If
   Else '4. G-
    HSV1.v_value = RGB1.Blue
    HSV1.s_saturation = RGB1.Blue - RGB1.Red
    If HSV1.s_saturation > 0 Then
     HSV1.h_hue = 765 + 255& * (RGB1.Blue - RGB1.Green) / HSV1.s_saturation
    End If
   End If
  End If
 End If
 If HSV1.v_value > 0 Then
  HSV1.s_saturation = HSV1.s_saturation / HSV1.v_value
 End If
End Sub

Public Sub HSVTYPE_From_RGBQUAD(HSV1 As HSVTYPE, RGB1 As RGBQUAD)
 If RGB1.Red >= RGB1.Green Then
  If RGB1.Green >= RGB1.Blue Then '1. G +
   HSV1.v_value = RGB1.Red
   HSV1.s_saturation = RGB1.Red - RGB1.Blue
   If HSV1.s_saturation > 0 Then
    HSV1.h_hue = 255& * (RGB1.Green - RGB1.Blue) / HSV1.s_saturation
   Else
    HSV1.h_hue = 0
   End If
  Else '  R -> B
   If RGB1.Red <= RGB1.Blue Then
    HSV1.v_value = RGB1.Blue
    HSV1.s_saturation = RGB1.Blue - RGB1.Green
    If HSV1.s_saturation > 0 Then
     HSV1.h_hue = 1020 + 255& * (RGB1.Red - RGB1.Green) / HSV1.s_saturation
    Else
     HSV1.h_hue = 0
    End If
   Else
    HSV1.v_value = RGB1.Red
    HSV1.s_saturation = RGB1.Red - RGB1.Green
    If HSV1.s_saturation > 0 Then
     HSV1.h_hue = 1275 + 255& * (RGB1.Red - RGB1.Blue) / HSV1.s_saturation
    Else
     HSV1.h_hue = 0
    End If
   End If
  End If 'G >= B
 Else 'R < G
  If RGB1.Red >= RGB1.Blue Then ' 2. r-
   HSV1.v_value = RGB1.Green
   HSV1.s_saturation = RGB1.Green - RGB1.Blue
   If HSV1.s_saturation > 0 Then
    HSV1.h_hue = 255 + 255& * (RGB1.Green - RGB1.Red) / HSV1.s_saturation
   Else
    HSV1.h_hue = 0
   End If
  Else
   If RGB1.Green >= RGB1.Blue Then '3. B+
    HSV1.v_value = RGB1.Green
    HSV1.s_saturation = RGB1.Green - RGB1.Red
    If HSV1.s_saturation > 0 Then
     HSV1.h_hue = 510 + 255& * (RGB1.Blue - RGB1.Red) / HSV1.s_saturation
    Else
     HSV1.h_hue = 0
    End If
   Else '4. G-
    HSV1.v_value = RGB1.Blue
    HSV1.s_saturation = RGB1.Blue - RGB1.Red
    If HSV1.s_saturation > 0 Then
     HSV1.h_hue = 765 + 255& * (RGB1.Blue - RGB1.Green) / HSV1.s_saturation
    Else
     HSV1.h_hue = 0
    End If
   End If
  End If
 End If
 If HSV1.v_value > 0 Then
  HSV1.s_saturation = HSV1.s_saturation / HSV1.v_value
 End If
End Sub



Public Sub RGBQUAD_From_HSV(RGB1 As RGBQUAD, hue_0_To_1530!, ByVal saturation_0_To_1!, value_0_To_255!)
Dim hue_and_sat!

 If value_0_To_255 > 0 Then
  value1 = value_0_To_255 + 0.5
  If saturation_0_To_1 > 0 Then
   maxim = hue_0_To_1530 - 1530& * Int(hue_0_To_1530 / 1530&)
   diff1 = saturation_0_To_1 * value_0_To_255
   subt = value1 - diff1
   diff1 = diff1 / 255
   If maxim <= 510 Then
    RGB1.Blue = Int(subt)
    If maxim <= 255 Then
     hue_and_sat = maxim * diff1!
     RGB1.Red = Int(value1)
     RGB1.Green = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 255) * diff1!
     RGB1.Green = Int(value1)
     RGB1.Red = Int(value1 - hue_and_sat)
    End If
   ElseIf maxim <= 1020 Then
    RGB1.Red = Int(subt)
    If maxim <= 765 Then
     hue_and_sat = (maxim - 510) * diff1!
     RGB1.Green = Int(value1)
     RGB1.Blue = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 765) * diff1!
     RGB1.Blue = Int(value1)
     RGB1.Green = Int(value1 - hue_and_sat)
    End If
   Else
    RGB1.Green = Int(subt)
    If maxim <= 1275 Then
     hue_and_sat = (maxim - 1020) * diff1!
     RGB1.Blue = Int(value1)
     RGB1.Red = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 1275) * diff1!
     RGB1.Red = Int(value1)
     RGB1.Blue = Int(value1 - hue_and_sat)
    End If
   End If
  Else 'saturation_0_To_1 <= 0
   RGB1.Red = Int(value1)
   RGB1.Green = Int(value1)
   RGB1.Blue = Int(value1)
  End If
 Else 'value_0_To_255 <= 0
  RGB1.Red = 0
  RGB1.Green = 0
  RGB1.Blue = 0
 End If
End Sub
Private Sub RGBQUAD_From_HSVTYPE_P(RGB1 As RGBQUAD, HSV1 As HSVTYPE)
Dim hue_and_sat!

 If HSV1.v_value > 0 Then
  value1 = HSV1.v_value + 0.5
  If HSV1.s_saturation > 0 Then
   maxim = HSV1.h_hue - 1530& * Int(HSV1.h_hue / 1530&)
   diff1 = HSV1.s_saturation * HSV1.v_value
   subt = value1 - diff1
   diff1 = diff1 / 255
   If maxim <= 510 Then
    RGB1.Blue = Int(subt)
    If maxim <= 255 Then
     hue_and_sat = maxim * diff1!
     RGB1.Red = Int(value1)
     RGB1.Green = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 255) * diff1!
     RGB1.Green = Int(value1)
     RGB1.Red = Int(value1 - hue_and_sat)
    End If
   ElseIf maxim <= 1020 Then
    RGB1.Red = Int(subt)
    If maxim <= 765 Then
     hue_and_sat = (maxim - 510) * diff1!
     RGB1.Green = Int(value1)
     RGB1.Blue = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 765) * diff1!
     RGB1.Blue = Int(value1)
     RGB1.Green = Int(value1 - hue_and_sat)
    End If
   Else
    RGB1.Green = Int(subt)
    If maxim <= 1275 Then
     hue_and_sat = (maxim - 1020) * diff1!
     RGB1.Blue = Int(value1)
     RGB1.Red = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 1275) * diff1!
     RGB1.Red = Int(value1)
     RGB1.Blue = Int(value1 - hue_and_sat)
    End If
   End If
  Else 'hsv1.s_saturation <= 0
   RGB1.Red = Int(value1)
   RGB1.Green = Int(value1)
   RGB1.Blue = Int(value1)
  End If
 Else 'hsv1.v_value <= 0
  RGB1.Red = 0
  RGB1.Green = 0
  RGB1.Blue = 0
 End If
End Sub
Public Function RGBHSV(hue_0_To_1530!, ByVal saturation_0_To_1!, value_0_To_255!) As Long
Dim hue_and_sat!

 If value_0_To_255 > 0 Then
  value1 = value_0_To_255 + 0.5
  If saturation_0_To_1 > 0 Then
   maxim = hue_0_To_1530 - 1530& * Int(hue_0_To_1530 / 1530&)
   diff1 = saturation_0_To_1 * value_0_To_255
   subt = value1 - diff1
   diff1 = diff1 / 255
   If maxim <= 510 Then
    BGBlu = Int(subt)
    If maxim <= 255 Then
     hue_and_sat = maxim * diff1!
     BGRed = Int(value1)
     BGGrn = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 255) * diff1!
     BGGrn = Int(value1)
     BGRed = Int(value1 - hue_and_sat)
    End If
   ElseIf maxim <= 1020 Then
    BGRed = Int(subt)
    If maxim <= 765 Then
     hue_and_sat = (maxim - 510) * diff1!
     BGGrn = Int(value1)
     BGBlu = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 765) * diff1!
     BGBlu = Int(value1)
     BGGrn = Int(value1 - hue_and_sat)
    End If
   Else
    BGGrn = Int(subt)
    If maxim <= 1275 Then
     hue_and_sat = (maxim - 1020) * diff1!
     BGBlu = Int(value1)
     BGRed = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 1275) * diff1!
     BGRed = Int(value1)
     BGBlu = Int(value1 - hue_and_sat)
    End If
   End If
   RGBHSV = BGRed Or BGGrn * 256& Or BGBlu * 65536
  Else 'saturation_0_To_1 <= 0
   RGBHSV = Int(value1) * GrayScaleRGB
  End If
 Else 'value_0_To_255 <= 0
  RGBHSV = 0&
 End If
End Function

Public Function ARGBHSV(hue_0_To_1530!, ByVal saturation_0_To_1!, value_0_To_255!) As Long
Dim hue_and_sat!

 If value_0_To_255 > 0 Then
  value1 = value_0_To_255 + 0.5
  If saturation_0_To_1 > 0 Then
   maxim = hue_0_To_1530 - 1530& * Int(hue_0_To_1530 / 1530&)
   diff1 = saturation_0_To_1 * value_0_To_255
   subt = value1 - diff1
   diff1 = diff1 / 255
   If maxim <= 510 Then
    BGBlu = Int(subt)
    If maxim <= 255 Then
     hue_and_sat = maxim * diff1!
     BGRed = Int(value1)
     BGGrn = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 255) * diff1!
     BGGrn = Int(value1)
     BGRed = Int(value1 - hue_and_sat)
    End If
   ElseIf maxim <= 1020 Then
    BGRed = Int(subt)
    If maxim <= 765 Then
     hue_and_sat = (maxim - 510) * diff1!
     BGGrn = Int(value1)
     BGBlu = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 765) * diff1!
     BGBlu = Int(value1)
     BGGrn = Int(value1 - hue_and_sat)
    End If
   Else
    BGGrn = Int(subt)
    If maxim <= 1275 Then
     hue_and_sat = (maxim - 1020) * diff1!
     BGBlu = Int(value1)
     BGRed = Int(subt + hue_and_sat)
    Else
     hue_and_sat = (maxim - 1275) * diff1!
     BGRed = Int(value1)
     BGBlu = Int(value1 - hue_and_sat)
    End If
   End If
   ARGBHSV = BGRed * 65536 Or BGGrn * 256& Or BGBlu
  Else 'saturation_0_To_1 <= 0
   ARGBHSV = Int(value1) * GrayScaleRGB
  End If
 Else 'value_0_To_255 <= 0
  ARGBHSV = 0&
 End If
End Function

