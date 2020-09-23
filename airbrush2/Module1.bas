Attribute VB_Name = "Module1"
Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0

Public Type BITMAPINFOHEADER
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

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type rgb
    red As Byte
    green As Byte
    blue As Byte
End Type

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long



Public Sub drawAirbrush(hdc As Long, X As Long, Y As Long, radius As Long, color As Long, pressure As Long)
    Dim iBitmap As Long, iDC As Long, i As Integer
    Dim bi24BitInfo As BITMAPINFO, bBytes() As Byte, Cnt As Long, xC As Long, yC As Long
    Dim aColor As rgb, tmpRad As String
    
    aColor = getRGB(color)
    
    'make sure the radius is an equal number
    tmpRad = CStr(radius)
    For i = 1 To 9 Step 2
        If Right(tmpRad, 1) = i Then
            radius = radius + 1
            Exit For
        End If
    Next
    
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = CLng(radius * 2)
        .biHeight = CLng(radius * 2)
    End With
    
    ReDim bBytes(1 To (bi24BitInfo.bmiHeader.biWidth + 1) * (bi24BitInfo.bmiHeader.biHeight + 1) * 3) As Byte
    
    iDC = CreateCompatibleDC(0)
    iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    
    SelectObject iDC, iBitmap
    BitBlt iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, hdc, X - radius, Y - radius, vbSrcCopy
    
    
    GetDIBits iDC, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
    
    Cnt = 1
    For yC = -radius To radius - 1
        For xC = -radius To radius - 1
            
            If (xC * xC) + (yC * yC) <= (radius * radius) - 1 Then
                aplha = CByte((255 * ((Sqr((radius * radius)) - Sqr((xC * xC) + (yC * yC))) / radius)) / 100 * pressure)
                
                bBytes(Cnt) = getAlpha(CByte(aplha), CLng(aColor.blue), CLng(bBytes(Cnt)))
                bBytes(Cnt + 1) = getAlpha(CByte(aplha), CLng(aColor.green), CLng(bBytes(Cnt + 1)))
                bBytes(Cnt + 2) = getAlpha(CByte(aplha), CLng(aColor.red), CLng(bBytes(Cnt + 2)))
                
            End If
            Cnt = Cnt + 3
        Next xC
    Next yC

    SetDIBitsToDevice hdc, X - radius, Y - radius, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, 0, 0, 0, bi24BitInfo.bmiHeader.biHeight, bBytes(1), bi24BitInfo, DIB_RGB_COLORS
    
    DeleteDC iDC
    DeleteObject iBitmap
End Sub

Private Function getAlpha(Alpha As Byte, Color1 As Long, color2 As Long)
    getAlpha = color2 + (((Color1 * Alpha) / 255) - ((color2 * Alpha) / 255))
    
End Function

Private Function getRGB(C As Long) As rgb
    Dim RealColor As Long
    
    If C < 0 Then
        TranslateColor C, 0, RealColor
        C = RealColor
    End If
    
    With getRGB
        .red = CByte(C And &HFF&)
        .green = CByte((C And &HFF00&) / 2 ^ 8)
        .blue = CByte((C And &HFF0000) / 2 ^ 16)
    End With
    
End Function
