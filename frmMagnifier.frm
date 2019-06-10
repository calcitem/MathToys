VERSION 5.00
Begin VB.Form frmMagnifier 
   Caption         =   "·Å´ó¾µ"
   ClientHeight    =   1965
   ClientLeft      =   8805
   ClientTop       =   1800
   ClientWidth     =   2250
   Icon            =   "frmMagnifier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   Begin VB.TextBox txtZoom 
      Height          =   285
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "1000%"
      Top             =   0
      Width           =   570
   End
   Begin VB.CheckBox chkOnTop 
      Height          =   270
      Left            =   240
      Picture         =   "frmMagnifier.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Value           =   1  'Checked
      Width           =   240
   End
   Begin VB.CheckBox chkGrid 
      Height          =   270
      Left            =   0
      Picture         =   "frmMagnifier.frx":0680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox picZoom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1635
      Left            =   30
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   103
      TabIndex        =   0
      Top             =   315
      Width           =   1605
      Begin VB.Line l2 
         X1              =   48
         X2              =   48
         Y1              =   56
         Y2              =   36
      End
      Begin VB.Line l1 
         X1              =   36
         X2              =   56
         Y1              =   48
         Y2              =   48
      End
   End
   Begin VB.HScrollBar hsbZoom 
      Height          =   270
      LargeChange     =   10
      Left            =   480
      Max             =   1000
      Min             =   25
      TabIndex        =   1
      Top             =   0
      Value           =   200
      Width           =   1200
   End
   Begin VB.Timer tmrZoom 
      Interval        =   50
      Left            =   1710
      Top             =   360
   End
End
Attribute VB_Name = "frmMagnifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Type PointAPI
    x   As Long
    y   As Long
End Type

Private Type SizeRect
    Left    As Long
    Top     As Long
    Width   As Long
    Height  As Long
End Type

Private Type RectAPI
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Const SRCCOPY           As Long = &HCC0020
Private Const PATCOPY           As Long = &HF00021

Private Const SWP_NOMOVE        As Long = 2
Private Const SWP_NOSIZE        As Long = 1
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_FLAGS         As Long = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
Private Const HWND_TOPMOST      As Long = -1
Private Const HWND_NOTOPMOST    As Long = -2

Private mfScale As Single
Private mlOldX  As Long
Private mlOldY  As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RectAPI) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Function CreateCheckeredBrush(ByVal hdc As Long, ByVal lColor1 As Long, ByVal lColor2 As Long) As Long
On Error Resume Next
Dim x           As Long
Dim y           As Long
Dim lRet        As Long
Dim hBitmapDC   As Long
Dim hBitmap     As Long
Dim hOldBitmap  As Long
    
    If lColor1 < 0 Then
        lColor1 = GetSysColor(lColor1 And &HFF&)
    End If
    If lColor2 < 0 Then
        lColor2 = GetSysColor(lColor2 And &HFF&)
    End If
    
    hBitmapDC = CreateCompatibleDC(hdc)
    hBitmap = CreateCompatibleBitmap(hdc, 8, 8)
    hOldBitmap = SelectObject(hBitmapDC, hBitmap)
    
    For y = 0 To 6 Step 2
        For x = 0 To 6 Step 2
            lRet = SetPixelV(hBitmapDC, x, y, lColor1)
            lRet = SetPixelV(hBitmapDC, x + 1, y, lColor2)
            lRet = SetPixelV(hBitmapDC, x, y + 1, lColor2)
            lRet = SetPixelV(hBitmapDC, x + 1, y + 1, lColor1)
        Next x
    Next y
    
    hBitmap = SelectObject(hBitmapDC, hOldBitmap)
    
    CreateCheckeredBrush = CreatePatternBrush(hBitmap)
    
    lRet = DeleteDC(hBitmapDC)
    lRet = DeleteObject(hBitmap)

End Function

Private Sub DoZoom(ptMouse As PointAPI)
On Error Resume Next
Dim lRet        As Long
Dim lTemp       As Long
Dim hWndDesk    As Long
Dim hDCDesk     As Long
Dim sizSrce     As SizeRect
Dim sizDest     As SizeRect

    hWndDesk = GetDesktopWindow()
    hDCDesk = GetDC(hWndDesk)
    
    With sizDest
        .Left = 0
        .Top = 0
        .Width = picZoom.ScaleWidth
        .Height = picZoom.ScaleHeight
    End With
    
    With sizSrce
        .Left = ptMouse.x - Int((sizDest.Width / 2) / mfScale)
        .Top = ptMouse.y - Int((sizDest.Height / 2) / mfScale)
        .Width = Int(sizDest.Width / mfScale)
        .Height = Int(sizDest.Height / mfScale)
        lTemp = Int(.Width * mfScale)
        If lTemp > sizDest.Width Then
            sizDest.Width = lTemp
        ElseIf lTemp < sizDest.Width Then
            .Width = .Width + 1
            sizDest.Width = lTemp + mfScale
        End If
        lTemp = Int(.Height * mfScale)
        If lTemp > sizDest.Height Then
            sizDest.Height = lTemp
        ElseIf lTemp < sizDest.Height Then
            .Height = .Height + 1
            sizDest.Height = lTemp + mfScale
        End If
    End With
    
    picZoom.Cls
    
    lRet = StretchBlt(picZoom.hdc, sizDest.Left, sizDest.Top, sizDest.Width, sizDest.Height, hDCDesk, sizSrce.Left, sizSrce.Top, sizSrce.Width, sizSrce.Height, SRCCOPY)
    
    lRet = ReleaseDC(hWndDesk, hDCDesk)
    
    If chkGrid.Value = vbChecked Then
        Call DrawGrid
    End If
    
    picZoom.Refresh
    
End Sub

Private Sub DrawGrid()
On Error Resume Next
Dim iWidth      As Integer
Dim iHeight     As Integer
Dim lRet        As Long
Dim hBrush      As Long
Dim hOldBrush   As Long
Dim fX          As Single
Dim fY          As Single

    If mfScale >= 3 Then
    
        hBrush = CreateCheckeredBrush(picZoom.hdc, &H808080, &HC0C0C0)
        hOldBrush = SelectObject(picZoom.hdc, hBrush)
        
        iWidth = picZoom.ScaleWidth
        iHeight = picZoom.ScaleHeight
        
        For fX = 0 To iWidth Step mfScale
            lRet = PatBlt(picZoom.hdc, Int(fX), 0, 1, iHeight, PATCOPY)
        Next
        For fY = 0 To iHeight Step mfScale
            lRet = PatBlt(picZoom.hdc, 0, Int(fY), iWidth, 1, PATCOPY)
        Next
        
        hBrush = SelectObject(picZoom.hdc, hOldBrush)
        lRet = DeleteObject(hBrush)
    
    End If
    
End Sub
Private Function ValidScale(ByVal fScale As Single) As Single
On Error Resume Next
    If fScale * 100 > hsbZoom.Max Then
        fScale = hsbZoom.Max / 100
    ElseIf fScale * 100 < hsbZoom.min Then
        fScale = hsbZoom.min / 100
    End If
    
    ValidScale = fScale
    
End Function

Private Sub LoadSettings()
On Error Resume Next
   ' Call RestoreFormSize(Me)
    'hsbZoom.Value = GetInitEntry("Settings", "Zoom", CStr(200))
    hsbZoom_Change
   ' chkGrid.Value = IIf(LCase$(GetInitEntry("Settings", "Grid", "False")) = "true", vbChecked, vbUnchecked)
    chkGrid_Click
   ' chkOnTop.Value = IIf(LCase$(GetInitEntry("Settings", "OnTop", "False")) = "true", vbChecked, vbUnchecked)
   chkOnTop_Click

End Sub

'Private Sub SaveSettings()


'Dim lRet As Long

   ' Call SaveFormSize(Me)
   ' lRet = SetInitEntry("Settings", "Zoom", hsbZoom.Value)
   ' lRet = SetInitEntry("Settings", "Grid", CStr(chkGrid.Value = vbChecked))
   ' lRet = SetInitEntry("Settings", "OnTop", CStr(chkOnTop.Value = vbChecked))

'End Sub

Private Sub chkGrid_Click()
On Error Resume Next
    mlOldX = -100
    
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
    
End Sub

Private Sub chkOnTop_Click()
On Error Resume Next
Dim lRet    As Long
Dim lWinPos As Long

    lWinPos = IIf(chkOnTop.Value = vbChecked, HWND_TOPMOST, HWND_NOTOPMOST)
    lRet = SetWindowPos(Me.hwnd, lWinPos, 0, 0, 0, 0, SWP_FLAGS)
    
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
    
End Sub

Private Sub Form_Load()
    'frmMain.Skin Me
    Call LoadSettings

End Sub


Private Sub Form_Resize()
On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        If Me.Width < 1680 Then
            Me.Width = 1680
        ElseIf Me.Height < 1680 Then
            Me.Height = 1680
        Else
            chkGrid.Move 0, 0
            chkOnTop.Move chkGrid.Width, 0
            hsbZoom.Move chkGrid.Width + chkOnTop.Width, 0, Me.ScaleWidth - txtZoom.Width - chkGrid.Width - chkOnTop.Width
            txtZoom.Move Me.ScaleWidth - txtZoom.Width, -1
            picZoom.Move 0, hsbZoom.Height, Me.ScaleWidth, Me.ScaleHeight - hsbZoom.Height
            l1.X1 = picZoom.Width / 2 - 10
            l1.X2 = picZoom.Width / 2 + 10
            l1.Y1 = picZoom.Height / 2
            l1.Y2 = l1.Y1
            l2.Y1 = picZoom.Height / 2 - 10
            l2.Y2 = picZoom.Height / 2 + 10
            l2.X1 = picZoom.Width / 2
            l2.X2 = l2.X1
        End If
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

   ' Call SaveSettings
    
End Sub


Private Sub hsbZoom_Change()
On Error Resume Next
    txtZoom.Text = Format$(hsbZoom.Value / 100, "####%")
    
    mfScale = CSng(hsbZoom.Value) / 100!
    
    If picZoom.Visible Then
        picZoom.SetFocus
    End If
    
    mlOldX = -100

End Sub


Private Sub hsbZoom_Scroll()
On Error Resume Next
    hsbZoom_Change
    
End Sub


Private Sub tmrZoom_Timer()
On Error Resume Next
Dim lRet    As Long
Dim ptMouse As PointAPI

Static lElapsed As Long

    If Me.WindowState <> vbMinimized Then
        lElapsed = lElapsed + tmrZoom.Interval
        lRet = GetCursorPos(ptMouse)
        With ptMouse
            If (.x <> mlOldX) Or (.y <> mlOldY) Or (lElapsed >= 250) Then
                Call DoZoom(ptMouse)
                If lElapsed >= 250 Then
                    If chkOnTop.Value = vbChecked Then
                        lRet = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
                    End If
                End If
                lElapsed = 0
            End If
            mlOldX = .x
            mlOldY = .y
        End With
    End If
End Sub


Private Sub txtZoom_GotFocus()
On Error Resume Next
    With txtZoom
        .Text = CStr(Val(.Text))
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    
End Sub


Private Sub txtZoom_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii > 31 And (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then
        Beep
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        picZoom.SetFocus
        DoEvents
        txtZoom.SetFocus
        KeyAscii = 0
    End If
    
End Sub


Private Sub txtZoom_LostFocus()
On Error Resume Next
    mfScale = ValidScale(Val(txtZoom.Text) / 100)
    
    hsbZoom.Value = mfScale * 100
    
    txtZoom.Text = Format$(mfScale, "####%")

End Sub


