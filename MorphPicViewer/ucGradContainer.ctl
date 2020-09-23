VERSION 5.00
Begin VB.UserControl MorphContainer 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
   ToolboxBitmap   =   "ucGradContainer.ctx":0000
End
Attribute VB_Name = "MorphContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*************************************************************************
'* MorphContainer - Owner-drawn gradient container control.              *
'* Matthew R. Usner et. al., April, 2005.                                *
'*************************************************************************
'* A completely owner-drawn replacement for VB's dull frame control.     *
'* Features include:                                                     *
'* - Separate gradients for header and container.                        *
'* - Container and header gradients can be drawn at any angle.           *
'* - Container background can be a gradient or bitmap.                   *
'* - Background bitmap can be tiled or stretched to fit control.         *
'* - Icon display capability.                                            *
'* - Unicode character display supported.                                *
'* - Ability to round each corner to user-specified curvature amounts.   *
'* - 12 different XP-style color themes are incorporated.                *
'* - Container can be collapsed and expanded by double-clicking header.  *
'* - Icon can be displayed in the left or right of the header.           *
'*************************************************************************
'* Legal:  Redistribution of this code, whole or in part, as source code *
'* or in binary form, alone or as part of a larger distribution or prod- *
'* uct, is forbidden for any commercial or for-profit use without the    *
'* author's explicit written permission.                                 *
'*                                                                       *
'* Redistribution of this code, as source code or in binary form, with   *
'* or without modification, is permitted provided that the following     *
'* conditions are met:                                                   *
'*                                                                       *
'* Redistributions of source code must include this list of conditions,  *
'* and the following acknowledgment:                                     *
'*                                                                       *
'* This code was developed by Matthew R. Usner.                          *
'* Source code, written in Visual Basic, is freely available for non-    *
'* commercial, non-profit use.                                           *
'*                                                                       *
'* Redistributions in binary form, as part of a larger project, must     *
'* include the above acknowledgment in the end-user documentation.       *
'* Alternatively, the above acknowledgment may appear in the software    *
'* itself, if and where such third-party acknowledgments normally appear.*
'*************************************************************************
'* Credits and Thanks:                                                   *
'* Originally based on 'XP Container Control' written by Cameron Groves. *
'* Carles P.V. - gradient code and the code for the ability to round     *
'* individual corners.                                                   *
'* Richard Mewett - Unicode support.                                     *
'* Franck Nunes - Expand/Collapse idea and implementation.               *
'*************************************************************************

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpDest As Any, ByRef lpSource As Any, ByVal iLen As Long)
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long

'  declares for Unicode support.
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
   dwOSVersionInfoSize                As Long
   dwMajorVersion                     As Long
   dwMinorVersion                     As Long
   dwBuildNumber                      As Long
   dwPlatformId                       As Long
   szCSDVersion                       As String * 128     '  Maintenance string for PSS usage
End Type
Private mWindowsNT                    As Boolean
Private Const DT_CALCRECT             As Long = &H400     ' if used, DrawText API just calculates rectangle.
Private Const DT_SINGLELINE           As Long = &H20      ' strip cr/lf from string before draw.
Private Const DT_NOPREFIX             As Long = &H800     ' ignore access key ampersand.
Private Const DT_LEFT                 As Long = &H0       ' draw from left edge of rectangle.
Private Const DT_NOCLIP               As Long = &H100     ' ignores right edge of rectangle when drawing.

' declares for Carles P.V.'s gradient paint routine.
Private Type BITMAPINFOHEADER
   biSize          As Long
   biWidth         As Long
   biHeight        As Long
   biPlanes        As Integer
   biBitCount      As Integer
   biCompression   As Long
   biSizeImage     As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed       As Long
   biClrImportant  As Long
End Type

' used to define the text drawing area.
Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

'  enum tied to .PictureMode property.
Public Enum MC_PictureModeOptions
   [Normal] = 0
   [Stretch] = 1
   [Tiled] = 2
End Enum

Public Enum IconSizeEnum
   [Display Full Size] = 0
   [Size To Header] = 1
End Enum

Public Enum MC_Themes
   [Blue] = 0
   [Dark Blue] = 1
   [Dark Green] = 2
   [Green] = 3
   [Light Blue] = 4
   [Light Green] = 5
   [Orange] = 6
   [Pastel Green] = 7
   [Purple] = 8
   [Red] = 9
   [Silver] = 10
   [Yellow] = 11
End Enum

Public Enum MC_CaptionAlignment
   [Left Justify] = 0
   [Right Justify] = 1
   [Center] = 2
End Enum

Public Enum AutoResizeEvent
   [Double Left Button Click] = 0
   [Double Right Button Click] = 1
End Enum

Public Enum IconAlignmentOptions
   [Align Left] = 0
   [Align Right] = 1
End Enum

Private Type BITMAP
   bmType       As Long
   bmWidth      As Long
   bmHeight     As Long
   bmWidthBytes As Long
   bmPlanes     As Integer
   bmBitsPixel  As Integer
   bmBits       As Long
End Type

Private Type POINTAPI
   X As Long
   Y As Long
End Type

'  property variables and constants.
Private m_PictureMode       As MC_PictureModeOptions ' normal, tiled or stretched background bitmap.
Private m_Picture           As Picture               ' container background bitmap.
Private m_IconAlignment     As IconAlignmentOptions  ' icon can be displayed in left or right of header.
Private m_Movable           As Boolean               ' allows container to be dragged using header.
Private m_CurveTopLeft      As Long                  ' the curvature of the top left corner.
Private m_CurveTopRight     As Long                  ' the curvature of the top right corner.
Private m_CurveBottomLeft   As Long                  ' the curvature of the bottom left corner.
Private m_CurveBottomRight  As Long                  ' the curvature of the bottom right corner.
Private m_HeaderVisible     As Boolean               ' flag that shows/hides header.
Private m_BackMiddleOut     As Boolean               ' flag for container background middle-out gradient.
Private m_HeaderMiddleOut   As Boolean               ' flag for header middle-out gradient.
Private m_Enabled           As Boolean               ' enabled/disabled flag.
Private m_HeaderAngle       As Single                ' the angle of the header gradient.
Private m_BackAngle         As Single                ' background gradient display angle
Private m_Iconsize          As IconSizeEnum          ' icon size - full or size to header
Private m_HeaderColor1      As OLE_COLOR             ' the first gradient color of the header.
Private m_HeaderColor2      As OLE_COLOR             ' the second gradient color of the header.
Private m_BackColor1        As OLE_COLOR             ' the first gradient color of the background.
Private m_BackColor2        As OLE_COLOR             ' the second gradient color of the background.
Private m_BorderWidth       As Integer               ' width, in pixels, of border.
Private m_BorderColor       As OLE_COLOR             ' color of border.
Private m_CaptionColor      As OLE_COLOR             ' text color of caption.
Private m_Caption           As String                ' caption text.
Private m_HeaderHeight      As Long                  ' height, in pixels, of the header.
Private m_CaptionFont       As StdFont               ' font used to display header text.
Private m_Alignment         As MC_CaptionAlignment   ' caption alignment (left, center, right).
Private m_Icon              As Picture               ' the icon or bitmap to display in the header.
Private m_Theme             As MC_Themes             ' XP-style color schemes.
Private m_Expanded          As Boolean               ' informs user when container is full size or collapsed.
Private m_AutoResize        As Boolean               ' container collapses/expands without any code from user.
Private m_AutoResizeEvent   As AutoResizeEvent       ' container autocolapses on single click or double click.

Private Const m_def_PictureMode = 0                 ' normal container background bitmap display.
Private Const m_def_IconAlignment = 0               ' initialize icon to left alignment.
Private Const m_def_Movable = False                 ' initialize container to fixed position.
Private Const m_def_CurveTopLeft = 0                ' initialize top left curvature to 0.
Private Const m_def_CurveTopRight = 0               ' initialize top right curvature to 0.
Private Const m_def_CurveBottomLeft = 0             ' initialize bottom left curvature to 0.
Private Const m_def_CurveBottomRight = 0            ' initialize bottom right curvature to 0.
Private Const m_def_HeaderVisible = True            ' initialize the header to be visible.
Private Const m_def_BackMiddleOut = True            ' initialize to a middle-out background gradient.
Private Const m_def_HeaderMiddleOut = True          ' initialize to a middle-out header gradient.
Private Const m_def_Enabled = 0                     ' initialize to disabled.
Private Const m_def_HeaderAngle = 90                ' initialize to horizontal header gradient.
Private Const m_def_BackAngle = 90                  ' initialize to horizontal background gradient.
Private Const m_def_Iconsize = 1                    ' initialize to 'size to header'
Private Const m_def_HeaderColor2 = &HF7E0D3
Private Const m_def_HeaderColor1 = &HEDC5A7
Private Const m_def_BackColor2 = &HFCF4EF
Private Const m_def_BackColor1 = &HFAE8DC
Private Const m_def_Caption = "MorphContainer"      ' default caption text.
Private Const m_def_BorderWidth = 1                 ' initialize border width to 1 pixel.
Private Const m_def_BorderColor = &HDCC1AD
Private Const m_def_Alignment = 0                   ' initalize text to left justification.
Private Const m_def_CaptionColor = &H7B2D02
Private Const m_def_hHeight = 25                    ' initialize header to 25 pixels in height.
Private Const m_def_Theme = 0
Private Const m_def_Expanded = True
Private Const m_def_AutoResize = True
Private Const m_def_AutoResizeEvent = 1

'  events.
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Resize()

'  miscellaneous control variables and constants.
Private Const OBJ_BITMAP     As Long = 7                  ' used to determine if picture is a bitmap.
Private Const RGN_DIFF       As Long = 4
Private Const DIB_RGB_COLORS As Long = 0
Private Const PI             As Single = 3.14159265358979
Private Const TO_DEG         As Single = 180 / PI
Private Const TO_RAD         As Single = PI / 180
Private Const INT_ROT        As Long = 1000
Private m_hMod               As Long
Private PreviousHeight       As Long    ' for collapsing/expanding container, container original height.
Private MousePosY            As Single  ' stores y coordinate of mouse (for collapse & move).
Private MouseButton          As Integer ' stores last mouse button clicked.
Private MouseButtonDown      As Boolean ' for dragging of container.
Private m_hBrush As Long                        ' pattern brush for bitmap tiling.

Private X               As Long, Y As Long, h1 As Long, h2 As Long, h3 As Long
Private wid             As Long, hgt As Long

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Events >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_DblClick()

'*************************************************************************
'* handles expansion or collapse of container on double-click.           *
'*************************************************************************

   If m_HeaderVisible = False Then
'     no header, don't allow autoresizing.
      RaiseEvent DblClick
      Exit Sub
   End If

   If m_AutoResize And MousePosY < m_HeaderHeight Then
      If (m_AutoResizeEvent = [Double Right Button Click] And MouseButton = vbRightButton) Or _
         (m_AutoResizeEvent = [Double Left Button Click] And MouseButton = vbLeftButton) Then
         If m_Expanded = True Then
            CollapseContainer
         Else
            ExpandContainer
         End If
      End If
   End If

   RaiseEvent DblClick

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   MouseButtonDown = False
   RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
   m_hMod = LoadLibrary("shell32.dll") ' Used to prevent crashes on Windows XP
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   MouseButtonDown = True
   MouseButton = Button 'Capture clicked button.
   RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* saves mouse Y position for collapse/expand, and also allows the cont- *
'* ainer to be dragged if the appropriate conditions are met.            *
'*************************************************************************

   Dim Ret As Long

'  capture mouse vertical position for possible use in collapse/expand.
   MousePosY = Y

'  allow the container to be dragged if all conditions are met.
   If MouseButtonDown And HeaderVisible And MousePosY < m_HeaderHeight And m_Movable Then
      If Button = vbLeftButton Then
         ReleaseCapture
         Ret = SendMessage(UserControl.hWnd, &H112, &HF012, 0)
      End If
   End If

   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Show()
   RedrawControl
End Sub

Private Sub UserControl_Resize()

'*************************************************************************
'* controls redrawing of container when resized.                         *
'*************************************************************************

   Dim H As Single

   On Error GoTo ErrHandler

   H = m_HeaderHeight * Screen.TwipsPerPixelY
   If UserControl.Height < H Then
      UserControl.Height = H
   End If
   RedrawControl
   RaiseEvent Resize

ErrHandler:

End Sub

Private Sub UserControl_Terminate()
   
'*************************************************************************
'* last event in lifecycle of a control.                                 *
'*************************************************************************

   FreeLibrary m_hMod ' Used to prevent crashes on Windows XP

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Graphics >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub RedrawControl()

'*************************************************************************
'* master routine for painting of container.                             *
'*************************************************************************

   UserControl.Cls

   If m_Expanded Then
      SetBackGround
   End If
   If m_HeaderVisible Then
      SetHeader
   End If
   CreateBorder

End Sub

Private Sub ApplyTheme()

'*************************************************************************
'* XP-style color schemes. By Cameron Groves.                            *
'*************************************************************************

    Select Case m_Theme

       Case [Blue]
          HeaderColor2 = &HF7E0D3
          HeaderColor1 = &HEDC5A7
          BackColor2 = &HFCF4EF
          BackColor1 = &HFAE8DC
          BorderColor = &HDCC1AD
          CaptionColor = &H7B2D02

       Case [Dark Blue]
          HeaderColor2 = &HECDCD3
          HeaderColor1 = &HDABAA8
          BackColor2 = &HF8F2EF
          BackColor1 = &HF1E5DD
          BorderColor = &HD6B4A0
          CaptionColor = &H4B2A17

       Case [Dark Green]
          HeaderColor2 = &HD8E5C8
          HeaderColor1 = &HB1CB92
          BackColor2 = &HF1F5EB
          BackColor1 = &HE1EBD5
          BorderColor = &HAAC688
          CaptionColor = &H213B00

       Case [Green]
          HeaderColor2 = &HE0EAE8
          HeaderColor1 = &HC2D6D1
          BackColor2 = &HF4F8F7
          BackColor1 = &HE7EFED
          BorderColor = &HBCD3CD
          CaptionColor = &H324741

       Case [Light Blue]
          HeaderColor2 = &HF1E3C8
          HeaderColor1 = &HE4C992
          BackColor2 = &HFAF5EB
          BackColor1 = &HF5EAD5
          BorderColor = &HE2C488
          CaptionColor = &H553900

       Case [Light Green]
          HeaderColor2 = &HDAF2E3
          HeaderColor1 = &HB5E5C8
          BackColor2 = &HF1FAF5
          BackColor1 = &HE3F5EA
          BorderColor = &HAEE3C3
          CaptionColor = &H245738

       Case [Orange]
          HeaderColor2 = &HD2E2FD
          HeaderColor1 = &HA7C6FA
          BackColor2 = &HEFF5FE
          BackColor1 = &HDDE9FD
          BorderColor = &H9FC0FA
          CaptionColor = &H16366D

       Case [Pastel Green]
          HeaderColor2 = &HE3E3D6
          HeaderColor1 = &HC9C9AE
          BackColor2 = &HF5F5F0
          BackColor1 = &HEAEAE0
          BorderColor = &HC4C4A6
          CaptionColor = &H39391D

       Case [Purple]
          HeaderColor2 = &HEAD7DF
          HeaderColor1 = &HD5B0BF
          BackColor2 = &HF7F1F3
          BackColor1 = &HEFE1E6
          BorderColor = &HD1A9B9
          CaptionColor = &H46202F

       Case [Red]
          HeaderColor2 = &HD6D2FB
          HeaderColor1 = &HAEA6F8
          BackColor2 = &HF0EFFE
          BackColor1 = &HE0DDFC
          BorderColor = &HA79EF7
          CaptionColor = &H1D156A

       Case [Silver]
          HeaderColor2 = &HECEAE9
          HeaderColor1 = &HD9D6D3
          BackColor2 = &HF8F7F7
          BackColor1 = &HF1EFEE
          BorderColor = &HD6D2CF
          CaptionColor = &H4A4744

       Case [Yellow]
          HeaderColor2 = &HE4FAFC
          HeaderColor1 = &HB9EEF4
          BackColor2 = &HEEFCFD
          BackColor1 = &HDCF7FA
          BorderColor = &H95E1EA
          CaptionColor = &H66D5E1

    End Select

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* converts color long COLORREF for api coloring purposes.               *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If

End Function

Private Sub SetBackGround()

'*************************************************************************
'* displays the control's background gradient or displays a bitmap if    *
'* one is specified.                                                     *
'*************************************************************************

   If IsPictureThere(m_Picture) Then
      Select Case m_PictureMode
         Case [Normal]
            Set UserControl.Picture = m_Picture
         Case [Tiled]
'           tile the background bitmap in the control, accounting for possible border and header.
            SetPattern m_Picture
            If m_HeaderVisible Then
               Tile hdc, m_BorderWidth, m_BorderWidth + m_HeaderHeight - 2, ScaleWidth - m_BorderWidth, ScaleHeight - m_BorderWidth
            Else
               Tile hdc, m_BorderWidth, m_BorderWidth, ScaleWidth - m_BorderWidth, ScaleHeight - m_BorderWidth
            End If
         Case [Stretch]
            StretchPicture
      End Select
   Else
      If m_HeaderVisible Then
'        if there is a visible header, the top of the background gradient is the row of pixels
'        under the header.  Otherwise, it is the top of the control, plus the border width.
         PaintGradient hdc, 0, m_HeaderHeight, ScaleWidth, ScaleHeight - m_HeaderHeight - m_BorderWidth, _
                       TranslateColor(m_BackColor1), TranslateColor(m_BackColor2), m_BackAngle, m_BackMiddleOut
      Else
         PaintGradient hdc, 0, 0, ScaleWidth, ScaleHeight, _
                       TranslateColor(m_BackColor1), TranslateColor(m_BackColor2), m_BackAngle, m_BackMiddleOut
      End If
   End If

End Sub

Private Sub StretchPicture()
   
'*************************************************************************
'* stretch bitmap to fit listbox background.  Thanks to LaVolpe for the  *
'* suggestion and AllAPI.net / VBCity.com for the learning to do it.     *
'*************************************************************************

   Dim TempBitmap As BITMAP       ' bitmap structure that temporarily holds picture.
   Dim CreateDC As Long           ' used in creating temporary bitmap structure virtual DC.
   Dim TempBitmapDC As Long       ' virtual DC of temporary bitmap structure.
   Dim TempBitmapOld As Long      ' used in destroying temporary bitmap structure virtual DC.
   Dim r As Long                  ' result long for StretchBlt call.

'  create a temporary bitmap and DC to place the picture in.
   GetObjectAPI m_Picture.Handle, Len(TempBitmap), TempBitmap
   CreateDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   TempBitmapDC = CreateCompatibleDC(CreateDC)
   TempBitmapOld = SelectObject(TempBitmapDC, m_Picture.Handle)

'  stretch it according to existence of border and/or header.
   If m_HeaderVisible Then
      r = StretchBlt(hdc, m_BorderWidth, m_BorderWidth + m_HeaderHeight - 1, ScaleWidth - m_BorderWidth, ScaleHeight - m_BorderWidth - m_HeaderHeight, TempBitmapDC, 0, 0, TempBitmap.bmWidth, TempBitmap.bmHeight, vbSrcCopy)
   Else
      r = StretchBlt(hdc, m_BorderWidth, m_BorderWidth, ScaleWidth - m_BorderWidth, ScaleHeight - m_BorderWidth, TempBitmapDC, 0, 0, TempBitmap.bmWidth, TempBitmap.bmHeight, vbSrcCopy)
   End If

'  destroy temporary bitmap DC.
   SelectObject TempBitmapDC, TempBitmapOld
   DeleteDC TempBitmapDC
   DeleteDC CreateDC

End Sub

Private Function IsPictureThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks for existence of a picture.  Thanks to Roger Gilchrist.        *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then
         IsPictureThere = Pic.Width <> 0
      End If
   End If

End Function

Private Sub CreateBorder()

'*************************************************************************
'* draws the border around the control, using appropriate curvatures.    *
'*************************************************************************

   Dim hRgn1   As Long   ' the outer region of the border.
   Dim hRgn2   As Long   ' the inner region of the border.
   Dim hBrush  As Long   ' the solid-color brush used to paint the combined border regions.

'  create the outer region.
   hRgn1 = pvGetRoundedRgn(0, 0, ScaleWidth, ScaleHeight, _
                           m_CurveTopLeft, m_CurveTopRight, _
                           m_CurveBottomLeft, m_CurveBottomRight)

'  create the inner region.
   hRgn2 = pvGetRoundedRgn(m_BorderWidth, m_BorderWidth, _
                           ScaleWidth - m_BorderWidth, ScaleHeight - m_BorderWidth, _
                           m_CurveTopLeft, m_CurveTopRight, _
                           m_CurveBottomLeft, m_CurveBottomRight)

'  combine the outer and inner regions.
   CombineRgn hRgn2, hRgn1, hRgn2, RGN_DIFF
'  create the brush used to color the combined regions.
   hBrush = CreateSolidBrush(TranslateColor(m_BorderColor))
'  color the combined regions.
   FillRgn hdc, hRgn2, hBrush

'  set the container's visibility region.
   SetWindowRgn hWnd, hRgn1, True

'  delete created objects to restore memory.
   DeleteObject hBrush
   DeleteObject hRgn1
   DeleteObject hRgn2

End Sub

Private Function pvGetRoundedRgn(ByVal x1 As Long, ByVal y1 As Long, _
                                 ByVal x2 As Long, ByVal y2 As Long, _
                                 ByVal TopLeftRadius As Long, ByVal TopRightRadius As Long, _
                                 ByVal BottomLeftRadius As Long, ByVal BottomRightRadius As Long) As Long

'*************************************************************************
'* allows each corner of the container to have its own curvature.        *
'* Code by Carles P.V.                                                   *
'*************************************************************************

   Dim hRgnMain As Long   ' the original "starting point" region.
   Dim hRgnTmp1 As Long   ' the first region that defines a corner's radius.
   Dim hRgnTmp2 As Long   ' the second region that defines a corner's radius.

'  bounding region.
   hRgnMain = CreateRectRgn(x1, y1, x2, y2)

'  top-left corner.
   hRgnTmp1 = CreateRectRgn(x1, y1, x1 + TopLeftRadius, y1 + TopLeftRadius)
   hRgnTmp2 = CreateEllipticRgn(x1, y1, x1 + 2 * TopLeftRadius, y1 + 2 * TopLeftRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

'  top-right corner.
   hRgnTmp1 = CreateRectRgn(x2, y1, x2 - TopRightRadius, y1 + TopRightRadius)
   hRgnTmp2 = CreateEllipticRgn(x2 + 1, y1, x2 + 1 - 2 * TopRightRadius, y1 + 2 * TopRightRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

'  bottom-left corner.
   hRgnTmp1 = CreateRectRgn(x1, y2, x1 + BottomLeftRadius, y2 - BottomLeftRadius)
   hRgnTmp2 = CreateEllipticRgn(x1, y2 + 1, x1 + 2 * BottomLeftRadius, y2 + 1 - 2 * BottomLeftRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

'  bottom-right corner.
   hRgnTmp1 = CreateRectRgn(x2, y2, x2 - BottomRightRadius, y2 - BottomRightRadius)
   hRgnTmp2 = CreateEllipticRgn(x2 + 1, y2 + 1, x2 + 1 - 2 * BottomRightRadius, y2 + 1 - 2 * BottomRightRadius)
   CombineRegions hRgnTmp1, hRgnTmp2, hRgnMain

   pvGetRoundedRgn = hRgnMain

End Function

Private Sub CombineRegions(ByVal Region1 As Long, ByVal Region2 As Long, ByVal MainRegion As Long)

'*************************************************************************
'* combines outer/inner rectangular regions for border painting.         *
'*************************************************************************

   CombineRgn Region1, Region1, Region2, RGN_DIFF
   CombineRgn MainRegion, MainRegion, Region1, RGN_DIFF
   DeleteObject Region1
   DeleteObject Region2

End Sub

'******************* bitmap tiling routines by Carles P.V.
' adapted from Carles' class titled "DIB Brush - Easy Image Tiling Using FillRect"
' at Planet Source Code, txtCodeId=40585.

Private Function SetPattern(Picture As StdPicture) As Boolean

'*************************************************************************
'* creates the brush pattern for tiling into the listbox.  By Carles P.V.*
'*************************************************************************

   Dim tBI       As BITMAP
   Dim tBIH      As BITMAPINFOHEADER
   Dim Buff()    As Byte 'Packed DIB

   Dim lhDC      As Long
   Dim lhOldBmp  As Long

   If (GetObjectType(Picture) = OBJ_BITMAP) Then

'     -- Get image info
      GetObject Picture, Len(tBI), tBI

'     -- Prepare DIB header and redim. Buff() array
      With tBIH
         .biSize = Len(tBIH) '40
         .biPlanes = 1
         .biBitCount = 24
         .biWidth = tBI.bmWidth
         .biHeight = tBI.bmHeight
         .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
      End With
      ReDim Buff(1 To Len(tBIH) + tBIH.biSizeImage) '[Header + Bits]

'     -- Create DIB brush
      lhDC = CreateCompatibleDC(0)
      If (lhDC <> 0) Then
         lhOldBmp = SelectObject(lhDC, Picture)

'        -- Build packed DIB:
'        - Merge Header
         CopyMemory Buff(1), tBIH, Len(tBIH)
'        - Get and merge DIB Bits
         GetDIBits lhDC, Picture, 0, tBI.bmHeight, Buff(Len(tBIH) + 1), tBIH, DIB_RGB_COLORS

         SelectObject lhDC, lhOldBmp
         DeleteDC lhDC

'        -- Create brush from packed DIB
         DestroyPattern
         m_hBrush = CreateDIBPatternBrushPt(Buff(1), DIB_RGB_COLORS)
      End If

   End If

   SetPattern = (m_hBrush <> 0)

End Function

Private Sub Tile(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

'*************************************************************************
'* performs the tiling of the bitmap on the control.  By Carles P.V.     *
'*************************************************************************

   Dim TileRect As RECT
   Dim PtOrg    As POINTAPI

   If (m_hBrush <> 0) Then
      SetRect TileRect, x1, y1, x2, y2
      SetBrushOrgEx hdc, x1, y1, PtOrg
'     -- Tile image
      FillRect hdc, TileRect, m_hBrush
   End If

End Sub

Private Sub DestroyPattern()
   
'*************************************************************************
'* destroys the pattern brush used to tile the bitmap.  By Carles P.V.   *
'*************************************************************************
   
   If (m_hBrush <> 0) Then
      DeleteObject m_hBrush
      m_hBrush = 0
   End If

End Sub

'******************* end of bitmap tiling routines by Carles P.V.

Private Sub SetHeader()

'*************************************************************************
'* displays the header gradient, header caption, and an icon if used.    *
'*************************************************************************

   Dim Clearance As Long

   If Not m_CaptionFont Is Nothing Then

'     fill header gradient.
      PaintGradient hdc, 0, 0, _
                    ScaleWidth, m_HeaderHeight, _
                    TranslateColor(m_HeaderColor1), TranslateColor(m_HeaderColor2), _
                    m_HeaderAngle, m_HeaderMiddleOut

'     obtain the width of one letter to use as a left/right caption display clearance.
      Clearance = TextWidthU(hdc, "A")

'     draw the caption.
      Dim TextRect As RECT  ' will define the text drawing region.

'     apply the font and text color.
      Set UserControl.Font = m_CaptionFont
      UserControl.ForeColor = TranslateColor(m_CaptionColor)

      With TextRect
'        define the text drawing area rectangle.
         If m_Alignment = vbCenter Then
            .Left = (ScaleWidth - TextWidthU(hdc, m_Caption)) / 2
         ElseIf m_Alignment = vbLeftJustify Then
            If IsThere(m_Icon) Then
'              provide for a left-hand clearance of one character plus height of header.
               .Left = Clearance + m_HeaderHeight
            Else
'              provide for a left-hand clearance of on character width.
               .Left = Clearance
            End If
         Else
'           provide a right-hand clearance of one character width.
            .Left = (ScaleWidth - TextWidthU(hdc, m_Caption)) - Clearance
         End If
'        define the rest of the text drawing rectangle.
         .Top = (m_HeaderHeight - TextHeight(m_Caption)) / 2
         .Bottom = .Top + TextHeight(m_Caption)
         .Right = .Left + TextWidthU(hdc, m_Caption)
      End With

'     draw the caption.
      DrawText hdc, m_Caption, -1, TextRect, 0

'     draw the icon, if one has been specified.
      If IsThere(m_Icon) Then
'        if specified, display the icon in its original size.
         If m_Iconsize = [Display Full Size] Then
            PaintPicture m_Icon, IconX, 2
         Else
'           otherwise, fit it into the confines defined by the header's height.
            PaintPicture m_Icon, IconX, 2, m_HeaderHeight - 2, m_HeaderHeight - 3
         End If
      End If

   End If

End Sub

Private Function TextWidthU(ByVal hdc As Long, sString As String) As Long

'*************************************************************************
'* a better alternative to the method .TextWidth.  Thanks to LaVolpe.    *
'*************************************************************************

   Dim Flags    As Long    ' the DT_CALCRECT flag calculates the width without displaying the text.
   Dim TextRect As RECT    ' a rectangle that will have the exact width of the text.

   SetRect TextRect, 0, 0, 0, 0
   Flags = DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_LEFT
   DrawText hdc, sString, -1, TextRect, Flags
   TextWidthU = TextRect.Right + 1

End Function

Private Sub DrawText(ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long)

'*************************************************************************
'* draws the text with Unicode support based on OS version.              *
'* Thanks to Richard Mewett.                                             *
'*************************************************************************

   If mWindowsNT Then
      DrawTextW hdc, StrPtr(lpString), nCount, lpRect, wFormat
   Else
      DrawTextA hdc, lpString, nCount, lpRect, wFormat
   End If

End Sub

Private Function IconX() As Long

'*************************************************************************
'* returns the X coordinate of icon based on IconAlignment property.     *
'*************************************************************************

   If m_IconAlignment = [Align Left] Then
      IconX = m_BorderWidth + 3 + m_BorderWidth
   Else
      If m_Iconsize = [Size To Header] Then
         IconX = ScaleWidth - m_HeaderHeight - 3 - m_BorderWidth
      Else
         IconX = ScaleWidth - ScaleX(m_Icon.Width, vbHimetric, vbPixels) - 3 - m_BorderWidth
      End If
   End If

End Function

Private Function IsThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks for existence of a picture by checking dimensions.             *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then
         IsThere = Pic.Width <> 0
      End If
   End If

End Function

Private Sub PaintGradient(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, _
                         ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, _
                         ByVal Angle As Single, ByVal bMOut As Boolean)

'*************************************************************************
'* Carles P.V.'s routine, modified by Matthew R. Usner for middle-out    *
'* gradient capability.  Original submission at PSC, txtCodeID=60580.    *
'*************************************************************************

   Dim uBIH      As BITMAPINFOHEADER
   Dim lBits()   As Long
   Dim lGrad()   As Long, lGrad2() As Long

   Dim lClr      As Long
   Dim R1        As Long, G1 As Long, b1 As Long
   Dim R2        As Long, G2 As Long, b2 As Long
   Dim dR        As Long, dG As Long, dB As Long

   Dim Scan      As Long
   Dim i         As Long, j As Long, k As Long
   Dim jIn       As Long
   Dim iEnd      As Long, jEnd As Long
   Dim Offset    As Long

   Dim lQuad     As Long
   Dim AngleDiag As Single
   Dim AngleComp As Single

   Dim g         As Long
   Dim luSin     As Long, luCos As Long
 
   If (Width > 0 And Height > 0) Then

'     when angle is >= 91 and <= 270, the colors
'     invert in MiddleOut mode.  This corrects that.
      If bMOut And Angle >= 91 And Angle <= 270 Then
         g = Color1
         Color1 = Color2
         Color2 = g
      End If

'     -- Right-hand [+] (ox=0º)
      Angle = -Angle + 90

'     -- Normalize to [0º;360º]
      Angle = Angle Mod 360
      If (Angle < 0) Then
         Angle = 360 + Angle
      End If

'     -- Get quadrant (0 - 3)
      lQuad = Angle \ 90

'     -- Normalize to [0º;90º]
        Angle = Angle Mod 90

'     -- Calc. gradient length ('distance')
      If (lQuad Mod 2 = 0) Then
         AngleDiag = Atn(Width / Height) * TO_DEG
      Else
         AngleDiag = Atn(Height / Width) * TO_DEG
      End If
      AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
      Angle = Angle * TO_RAD
      g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem

'     -- Decompose colors
      If (lQuad > 1) Then
         lClr = Color1
         Color1 = Color2
         Color2 = lClr
      End If
      R1 = (Color1 And &HFF&)
      G1 = (Color1 And &HFF00&) \ 256
      b1 = (Color1 And &HFF0000) \ 65536
      R2 = (Color2 And &HFF&)
      G2 = (Color2 And &HFF00&) \ 256
      b2 = (Color2 And &HFF0000) \ 65536

'     -- Get color distances
      dR = R2 - R1
      dG = G2 - G1
      dB = b2 - b1

'     -- Size gradient-colors array
      ReDim lGrad(0 To g - 1)
      ReDim lGrad2(0 To g - 1)

'     -- Calculate gradient-colors
      iEnd = g - 1
      If (iEnd = 0) Then
'        -- Special case (1-pixel wide gradient)
         lGrad2(0) = (b1 \ 2 + b2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
         For i = 0 To iEnd
            lGrad2(i) = b1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
         Next i
      End If

'     'if block' added by Matthew R. Usner - accounts for possible MiddleOut gradient draw.
      If bMOut Then
         k = 0
         For i = 0 To iEnd Step 2
            lGrad(k) = lGrad2(i)
            k = k + 1
         Next i
         For i = iEnd - 1 To 1 Step -2
            lGrad(k) = lGrad2(i)
            k = k + 1
         Next i
      Else
         For i = 0 To iEnd
            lGrad(i) = lGrad2(i)
         Next i
      End If

'     -- Size DIB array
      ReDim lBits(Width * Height - 1) As Long
      iEnd = Width - 1
      jEnd = Height - 1
      Scan = Width

'     -- Render gradient DIB
      Select Case lQuad

         Case 0, 2
            luSin = Sin(Angle) * INT_ROT
            luCos = Cos(Angle) * INT_ROT
            Offset = 0
            jIn = 0
            For j = 0 To jEnd
               For i = 0 To iEnd
                  lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
               Next i
               jIn = jIn + luCos
               Offset = Offset + Scan
            Next j

         Case 1, 3
            luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
            luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
            Offset = jEnd * Scan
            jIn = 0
            For j = 0 To jEnd
               For i = 0 To iEnd
                  lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
               Next i
               jIn = jIn + luCos
               Offset = Offset - Scan
            Next j

      End Select

'     -- Define DIB header
      With uBIH
         .biSize = 40
         .biPlanes = 1
         .biBitCount = 32
         .biWidth = Width
         .biHeight = Height
      End With

'     -- Paint it!
      Call StretchDIBits(hdc, X, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)

    End If
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Public Methods >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Sub ExpandContainer()
   If Expanded Then
      Exit Sub
   End If
   UserControl.Cls
   m_Expanded = True
   UserControl.Height = PreviousHeight
   DoEvents
End Sub

Public Sub CollapseContainer()
   If Not (Expanded) Then
      Exit Sub
   End If
   PreviousHeight = UserControl.Height
   UserControl.Cls
   UserControl.BackColor = UserControl.Ambient.BackColor
   m_Expanded = False
   UserControl.Height = m_HeaderHeight
   DoEvents
End Sub

Public Sub Refresh()
   UserControl.Refresh
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Properties >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_InitProperties()

'*************************************************************************
'* initialize properties to the default constants.                       *
'*************************************************************************

   Set m_Icon = Nothing
   Set m_CaptionFont = Ambient.Font
   Set m_Picture = LoadPicture("")
   m_HeaderAngle = m_def_HeaderAngle
   m_BackAngle = m_def_BackAngle
   m_HeaderColor2 = m_def_HeaderColor2
   m_HeaderColor1 = m_def_HeaderColor1
   m_BackColor2 = m_def_BackColor2
   m_BackColor1 = m_def_BackColor1
   m_BorderColor = m_def_BorderColor
   m_CaptionColor = m_def_CaptionColor
   m_Caption = m_def_Caption
   m_Alignment = m_def_Alignment
   m_HeaderHeight = m_def_hHeight
   m_Enabled = m_def_Enabled
   m_BorderWidth = m_def_BorderWidth
   m_BackMiddleOut = m_def_BackMiddleOut
   m_HeaderMiddleOut = m_def_HeaderMiddleOut
   m_HeaderVisible = m_def_HeaderVisible
   m_CurveTopLeft = m_def_CurveTopLeft
   m_CurveTopRight = m_def_CurveTopRight
   m_CurveBottomLeft = m_def_CurveBottomLeft
   m_CurveBottomRight = m_def_CurveBottomRight
   m_Theme = m_def_Theme
   m_Expanded = m_def_Expanded
   m_AutoResize = m_def_AutoResize
   m_AutoResizeEvent = m_def_AutoResizeEvent
   m_Movable = m_def_Movable
   m_IconAlignment = m_def_IconAlignment
   m_PictureMode = m_def_PictureMode

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*************************************************************************
'* read properties in the property bag.                                  *
'*************************************************************************

    With PropBag
        Set m_Picture = .ReadProperty("Picture", Nothing)
        Set m_Icon = .ReadProperty("HeaderIcon", Nothing)
        Set m_CaptionFont = .ReadProperty("CaptionFont", Ambient.Font)
        m_Iconsize = .ReadProperty("IconSize", m_def_Iconsize)
        m_HeaderAngle = .ReadProperty("HeaderAngle", m_def_HeaderAngle)
        m_BackAngle = .ReadProperty("BackAngle", m_def_BackAngle)
        m_HeaderColor2 = .ReadProperty("HeaderColor2", m_def_HeaderColor2)
        m_HeaderColor1 = .ReadProperty("HeaderColor1", m_def_HeaderColor1)
        m_BackColor2 = .ReadProperty("BackColor2", m_def_BackColor2)
        m_BackColor1 = .ReadProperty("BackColor1", m_def_BackColor1)
        m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
        m_CaptionColor = .ReadProperty("CaptionColor", m_def_CaptionColor)
        m_Caption = .ReadProperty("Caption", m_def_Caption)
        m_Alignment = .ReadProperty("CaptionAlignment", m_def_Alignment) 'modified by Franck Nunes
        m_HeaderHeight = .ReadProperty("HeaderHeight", m_def_hHeight)
        m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
        m_BorderWidth = .ReadProperty("BorderWidth", m_def_BorderWidth)
        m_BackMiddleOut = .ReadProperty("BackMiddleOut", m_def_BackMiddleOut)
        m_HeaderMiddleOut = .ReadProperty("HeaderMiddleOut", m_def_HeaderMiddleOut)
        m_HeaderVisible = .ReadProperty("HeaderVisible", m_def_HeaderVisible)
        m_CurveTopLeft = .ReadProperty("CurveTopLeft", m_def_CurveTopLeft)
        m_CurveTopRight = .ReadProperty("CurveTopRight", m_def_CurveTopRight)
        m_CurveBottomLeft = .ReadProperty("CurveBottomLeft", m_def_CurveBottomLeft)
        m_CurveBottomRight = .ReadProperty("CurveBottomRight", m_def_CurveBottomRight)
        m_Theme = .ReadProperty("Theme", m_def_Theme)
        m_Expanded = .ReadProperty("Expanded", m_def_Expanded)
        m_AutoResize = .ReadProperty("AutoResize", m_def_AutoResize)
        m_AutoResizeEvent = .ReadProperty("AutoResizeOn", m_def_AutoResizeEvent)
        m_Movable = .ReadProperty("Movable", m_def_Movable)
        m_IconAlignment = .ReadProperty("IconAlignment", m_def_IconAlignment)
        m_PictureMode = PropBag.ReadProperty("PictureMode", m_def_PictureMode)
    End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'*************************************************************************
'* write the properties in the property bag.                             *
'*************************************************************************

   With PropBag
      .WriteProperty "Picture", m_Picture, Nothing
      .WriteProperty "HeaderAngle", m_HeaderAngle, m_def_HeaderAngle
      .WriteProperty "BackAngle", m_BackAngle, m_def_BackAngle
      .WriteProperty "IconSize", m_Iconsize, m_def_Iconsize
      .WriteProperty "HeaderColor2", m_HeaderColor2, m_def_HeaderColor2
      .WriteProperty "HeaderColor1", m_HeaderColor1, m_def_HeaderColor1
      .WriteProperty "BackColor2", m_BackColor2, m_def_BackColor2
      .WriteProperty "BackColor1", m_BackColor1, m_def_BackColor1
      .WriteProperty "BorderColor", m_BorderColor, m_def_BorderColor
      .WriteProperty "CaptionColor", m_CaptionColor, m_def_CaptionColor
      .WriteProperty "Caption", m_Caption, m_def_Caption
      .WriteProperty "CaptionAlignment", m_Alignment, m_def_Alignment
      .WriteProperty "HeaderHeight", m_HeaderHeight, m_def_hHeight
      .WriteProperty "CaptionFont", m_CaptionFont, Ambient.Font
      .WriteProperty "HeaderIcon", m_Icon, Nothing
      .WriteProperty "Enabled", m_Enabled, m_def_Enabled
      .WriteProperty "BorderWidth", m_BorderWidth, m_def_BorderWidth
      .WriteProperty "BackMiddleOut", m_BackMiddleOut, m_def_BackMiddleOut
      .WriteProperty "HeaderMiddleOut", m_HeaderMiddleOut, m_def_HeaderMiddleOut
      .WriteProperty "HeaderVisible", m_HeaderVisible, m_def_HeaderVisible
      .WriteProperty "CurveTopLeft", m_CurveTopLeft, m_def_CurveTopLeft
      .WriteProperty "CurveTopRight", m_CurveTopRight, m_def_CurveTopRight
      .WriteProperty "CurveBottomLeft", m_CurveBottomLeft, m_def_CurveBottomLeft
      .WriteProperty "CurveBottomRight", m_CurveBottomRight, m_def_CurveBottomRight
      .WriteProperty "Theme", m_Theme, m_def_Theme
      .WriteProperty "Expanded", m_Expanded, m_def_Expanded
      .WriteProperty "AutoResize", m_AutoResize, m_def_AutoResize
      .WriteProperty "AutoResizeOn", m_AutoResizeEvent, m_def_AutoResizeEvent
      .WriteProperty "Movable", m_Movable, m_def_Movable
      .WriteProperty "IconAlignment", m_IconAlignment, m_def_IconAlignment
      .WriteProperty "PictureMode", m_PictureMode, m_def_PictureMode
   End With

End Sub

Public Property Get AutoResize() As Boolean
Attribute AutoResize.VB_Description = "If True, container can be collapsed and expanded by double-clicking the left or right mouse button (which button is defined by AutoResizeOn property)."
Attribute AutoResize.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoResize = m_AutoResize
End Property

Public Property Let AutoResize(ByVal New_AutoResize As Boolean)
    m_AutoResize = New_AutoResize
    PropertyChanged "AutoResize"
End Property

Public Property Get AutoResizeOn() As AutoResizeEvent
Attribute AutoResizeOn.VB_Description = "Determines whether the right or left mouse button can be double-clicked to collapse and expand container when AutoResize property is set to True."
Attribute AutoResizeOn.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoResizeOn = m_AutoResizeEvent
End Property

Public Property Let AutoResizeOn(ByVal vNewValue As AutoResizeEvent)
   m_AutoResizeEvent = vNewValue
   PropertyChanged "AutoResizeOn"
End Property

Public Property Get BackAngle() As Single
Attribute BackAngle.VB_Description = "The angle, in degrees, of the container background gradient."
Attribute BackAngle.VB_ProcData.VB_Invoke_Property = ";BG Graphics"
   BackAngle = m_BackAngle
End Property

Public Property Let BackAngle(ByVal New_BackAngle As Single)
'  do some bounds checking.
   If New_BackAngle > 360 Then
      New_BackAngle = 360
   ElseIf New_BackAngle < 0 Then
      New_BackAngle = 0
   End If
   m_BackAngle = New_BackAngle
   PropertyChanged "BackAngle"
   RedrawControl
End Property

Public Property Get BackColor1() As OLE_COLOR
Attribute BackColor1.VB_Description = "The first color of the container background gradient."
Attribute BackColor1.VB_ProcData.VB_Invoke_Property = ";BG Graphics"
   BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
   m_BackColor1 = New_BackColor1
   PropertyChanged "BackColor1"
   RedrawControl
End Property

Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "The second color of the container background gradient."
Attribute BackColor2.VB_ProcData.VB_Invoke_Property = ";BG Graphics"
   BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
   m_BackColor2 = New_BackColor2
   PropertyChanged "BackColor2"
   RedrawControl
End Property

Public Property Get BackMiddleOut() As Boolean
Attribute BackMiddleOut.VB_Description = "If True, the container background is drawn in middle-out mode (Color1>Color2>Color1)."
Attribute BackMiddleOut.VB_ProcData.VB_Invoke_Property = ";BG Graphics"
   BackMiddleOut = m_BackMiddleOut
End Property

Public Property Let BackMiddleOut(ByVal New_BackMiddleOut As Boolean)
   m_BackMiddleOut = New_BackMiddleOut
   PropertyChanged "BackMiddleOut"
   RedrawControl
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "The color of the container's border."
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   m_BorderColor = New_BorderColor
   PropertyChanged "BorderColor"
   RedrawControl
End Property

Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "The width, in pixels, of the container's border."
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
   m_BorderWidth = New_BorderWidth
   PropertyChanged "BorderWidth"
   RedrawControl
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "The caption to be displayed in the container's header."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Caption"
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   RedrawControl
End Property

Public Property Get CaptionAlignment() As MC_CaptionAlignment
Attribute CaptionAlignment.VB_Description = "The alignment to use when displaying the caption (left, right, or center justification)."
Attribute CaptionAlignment.VB_ProcData.VB_Invoke_Property = ";Caption"
   CaptionAlignment = m_Alignment
End Property

Public Property Let CaptionAlignment(ByVal vNewAlignment As MC_CaptionAlignment)
   m_Alignment = vNewAlignment
   PropertyChanged "CaptionAlignment"
   RedrawControl
End Property

Public Property Get CaptionColor() As OLE_COLOR
Attribute CaptionColor.VB_Description = "The color to use when displaying the caption."
Attribute CaptionColor.VB_ProcData.VB_Invoke_Property = ";Caption"
   CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
   m_CaptionColor = New_CaptionColor
   PropertyChanged "CaptionColor"
   RedrawControl
End Property

Public Property Get CaptionFont() As Font
Attribute CaptionFont.VB_Description = "The font to use when displaying the caption."
Attribute CaptionFont.VB_ProcData.VB_Invoke_Property = ";Caption"
   Set CaptionFont = m_CaptionFont
End Property

Public Property Set CaptionFont(ByVal vNewCaptionFont As Font)
   Set m_CaptionFont = vNewCaptionFont
   PropertyChanged "CaptionFont"
   RedrawControl
End Property

Public Property Get CurveBottomLeft() As Long
Attribute CurveBottomLeft.VB_Description = "The amount of curvature for the bottom left corner of the container."
Attribute CurveBottomLeft.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveBottomLeft = m_CurveBottomLeft
End Property

Public Property Let CurveBottomLeft(ByVal New_CurveBottomLeft As Long)
   m_CurveBottomLeft = New_CurveBottomLeft
   PropertyChanged "CurveBottomLeft"
   RedrawControl
End Property

Public Property Get CurveBottomRight() As Long
Attribute CurveBottomRight.VB_Description = "The amount of curvature for the bottom right corner of the container."
Attribute CurveBottomRight.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveBottomRight = m_CurveBottomRight
End Property

Public Property Let CurveBottomRight(ByVal New_CurveBottomRight As Long)
   m_CurveBottomRight = New_CurveBottomRight
   PropertyChanged "CurveBottomRight"
   RedrawControl
End Property

Public Property Get CurveTopLeft() As Long
Attribute CurveTopLeft.VB_Description = "The amount of curvature for the top left corner of the container."
Attribute CurveTopLeft.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveTopLeft = m_CurveTopLeft
End Property

Public Property Let CurveTopLeft(ByVal New_CurveTopLeft As Long)
   m_CurveTopLeft = New_CurveTopLeft
   PropertyChanged "CurveTopLeft"
   RedrawControl
End Property

Public Property Get CurveTopRight() As Long
Attribute CurveTopRight.VB_Description = "The amount of curvature for the top right corner of the container."
Attribute CurveTopRight.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   CurveTopRight = m_CurveTopRight
End Property

Public Property Let CurveTopRight(ByVal New_CurveTopRight As Long)
   m_CurveTopRight = New_CurveTopRight
   PropertyChanged "CurveTopRight"
   RedrawControl
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   PropertyChanged "Enabled"
End Property

Public Property Get Expanded() As Boolean
   Expanded = m_Expanded
End Property

Public Property Get HeaderAngle() As Single
Attribute HeaderAngle.VB_Description = "The angle (in degrees) of the header gradient."
Attribute HeaderAngle.VB_ProcData.VB_Invoke_Property = ";Header Graphics"
   HeaderAngle = m_HeaderAngle
End Property

Public Property Let HeaderAngle(ByVal New_HeaderAngle As Single)
'  do some bounds checking.
   If New_HeaderAngle > 360 Then
      New_HeaderAngle = 360
   ElseIf New_HeaderAngle < 0 Then
      New_HeaderAngle = 0
   End If
   m_HeaderAngle = New_HeaderAngle
   PropertyChanged "HeaderAngle"
   RedrawControl
End Property

Public Property Get HeaderColor1() As OLE_COLOR
Attribute HeaderColor1.VB_Description = "The first color of the header gradient."
Attribute HeaderColor1.VB_ProcData.VB_Invoke_Property = ";Header Graphics"
   HeaderColor1 = m_HeaderColor1
End Property

Public Property Let HeaderColor1(ByVal New_HeaderColor1 As OLE_COLOR)
   m_HeaderColor1 = New_HeaderColor1
   PropertyChanged "HeaderColor1"
   RedrawControl
End Property

Public Property Get HeaderColor2() As OLE_COLOR
Attribute HeaderColor2.VB_Description = "The second color of the header gradient."
Attribute HeaderColor2.VB_ProcData.VB_Invoke_Property = ";Header Graphics"
   HeaderColor2 = m_HeaderColor2
End Property

Public Property Let HeaderColor2(ByVal New_HeaderColor2 As OLE_COLOR)
   m_HeaderColor2 = New_HeaderColor2
   PropertyChanged "HeaderColor2"
   RedrawControl
End Property

Public Property Get HeaderHeight() As Long
Attribute HeaderHeight.VB_Description = "The height, in pixels, of the container's header."
Attribute HeaderHeight.VB_ProcData.VB_Invoke_Property = ";Header Graphics"
   HeaderHeight = m_HeaderHeight
End Property

Public Property Let HeaderHeight(ByVal vNewHeight As Long)
   m_HeaderHeight = vNewHeight
   PropertyChanged "HeaderHeight"
   RedrawControl
End Property

Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "The icon or small bitmap to display in the header."
Attribute Icon.VB_ProcData.VB_Invoke_Property = ";Header Graphics"
   Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal vNewIcon As Picture)
   Set m_Icon = vNewIcon
   PropertyChanged "HeaderIcon"
   RedrawControl
End Property

Public Property Get HeaderMiddleOut() As Boolean
Attribute HeaderMiddleOut.VB_Description = "If True, the header gradient is drawn in middle-out mode (Color1>Color2>Color1)."
Attribute HeaderMiddleOut.VB_ProcData.VB_Invoke_Property = ";Header Graphics"
   HeaderMiddleOut = m_HeaderMiddleOut
End Property

Public Property Let HeaderMiddleOut(ByVal New_HeaderMiddleOut As Boolean)
   m_HeaderMiddleOut = New_HeaderMiddleOut
   PropertyChanged "HeaderMiddleOut"
   RedrawControl
End Property

Public Property Get HeaderVisible() As Boolean
Attribute HeaderVisible.VB_Description = "If True, the header is visible."
Attribute HeaderVisible.VB_ProcData.VB_Invoke_Property = ";Header Graphics"
   HeaderVisible = m_HeaderVisible
End Property

Public Property Let HeaderVisible(ByVal New_HeaderVisible As Boolean)
   m_HeaderVisible = New_HeaderVisible
   PropertyChanged "HeaderVisible"
   RedrawControl
End Property

Public Property Get IconSize() As IconSizeEnum
Attribute IconSize.VB_Description = "Allows the icon to be displayed regular size, or so it fits the header"
Attribute IconSize.VB_ProcData.VB_Invoke_Property = ";Header Graphics"
   IconSize = m_Iconsize
End Property

Public Property Let IconSize(ByVal New_IconSize As IconSizeEnum)
   m_Iconsize = New_IconSize
   PropertyChanged "IconSize"
   RedrawControl
End Property

Public Property Get Theme() As MC_Themes
Attribute Theme.VB_Description = "Allows selection of 1 of 12 predefined XP-style color schemes."
Attribute Theme.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As MC_Themes)
   m_Theme = New_Theme
   PropertyChanged "Theme"
   ApplyTheme
End Property

Public Property Get Movable() As Boolean
Attribute Movable.VB_Description = "If True, the container can be moved around the form by dragging the header."
Attribute Movable.VB_ProcData.VB_Invoke_Property = ";Behavior"
   Movable = m_Movable
End Property

Public Property Let Movable(ByVal New_Movable As Boolean)
   m_Movable = New_Movable
   PropertyChanged "Movable"
End Property

Public Property Get IconAlignment() As IconAlignmentOptions
Attribute IconAlignment.VB_Description = "Allows the icon to be displayed on either the left or right side of the header."
Attribute IconAlignment.VB_ProcData.VB_Invoke_Property = ";Header Graphics"
   IconAlignment = m_IconAlignment
End Property

Public Property Let IconAlignment(ByVal New_IconAlignment As IconAlignmentOptions)
   m_IconAlignment = New_IconAlignment
   PropertyChanged "IconAlignment"
   RedrawControl
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
   hWnd = UserControl.hWnd
End Property

Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
   hdc = UserControl.hdc
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "The picture to display in lieu of a gradient in the container's background."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";BG Graphics"
   Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
   Set m_Picture = New_Picture
   PropertyChanged "Picture"
   RedrawControl
End Property

Public Property Get PictureMode() As MC_PictureModeOptions
Attribute PictureMode.VB_Description = "Allows background bitmap to be displayed in Normal, Tiled or Stretched formats."
Attribute PictureMode.VB_ProcData.VB_Invoke_Property = ";BG Graphics"
   PictureMode = m_PictureMode
End Property

Public Property Let PictureMode(ByVal New_PictureMode As MC_PictureModeOptions)
   m_PictureMode = New_PictureMode
   RedrawControl
   PropertyChanged "PictureMode"
End Property
