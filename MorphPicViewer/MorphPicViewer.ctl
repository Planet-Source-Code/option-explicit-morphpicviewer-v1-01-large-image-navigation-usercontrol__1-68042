VERSION 5.00
Begin VB.UserControl MorphPicViewer 
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
   ToolboxBitmap   =   "MorphPicViewer.ctx":0000
End
Attribute VB_Name = "MorphPicViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'*************************************************************************
'* MorphPicViewer v1.01 - Ownerdrawn large image viewer usercontrol.     *
'* By Matthew R. Usner, February, 2007, for www.planet-source-code.com   *
'* Update 24 Mar 2007 - Added .ColorAtCursor property by request.        *
'* The most up-to-date version of this control can always be found at:   *
'* www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=68042&lngWId=1       *
'* Copyright 2007-2008, Matthew R. Usner.  All rights reserved.          *
'*************************************************************************
'*            ...A small control for viewing large images...             *
'* I've seen a lot of projects on PSC for scrolling large images and     *
'* they all seem to involve slapping a couple of scrollbars on a VB      *
'* PictureBox. Navigating large images by fumbling with scrollbars is    *
'* very awkward, for me anyway. So just for fun (I have no use for this) *
'* I cooked up this little beastie. Its ClickNavigation feature provides *
'* a much faster [is 'instantly' fast enough for you? ;)], more natural  *
'* way to navigate large images than anything I've seen so far on PSC.   *
'* Features include:                                                     *
'* - Image display in either Normal or Stretched (display-to-fit) modes. *
'* - Aspect ratio may be optionally maintained in Stretched viewing mode.*
'* - Two image navigation modes are available for viewing large images - *
'*   ClickNavigation and DragNavigation.  ClickNavigation is a unique    *
'*   feature that displays entire image when the Ctrl key is held down,  *
'*   maintaining aspect ratio if mandated by the .KeepAspectRatio prop-  *
'*   erty. A rectangle outlines the portion of the image that's currently*
'*   being displayed in Normal mode.  Clicking on the image moves the    *
'*   rectangle, and when Normal view is reestablished by releasing Ctrl, *
'*   the view is changed to the new coordinates.  DragNavigation allows  *
'*   user to simply drag the image around in Normal mode until the des-  *
'*   ired area is in view.                                               *
'* - Any portion of image can also be displayed via code using the       *
'*   .DisplayImage method, supplying the X and Y coordinates of the      *
'*   upper left hand corner of the image portion you wish to display.    *
'* - Control can be used as a simple PictureBox replacement; image nav-  *
'*   igation can be disabled for straightforward image display (although *
'*   image stretching and aspect ratio features are always available).   *
'*   Control is also a container as is the standard VB PictureBox.       *
'*************************************************************************
'* Miscellaneous Notes:                                                  *
'*  - Since this uses region code, DO NOT use the End button in the IDE. *
'*    Use Unload Me in form code.  This ensures proper object deletion.  *
'*  - When the term 'Display Area' is used in this control, it refers to *
'*    the control's area not counting the border (if one is defined).    *
'*  - There's no fancy gadgetry like zooming or whatever, it was meant   *
'*    to be a large image viewer ONLY.  As many of you know I sometimes  *
'*    get carried away with features in my controls and this time I was  *
'*    determined to stick with a minimalist approach!  However, bug rep- *
'*    orts and enhancement ideas are always welcome... and votes are     *
'*    always appreciated.  Hope someone finds a use for this... enjoy.   *
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
'* This control was developed by Matthew R. Usner.                       *
'* Source code, written in Visual Basic, is freely available for non-    *
'* commercial, non-profit use.                                           *
'*                                                                       *
'* Redistributions in binary form, as part of a larger project, must     *
'* include the above acknowledgment in the end-user documentation.       *
'* Alternatively, the above acknowledgment may appear in the software    *
'* itself, if and where such third-party acknowledgments normally appear.*
'*************************************************************************
'* Credits and Thanks:                                                   *
'* Carles P.V. - original gradient generation code; modified by me.      *
'* LaVolpe - border region creation code; StdPicture.Render example.     *
'* Heriberto Mantilla Santamaria - SetStretchBltMode technique.          *
'*************************************************************************

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Private Const STRETCH_HALFTONE As Long = &H4&        ' for correctly displaying colors of stretched image.

'  used in creating trapezoidal border segments.
Private Type POINTAPI
   X                           As Long
   Y                           As Long
End Type

'  used to define various graphics areas.
Private Type RECT
   Left                        As Long
   Top                         As Long
   Right                       As Long
   Bottom                      As Long
End Type

'  declares for gradient painting.
Private Type BITMAPINFOHEADER
   biSize                      As Long
   biWidth                     As Long
   biHeight                    As Long
   biPlanes                    As Integer
   biBitCount                  As Integer
   biCompression               As Long
   biSizeImage                 As Long
   biXPelsPerMeter             As Long
   biYPelsPerMeter             As Long
   biClrUsed                   As Long
   biClrImportant              As Long
End Type

' gradient generation constants.
Private Const DIB_RGB_COLORS   As Long = 0
Private Const PI               As Single = 3.14159265358979
Private Const TO_DEG           As Single = 180 / PI
Private Const TO_RAD           As Single = PI / 180
Private Const INT_ROT          As Long = 1000

' holds gradient information for background.
Private uBIH_BG                As BITMAPINFOHEADER
Private lBits_BG()             As Long

'  gradient information for horizontal and vertical border segments.
Private SegV1uBIH              As BITMAPINFOHEADER
Private SegV1lBits()           As Long
Private SegV2uBIH              As BITMAPINFOHEADER
Private SegV2lBits()           As Long
Private SegH1uBIH              As BITMAPINFOHEADER
Private SegH1lBits()           As Long
Private SegH2uBIH              As BITMAPINFOHEADER
Private SegH2lBits()           As Long

' constants defining the four border segments.
Private Const TOP_SEGMENT      As Long = 0
Private Const RIGHT_SEGMENT    As Long = 1
Private Const BOTTOM_SEGMENT   As Long = 2
Private Const LEFT_SEGMENT     As Long = 3

Private BorderSegment(0 To 3)  As Long               ' holds region pointers for border segments.

' declares for virtual image (unstretched) bitmap.
Private VirtualDC_Image        As Long               ' handle of the created DC.
Private mMemoryBitmap_Image    As Long               ' handle of the created bitmap.
Private mOriginalBitmap_Image  As Long               ' used in destroying virtual DC.

' declares for virtual horizontal border segment bitmap.
Private VirtualDC_SegH         As Long               ' handle of the created DC.
Private mMemoryBitmap_SegH     As Long               ' handle of the created bitmap.
Private mOriginalBitmap_SegH   As Long               ' used in destroying virtual DC.

' declares for virtual vertical border segment bitmap.
Private VirtualDC_SegV         As Long               ' handle of the created DC.
Private mMemoryBitmap_SegV     As Long               ' handle of the created bitmap.
Private mOriginalBitmap_SegV   As Long               ' used in destroying virtual DC.

'  enum tied to .PictureMode property.
Public Enum MPV_PicModeOptions
   mpv_Normal
   mpv_Stretch
End Enum

'  property variables and default value constants.
Private m_BackAngle            As Single             ' background gradient display angle.
Private m_BackColor1           As OLE_COLOR          ' the first gradient color of the background.
Private m_BackColor2           As OLE_COLOR          ' the second gradient color of the background.
Private m_BackMiddleOut        As Boolean            ' flag for background middle-out gradient.
Private m_BorderColor1         As OLE_COLOR          ' color 1 of border gradient.
Private m_BorderColor2         As OLE_COLOR          ' color 2 of border gradient.
Private m_BorderMiddleOut      As Boolean            ' if True, linear border gradient is middle-out.
Private m_BorderWidth          As Integer            ' width, in pixels, of border.
Private m_ColorAtCursor        As Long               ' color of pixel mouse is hovering over.
Private m_DisBorderBGColor1    As OLE_COLOR          ' disabled first border gradient color.
Private m_DisBorderBGColor2    As OLE_COLOR          ' disabled second border gradient color.
Private m_DispX1               As Long               ' x coordinate of top left image displayed portion.
Private m_DispY1               As Long               ' y coordinate of top left image displayed portion.
Private m_Enabled              As Boolean            ' control enabled/disabled flag.
Private m_FocusBorderColor1    As OLE_COLOR          ' focus first border gradient color.
Private m_FocusBorderColor2    As OLE_COLOR          ' focus second border gradient color.
Private m_KeepAspectRatio      As Boolean            ' if True, stretched image maintains aspect ratio.
Private m_NavModeEnabled       As Boolean            ' if true, Navigation Mode can be used.
Private m_NavRectColor         As OLE_COLOR          ' the color of the nav mode viewport rectangle.
Private m_PicHeight            As Long               ' image height, in pixels.
Private m_Picture              As Picture            ' image to display.
Private m_PictureMode          As MPV_PicModeOptions ' normal or stretched image.
Private m_PicWidth             As Long               ' image width, in pixels.
Private m_PicX                 As Long               ' X coordinate of top left current displayed portion.
Private m_PicY                 As Long               ' Y coordinate of top left current displayed portion.

'  events.
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Resize()

'  miscellaneous control variables and constants.
Private ActiveBorderColor1     As Long               ' first gradient color of active border.
Private ActiveBorderColor2     As Long               ' second gradient color of active border.
Private ActiveBGColor1         As Long               ' first gradient color of active background.
Private ActiveBGColor2         As Long               ' second gradient color of active background.
Private HasFocus               As Boolean            ' indicates if control has focus.
Private ClickNavigationActive  As Boolean            ' flag to indicate ClickNavigation mode.
Private DragNavigationActive   As Boolean            ' flag to indicate DragNavigation mode.

'  defines the part of the image currently being displayed (normal mode).
Private CurrentXY              As RECT

'  control X and Y coordinates where to display the top left corner of the image in Normal display mode.
Private DisplayXY              As POINTAPI

'  current normal-mode image portion coordinates prior to moving image in DragNavigation (MouseDown event).
Private SaveXY                 As RECT

' constant describing if the mouse pointer is currently located within the image (not just display area).
Private Const MOUSE_IN_IMAGE   As Long = 1
Private MouseLocation          As Long               ' holds the above constant if pointer over image.

' passed to the DisplayImage procedure.
Private Const IMAGE_NORMAL     As Long = 0           ' display image in normal size.
Private Const IMAGE_STRETCHED  As Long = 1           ' display image sized to fit display area.

' mouse coordinates when MouseDown event triggers DragNavigation mode.
Private SaveMouseXY            As POINTAPI

' navigation rectangle coordinates and dimensions.
Private NavXY                  As RECT               ' coordinates of navigation rectangle.
Private NavWidth               As Long               ' navigation rectangle width, in pixels.
Private NavHeight              As Long               ' navigation rectangle height, in pixels.

' image display area dimensions (excludes control's border if one is defined).
Private DisplayAreaWidth       As Long               ' display area width, in pixels.
Private DisplayAreaHeight      As Long               ' display area height, in pixels.

' the dimensions and display area coordinates of the top left corner of the image when displayed in stretched
' mode (maintaining aspect ratio).  These coordinates are used when drawing the navigation rectangle.
Private AspRatXY               As POINTAPI
Private AspRatWidth            As Long               ' width, in pixels, of image with aspect ratio.
Private AspRatHeight           As Long               ' height, in pixels, of image with aspect ratio.

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Events >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_GotFocus()

'*************************************************************************
'* updates border colors when control receives focus.                    *
'*************************************************************************

   If Not m_Enabled Then Exit Sub

   HasFocus = True
   UpdateBorder

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* activates ClickNavigation mode if .NavModeEnabled property allows it. *
'*************************************************************************

   If Not m_Enabled Then Exit Sub

'  if Ctrl pressed, we're in Normal image display mode (i.e. not Stretched mode), the image
'  is larger than the display area, and are allowing navigation, start ClickNavigation mode.
   If (Shift And vbCtrlMask) > 0 And m_PictureMode = mpv_Normal And m_NavModeEnabled And Not DragNavigationActive And Not ClickNavigationActive And Not FullImageContainedInDisplayArea Then
      ClickNavigationActive = True
      DisplayImage IMAGE_STRETCHED
   End If

   RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* cancels ClickNavigation mode if Ctrl key has been released.           *
'*************************************************************************

   If Not m_Enabled Then Exit Sub

'  if we're in ClickNavigation mode and Ctrl key has been released, exit ClickNavigation mode.
   If ClickNavigationActive And Not (Shift And vbCtrlMask) > 0 Then
      ClickNavigationActive = False
      DisplayImage IMAGE_NORMAL    ' redraw image in Normal mode.
   End If

   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub UserControl_LostFocus()

'*************************************************************************
'* updates border colors when control loses focus.                       *
'*************************************************************************

   If Not m_Enabled Then Exit Sub

   HasFocus = False
   GetActiveBorderColors
'  if we're in the ClickNavigation screen and focus is lost, return to normal display.
'  If Ctrl is still down when focus is regained, we'll reenter ClickNavigation mode.
   If ClickNavigationActive Then
      ClickNavigationActive = False
      DisplayImage IMAGE_NORMAL    ' redraw image in Normal mode.
   End If
   UpdateBorder

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* allows user to move the navigation rectangle in ClickNavigation Mode. *
'* Also sets DragNavigation mode if ClickNavigation is inactive.         *
'*************************************************************************

   If Not m_Enabled Or Button = vbRightButton Then Exit Sub

   If m_NavModeEnabled And Not ClickNavigationActive And Not FullImageContainedInDisplayArea And m_PictureMode <> [mpv_Stretch] Then

      DragNavigationActive = True
'     save the current image coordinates.
      SaveXY.Left = CurrentXY.Left
      SaveXY.Top = CurrentXY.Top
'     save the current mouse coordinates.
      SaveMouseXY.X = X
      SaveMouseXY.Y = Y
      If m_NavModeEnabled Then
         Set UserControl.MouseIcon = LoadResPicture(102, 2)   ' set cursor to grasping hand.
         UserControl.MousePointer = 99
      End If

   ElseIf ClickNavigationActive Then

'     determine the new navigation rectangle coordinates, centered on mousedown x,y.
      NavXY.Left = IIf(m_PicWidth > DisplayAreaWidth, (X - AspRatXY.X) - (NavWidth \ 2), 0)
      NavXY.Top = IIf(m_PicHeight > DisplayAreaHeight, (Y - AspRatXY.Y) - (NavHeight \ 2), 0)
      NavXY.Right = IIf(m_PicWidth > DisplayAreaWidth, NavXY.Left + NavWidth - 1, NavXY.Left + AspRatWidth - 1)
      NavXY.Bottom = IIf(m_PicHeight > DisplayAreaHeight, NavXY.Top + NavHeight - 1, NavXY.Top + AspRatHeight - 1)

'     calculate the new normal mode display coordinates.
      CurrentXY.Left = CLng(NavXY.Left / AspRatWidth * m_PicWidth)
      CurrentXY.Top = CLng(NavXY.Top / AspRatHeight * m_PicHeight)
      CurrentXY.Right = CLng(NavXY.Right / AspRatWidth * m_PicWidth)
      CurrentXY.Bottom = CLng(NavXY.Bottom / AspRatHeight * m_PicHeight)

'     ensure that visible image portion always stays completely within display area boundaries.
      AdjustVisibleImageCoordinates

'     set the properties that define the top left coordinates of current viewed image portion.
      m_DispX1 = CurrentXY.Left
      m_DispY1 = CurrentXY.Top

'     redraw the stretched-mode image, showing new navigation rectangle.
      DisplayImage IMAGE_STRETCHED

   End If

   RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* cancels DragNavigation mode if left button has been released.         *
'*************************************************************************
   
   If Not m_Enabled Or Button = vbRightButton Then Exit Sub

   If DragNavigationActive Then
      DragNavigationActive = False
'     since we were in DragNavigation mode, set cursor to pointing finger.
      If m_NavModeEnabled And m_PictureMode <> mpv_Stretch Then Set UserControl.MouseIcon = LoadResPicture(101, 2)
   End If

   RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'*************************************************************************
'* if in DragNavigation mode, moving the mouse with the left button down *
'* moves the image in the direction of mouse pointer movement.           *
'*************************************************************************

   If Not m_Enabled Then Exit Sub

'  make sure mouse pointer is set properly according to control state.
   If m_PictureMode = mpv_Stretch And UserControl.MousePointer <> vbDefault Then
      UserControl.MousePointer = vbDefault
   ElseIf m_NavModeEnabled And m_PictureMode <> mpv_Stretch And UserControl.MousePointer = vbDefault Then
      Set UserControl.MouseIcon = LoadResPicture(101, 2) ' set pointer to pointing finger.
      UserControl.MousePointer = 99
   End If

'  determine the part of the control the mouse pointer is currently over.
   MouseLocation = CurrentMouseLocation(X, Y)
   m_ColorAtCursor = IIf(MouseLocation = MOUSE_IN_IMAGE, GetPixel(hdc, X, Y), -1)
   
   'If MouseLocation = MOUSE_IN_IMAGE Then
   '   m_ColorAtCursor = GetPixel(hdc, X, Y)
   ' Else
   '   m_ColorAtCursor = -1
   'End If

   If (Not FullImageContainedInDisplayArea) And DragNavigationActive And HasFocus Then

'     determine the new top left / bottom right corner coordinates of visible portion of image.
'     can only move horizontally if image width is greater than display area width.
      If m_PicWidth > DisplayAreaWidth Then
'        (x - SaveMouseXY.x) represents the horizontal pixel distance mouse has traveled since MouseDown.
         CurrentXY.Left = SaveXY.Left - (X - SaveMouseXY.X)
         CurrentXY.Right = CurrentXY.Left + DisplayAreaWidth - 1
      End If

'     can only move vertically if image height is greater than display area height.
      If m_PicHeight > DisplayAreaHeight Then
'        (y - SaveMouseXY.y) represents the vertical pixel distance mouse has traveled since MouseDown.
         CurrentXY.Top = SaveXY.Top - (Y - SaveMouseXY.Y)
         CurrentXY.Bottom = CurrentXY.Top + DisplayAreaHeight - 1
      End If

'     ensure that visible image portion always stays completely within display area boundaries.
      AdjustVisibleImageCoordinates
      m_DispX1 = CurrentXY.Left ' update X coordinate of upper left displayed portion property.
      m_DispY1 = CurrentXY.Top  ' update Y coordinate of upper left displayed portion property.

'     display the appropriate portion of the image.
      DisplayImage IMAGE_NORMAL

   End If

'  set the PicX and PicY properties.  Only available in Normal view mode.
   m_PicX = IIf(Not MouseLocation = MOUSE_IN_IMAGE Or m_PictureMode = [mpv_Stretch], -1, IIf(m_PicWidth <= DisplayAreaWidth, X - DisplayXY.X, CurrentXY.Left + X - m_BorderWidth))
   m_PicY = IIf(Not MouseLocation = MOUSE_IN_IMAGE Or m_PictureMode = [mpv_Stretch], -1, IIf(m_PicHeight <= DisplayAreaHeight, Y - DisplayXY.Y, CurrentXY.Top + Y - m_BorderWidth))

   RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_Resize()

'*************************************************************************
'* controls redrawing of container when resized.                         *
'*************************************************************************

   If Ambient.UserMode Then Exit Sub
   InitializeControl
   RaiseEvent Resize

End Sub

Private Sub UserControl_Terminate()
   
'*************************************************************************
'* last event in lifecycle of the control.  Destroy all created objects. *
'*************************************************************************

'  delete border segment region objects.
   DeleteBorderSegmentObjects

'  destroy the virtual DCs used to store segment gradients.
   DestroyVirtualDC VirtualDC_SegH, mMemoryBitmap_SegH, mOriginalBitmap_SegH
   DestroyVirtualDC VirtualDC_SegV, mMemoryBitmap_SegV, mOriginalBitmap_SegV

'  destroy the image virtual DC.
   DestroyVirtualDC VirtualDC_Image, mMemoryBitmap_Image, mOriginalBitmap_Image

End Sub

Private Sub DeleteBorderSegmentObjects()

'*************************************************************************
'* destroys the border segment objects if they exist, to prevent leaks.  *
'*************************************************************************

   If BorderSegment(TOP_SEGMENT) Then DeleteObject BorderSegment(TOP_SEGMENT)
   If BorderSegment(RIGHT_SEGMENT) Then DeleteObject BorderSegment(RIGHT_SEGMENT)
   If BorderSegment(BOTTOM_SEGMENT) Then DeleteObject BorderSegment(BOTTOM_SEGMENT)
   If BorderSegment(LEFT_SEGMENT) Then DeleteObject BorderSegment(LEFT_SEGMENT)

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Graphics >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub InitializeControl()

'*************************************************************************
'* sets up graphics and initial image display.                           *
'*************************************************************************

   GetActiveBorderColors          ' determine the border colorset to use.
   GetActiveBackgroundColors      ' determine the background colorset to use.
   CalculateBackgroundGradient    ' set up the background gradient in the background virtual DC.
   InitializeBorderGraphics       ' set up the border segments.
   ProcessImage                   ' determine image statistics.
   RedrawControl                  ' paint the viewer control.

End Sub

Private Sub RedrawControl()

'*************************************************************************
'* master routine for painting of control.                               *
'*************************************************************************

   If m_BorderWidth > 0 Then DisplayBorder
   SetBackGround
   UserControl.Refresh

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* converts color long COLORREF for api coloring purposes.               *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then TranslateColor = -1

End Function

Private Sub SetBackGround()

'*************************************************************************
'* displays the image, or a background gradient if no image specified.   *
'*************************************************************************

   If m_Enabled Then
      If IsPictureThere(m_Picture) Then
         If m_PictureMode = mpv_Normal Then
            If ClickNavigationActive Then
               DisplayImage IMAGE_STRETCHED ' display Stretched.
            Else
               DisplayImage IMAGE_NORMAL    ' display normal size, centered if applicable.
            End If
         Else    ' m_PictureMode = mpv_Stretch.
            DisplayImage IMAGE_STRETCHED    ' stretch/shrink the picture to fit the control.
         End If
      Else
         DisplayGradientBackground          ' no image - display gradient, accounting for border.
      End If
   Else
      DisplayGradientBackground             ' background will display in disabled colors.
   End If

End Sub

Private Sub DisplayGradientBackground()

'*************************************************************************
'* paints the background gradient on control, accounting for border.     *
'*************************************************************************

   Call StretchDIBits(hdc, m_BorderWidth, m_BorderWidth, ScaleWidth - (m_BorderWidth * 2), ScaleHeight - (m_BorderWidth * 2), _
                      m_BorderWidth, m_BorderWidth, ScaleWidth - (m_BorderWidth * 2), ScaleHeight - (m_BorderWidth * 2), lBits_BG(0), uBIH_BG, DIB_RGB_COLORS, vbSrcCopy)

End Sub

Private Function IsPictureThere(ByVal Pic As StdPicture) As Boolean

'*************************************************************************
'* checks for existence of a picture.  Thanks to Roger Gilchrist.        *
'*************************************************************************

   If Not Pic Is Nothing Then
      If Pic.Height <> 0 Then IsPictureThere = Pic.Width <> 0
   End If

End Function

Private Sub UpdateBorder()

'*************************************************************************
'* redraws border when GotFocus and LostFocus events are detected.       *
'*************************************************************************

   GetActiveBorderColors
   If m_BorderWidth < 1 Then Exit Sub
   DisplayBorder
   UserControl.Refresh

End Sub

Private Sub DisplayBorder()

'*************************************************************************
'* displays the four border segments.                                    *
'*************************************************************************

'  display each border segment.
   DisplaySegment TOP_SEGMENT, 0, 0
   DisplaySegment LEFT_SEGMENT, 0, 0
   DisplaySegment RIGHT_SEGMENT, ScaleWidth - m_BorderWidth, 0
   DisplaySegment BOTTOM_SEGMENT, -1, ScaleHeight - m_BorderWidth

End Sub

Private Function CurrentMouseLocation(ByVal X As Single, ByVal Y As Single) As Long

'*************************************************************************
'* sets the MouseLocation variable according to mouse pointer location.  *
'*************************************************************************

   If Not ClickNavigationActive And Not DragNavigationActive Then
      If X >= DisplayXY.X And X < (DisplayXY.X + m_PicWidth) And X < ScaleWidth - m_BorderWidth And Y >= DisplayXY.Y And Y < (DisplayXY.Y + m_PicHeight) And Y < ScaleHeight - m_BorderWidth Then
         CurrentMouseLocation = MOUSE_IN_IMAGE
      End If
   End If

End Function

Private Sub DisplaySegment(ByVal SegmentNdx As Long, ByVal StartX As Long, ByVal StartY As Long)

'*************************************************************************
'* displays one border segment.  Border segment gradients are displayed  *
'* to virtual bitmaps on the fly so that correct gradient orientation    *
'* is maintained if the .MiddleOut property is set to False.             *
'*************************************************************************

'  position the border segment region in the correct location.
   MoveRegionToXY BorderSegment(SegmentNdx), StartX, StartY

   Select Case SegmentNdx

      Case LEFT_SEGMENT
         PaintVerticalGradient m_BorderWidth, ScaleHeight, SegV1uBIH, SegV1lBits()
         BlitToRegion VirtualDC_SegV, hdc, m_BorderWidth, ScaleHeight, BorderSegment(SegmentNdx), StartX, StartY

      Case RIGHT_SEGMENT
         If m_BorderMiddleOut Then
            PaintVerticalGradient m_BorderWidth, ScaleHeight, SegV1uBIH, SegV1lBits()
         Else
            PaintVerticalGradient m_BorderWidth, ScaleHeight, SegV2uBIH, SegV2lBits()
         End If
         BlitToRegion VirtualDC_SegV, hdc, m_BorderWidth, ScaleHeight, BorderSegment(SegmentNdx), StartX, StartY

      Case TOP_SEGMENT
         PaintHorizontalGradient m_BorderWidth, ScaleWidth, SegH1uBIH, SegH1lBits()
         BlitToRegion VirtualDC_SegH, hdc, ScaleWidth, m_BorderWidth, BorderSegment(SegmentNdx), StartX, StartY

      Case BOTTOM_SEGMENT
         If m_BorderMiddleOut Then
            PaintHorizontalGradient m_BorderWidth, ScaleWidth, SegH1uBIH, SegH1lBits()
         Else
            PaintHorizontalGradient m_BorderWidth, ScaleWidth, SegH2uBIH, SegH2lBits()
         End If
         BlitToRegion VirtualDC_SegH, hdc, ScaleWidth, m_BorderWidth, BorderSegment(SegmentNdx), StartX, StartY

   End Select

End Sub

Private Sub PaintHorizontalGradient(ByVal BorderWidth As Long, ByVal TargetWidth As Long, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

'*************************************************************************
'* paints appropriate horizontal gradient to horizontal virtual bitmap.  *
'*************************************************************************

   Call StretchDIBits(VirtualDC_SegH, 0, 0, TargetWidth, BorderWidth, 0, 1, TargetWidth, BorderWidth - 1, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)

End Sub

Private Sub PaintVerticalGradient(ByVal BorderWidth As Long, ByVal TargetHeight, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

'*************************************************************************
'* paints appropriate vertical gradient to vertical virtual bitmap.      *
'*************************************************************************

   Call StretchDIBits(VirtualDC_SegV, 0, 0, BorderWidth, TargetHeight, 1, 0, BorderWidth - 1, TargetHeight, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)

End Sub

Private Sub BlitToRegion(ByVal SourceDC As Long, ByVal DestDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal Region As Long, ByVal XPos As Long, ByVal YPos As Long)

'*************************************************************************
'* blits the contents of a source DC to a non-rectangular region in a    *
'* destination DC.  A clipping region is selected in the destination DC, *
'* then the source DC is blitted to that location.  Technique is used in *
'* this control to blit to the trapezoid-shaped border regions.          *
'*************************************************************************

'  move the region to the desired position.
   MoveRegionToXY Region, XPos, YPos
'  select a clipping region consisting of the segment parameter.
   SelectClipRgn DestDC, Region
'  blit the virtual bitmap to the control or form using the clip region as a mask.
   BitBlt DestDC, XPos, YPos, lWidth, lHeight, SourceDC, 0, 0, vbSrcCopy
'  remove the clipping region constraint from the control.
   SelectClipRgn DestDC, ByVal 0&
'  reset the region coordinates to 0,0.
   MoveRegionToXY Region, 0, 0

End Sub

Private Sub MoveRegionToXY(ByVal Rgn As Long, ByVal X As Long, ByVal Y As Long)

'*************************************************************************
'* moves the supplied region to absolute X,Y coordinates.                *
'*************************************************************************

   Dim r As RECT    ' holds current X and Y coordinates of region.

'  get the current X,Y coordinates of the region.
   GetRgnBox Rgn, r
'  shift the region to 0,0 then to X,Y.
   OffsetRgn Rgn, -r.Left + X, -r.Top + Y

End Sub

Private Function FullImageContainedInDisplayArea() As Boolean

'*************************************************************************
'* returns True if image can be fully contained in the display area.     *
'*************************************************************************

   FullImageContainedInDisplayArea = (m_PicWidth <= DisplayAreaWidth And m_PicHeight <= DisplayAreaHeight)

End Function

Private Sub AdjustVisibleImageCoordinates()

'*************************************************************************
'* ensures that visible image portion always stays completely within     *
'* display area boundaries.                                              *
'*************************************************************************

   If CurrentXY.Top < 0 Then
      CurrentXY.Top = 0
      CurrentXY.Bottom = DisplayAreaHeight - 1
   ElseIf CurrentXY.Bottom > m_PicHeight - 1 Then
      CurrentXY.Bottom = m_PicHeight - 1
      CurrentXY.Top = CurrentXY.Bottom - DisplayAreaHeight + 1
   End If

   If CurrentXY.Left < 0 Then
      CurrentXY.Left = 0
      CurrentXY.Right = DisplayAreaWidth - 1
   ElseIf CurrentXY.Right > m_PicWidth - 1 Then
      CurrentXY.Right = m_PicWidth - 1
      CurrentXY.Left = CurrentXY.Right - DisplayAreaWidth + 1
   End If

End Sub

Private Sub CreateVirtualDC(TargetDC As Long, vDC As Long, mMB As Long, mOB As Long, ByVal vWidth As Long, ByVal vHeight As Long)

'*************************************************************************
'* creates virtual bitmaps for background and cells.                     *
'*************************************************************************

'  make sure region does not already exist - just a safety net.
   If vDC <> 0 Then DestroyVirtualDC vDC, mMB, mOB

'  create a memory device context to use.
   vDC = CreateCompatibleDC(TargetDC)

'  define it as a bitmap so that drawing can be performed to the virtual DC.
   mMB = CreateCompatibleBitmap(TargetDC, vWidth, vHeight)
   mOB = SelectObject(vDC, mMB)

End Sub

Private Sub DestroyVirtualDC(ByRef vDC As Long, ByVal mMB As Long, ByVal mOB As Long)

'*************************************************************************
'* eliminates a virtual dc bitmap on control's termination.              *
'*************************************************************************

   If vDC = 0 Then Exit Sub           ' no virtual DC to delete.
   Call SelectObject(vDC, mOB)
   Call DeleteObject(mMB)
   Call DeleteDC(vDC)
   vDC = 0

End Sub

Private Sub CalculateGradient(Width As Long, Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Angle As Single, ByVal bMOut As Boolean, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

'*************************************************************************
'* Carles P.V.'s routine, modified by Matthew R. Usner for middle-out    *
'* gradient capability.  Also modified to just calculate the gradient,   *
'* not draw it.  Original submission at PSC, txtCodeID=60580.            *
'*************************************************************************

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

'     when angle is >= 91 and <= 270, the colors invert in MiddleOut mode.  This corrects that.
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

'     'if' block added by Matthew R. Usner - accounts for possible MiddleOut gradient draw.
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

   End If

End Sub

Private Function CreateDiagRectRegion(ByVal cx As Long, ByVal cy As Long, SideAStyle As Integer, SideBStyle As Integer) As Long

'**************************************************************************
'* Author: LaVolpe                                                        *
'* Original submission at txtCodeId=58562.  Thanks Keith.                 *
'* the cx & cy parameters are the respective width & height of the region *
'* the passed values may be modified which coder can use for other purp-  *
'* oses like drawing borders or calculating the client/clipping region.   *
'* SideAStyle is -1, 0 or 1 depending on horizontal/vertical shape,       *
'*            reflects the left or top side of the region                 *
'*            -1 draws left/top edge like /                               *
'*            0 draws left/top edge like  |                               *
'*            1 draws left/top edge like  \                               *
'* SideBStyle is -1, 0 or 1 depending on horizontal/vertical shape,       *
'*            reflects the right or bottom side of the region             *
'*            -1 draws right/bottom edge like \                           *
'*            0 draws right/bottom edge like  |                           *
'*            1 draws right/bottom edge like  /                           *
'**************************************************************************

   Dim tpts(0 To 4) As POINTAPI    ' holds polygonal region vertices.

   If cx > cy Then ' horizontal

'     absolute minimum width & height of a trapezoid
      If Abs(SideAStyle + SideBStyle) = 2 Then ' has 2 opposing slanted sides
         If cx < cy * 2 Then cy = cx \ 2
      End If

      If SideAStyle < 0 Then
         tpts(0).X = cy - 1
         tpts(1).X = -1
      ElseIf SideAStyle > 0 Then
         tpts(1).X = cy
      End If
      tpts(1).Y = cy

      tpts(2).X = cx + Abs(SideBStyle < 0)
      If SideBStyle > 0 Then tpts(2).X = tpts(2).X - cy
      tpts(2).Y = cy

      tpts(3).X = cx + Abs(SideBStyle < 0)
      If SideBStyle < 0 Then tpts(3).X = tpts(3).X - cy

   Else

'     absolute minimum width & height of a trapezoid
      If Abs(SideAStyle + SideBStyle) = 2 Then ' has 2 opposing slanted sides
         If cy < cx * 2 Then cx = cy \ 2
      End If

      If SideAStyle < 0 Then
         tpts(0).Y = cx - 1
         tpts(3).Y = -1
      ElseIf SideAStyle > 0 Then
         tpts(3).Y = cx - 1
         tpts(0).Y = -1
      End If

      tpts(1).Y = cy
      If SideBStyle < 0 Then tpts(1).Y = tpts(1).Y - cx
      tpts(2).X = cx

      tpts(2).Y = cy
      If SideBStyle > 0 Then tpts(2).Y = tpts(2).Y - cx
      tpts(3).X = cx

   End If

   tpts(4) = tpts(0)

   CreateDiagRectRegion = CreatePolygonRgn(tpts(0), UBound(tpts) + 1, 2)

End Function

Public Sub DisplayImage(DisplayMode As Long, Optional ByVal XPos As Long = -1, Optional ByVal YPos As Long = -1)

'*************************************************************************
'* public method/internal procedure that displays image on control,      *
'* accounting for normal/stretched display, navigation modes and current *
'* portion of image being displayed. Image is automatically centered in  *
'* middle of display area if height or width of picture is less than     *
'* height or width of display area.  If picture is equal to or larger    *
'* than display area size, image starts at X=BorderWidth, Y=BorderWidth  *
'* within display area.  If called from an app as a method, supplying    *
'* the XY coordinates of the top left corner of the portion of the       *
'* desired image portion allows you to view any part of image via code.  *
'*************************************************************************

   If Not DragNavigationActive Then DisplayGradientBackground

'  render the appropriate image portion to the control.
   If DragNavigationActive Then '  use blit to 'move' image.
      BitBlt hdc, DisplayXY.X, DisplayXY.Y, DisplayAreaWidth, DisplayAreaHeight, VirtualDC_Image, CurrentXY.Left, CurrentXY.Top, vbSrcCopy
   ElseIf ClickNavigationActive Then '  redisplay in stretched mode and display the navigation rectangle.
      DisplayStretchedModeBitmap
      DisplayNavigationRectangle
   Else
'     otherwise, display image in Normal mode according to DisplayMode (normal/stretched).
      If DisplayMode = IMAGE_NORMAL Then
         If XPos > -1 Then ' procedure was called as public method from app with X coordinate supplied.
'           make sure image stays in bounds.
            If XPos + DisplayAreaWidth > m_PicWidth Then XPos = m_PicWidth - DisplayAreaWidth
            CurrentXY.Left = XPos
            CurrentXY.Right = XPos + DisplayAreaWidth - 1
            m_DispX1 = XPos ' update upper left X coordinate of displayed portion property.
         End If
         If YPos > -1 Then ' procedure was called as public method from app with Y coordinate supplied.
'           make sure image stays in bounds.
            If YPos + DisplayAreaHeight > m_PicHeight Then YPos = m_PicHeight - DisplayAreaHeight
            CurrentXY.Top = YPos
            CurrentXY.Bottom = YPos + DisplayAreaHeight - 1
            m_DispY1 = YPos ' update upper left Y coordinate of displayed portion property.
         End If
         BitBlt hdc, DisplayXY.X, DisplayXY.Y, DisplayAreaWidth, DisplayAreaHeight, VirtualDC_Image, CurrentXY.Left, CurrentXY.Top, vbSrcCopy
      Else
         DisplayStretchedModeBitmap
      End If
   End If
   UserControl.Refresh

End Sub

Private Sub DisplayStretchedModeBitmap()

'*************************************************************************
'* displays image in stretched format.                                   *
'*************************************************************************

   Dim mOrigTone As Long    ' holds original StretchBlt mode when mode is changed.
   Dim mResult   As Long    ' for reverting back to the original StretchBlt mode.

   On Error GoTo StretchPicErr

'  from API Guide: "Map pixels from the source rectangle into blocks of pixels in the destination rectangle.
'  The average color over the destination block of pixels approximates the color of the source pixels."
'  Translation: maintain color and tone in stretched/compressed image so image isn't grainy and offcolor.
'  I found this in Heriberto Mantilla Santamaria's Task Screen submission at txtCodeId=67775, thanks buddy.
   mOrigTone = SetStretchBltMode(hdc, STRETCH_HALFTONE)
'  stretch the image onto control display area.
   StretchBlt hdc, AspRatXY.X, AspRatXY.Y, AspRatWidth, AspRatHeight, VirtualDC_Image, 0, 0, m_PicWidth, m_PicHeight, vbSrcCopy
   mResult = SetStretchBltMode(hdc, mOrigTone)

StretchPicErr:

End Sub

Private Sub DisplayNavigationRectangle()

'*************************************************************************
'* displays a 1-pixel rectangle in ClickNavigation mode that surrounds   *
'* the portion of the image being displayed in Normal mode.              *
'*************************************************************************

   Dim hBrush As Long    ' the brush pattern used to 'paint' the navigation rectangle.

'  given the position and dimensions of the stretched image, plot the equivalent points
'  in the normal-sized image; calculate navigation rectangle coordinates and dimensions.
   NavXY.Left = CLng(CurrentXY.Left / m_PicWidth * AspRatWidth) + AspRatXY.X
   NavXY.Top = CLng(CurrentXY.Top / m_PicHeight * AspRatHeight) + AspRatXY.Y
   NavXY.Right = CLng(CurrentXY.Right / m_PicWidth * AspRatWidth) + AspRatXY.X
   NavXY.Bottom = CLng(CurrentXY.Bottom / m_PicHeight * AspRatHeight) + AspRatXY.Y
   NavWidth = NavXY.Right - NavXY.Left + 1
   NavHeight = NavXY.Bottom - NavXY.Top + 1

'  create and apply the color brush.
   hBrush = CreateSolidBrush(TranslateColor(m_NavRectColor))

'  draw the rectangle frame and delete the color brush.
   FrameRect hdc, NavXY, hBrush
   DeleteObject hBrush

End Sub

Private Sub ProcessImage()

'*************************************************************************
'* prepares a new image for display.                                     *
'*************************************************************************

   If Not IsPictureThere(m_Picture) Then Exit Sub

'  retrieve image statistics.
   GetImageInfo

'  create a virtual bitmap the height and width of the image (in pixels).
   CreateVirtualDC hdc, VirtualDC_Image, mMemoryBitmap_Image, mOriginalBitmap_Image, m_PicWidth, m_PicHeight

'  transfer the image in the .Picture property to the virtual bitmap.
'  Thanks to LaVolpe for his "FYI: StdPicture.Render Function" submission at txtCodeId=58041.
   m_Picture.Render VirtualDC_Image + 0, 0, 0, m_PicWidth + 0, m_PicHeight + 0, 0, ScaleY(m_PicHeight, vbPixels, vbHimetric), _
                    ScaleX(m_PicWidth, vbPixels, vbHimetric), -ScaleY(m_PicHeight, vbPixels, vbHimetric), ByVal 0&

'  get aspect ratio dimensions so image can be displayed stretched or compressed if indicated.
   If m_KeepAspectRatio Then
      AspRatHeight = IIf(DisplayAreaWidth < (m_PicWidth * (DisplayAreaHeight / m_PicHeight)), (m_PicHeight * (DisplayAreaWidth / m_PicWidth)), DisplayAreaHeight)
      AspRatWidth = IIf(DisplayAreaWidth < (m_PicWidth * (DisplayAreaHeight / m_PicHeight)), DisplayAreaWidth, (m_PicWidth * (DisplayAreaHeight / m_PicHeight)))
      AspRatXY.X = (ScaleWidth \ 2) - (AspRatWidth \ 2)
      AspRatXY.Y = (ScaleHeight \ 2) - (AspRatHeight \ 2)
   Else
      AspRatHeight = DisplayAreaHeight
      AspRatWidth = DisplayAreaWidth
      AspRatXY.X = m_BorderWidth
      AspRatXY.Y = m_BorderWidth
   End If

End Sub

Private Sub GetImageInfo()

'*************************************************************************
'* obtains various statistics pertaining to image display.               *
'*************************************************************************

'  the width and height of the image in pixels.
   m_PicWidth = IIf(IsPictureThere(m_Picture) = False, -1, ScaleX(m_Picture.Width, vbHimetric, vbPixels))
   m_PicHeight = IIf(IsPictureThere(m_Picture) = False, -1, ScaleX(m_Picture.Height, vbHimetric, vbPixels))
   
'  the width and the height of the display area in pixels.
   DisplayAreaWidth = ScaleWidth - m_BorderWidth * 2
   DisplayAreaHeight = ScaleHeight - m_BorderWidth * 2

'  X,Y of top left part of display area (Normal view) that will have image displayed.
   DisplayXY.X = IIf(m_PicWidth < DisplayAreaWidth, (ScaleWidth \ 2) - (m_PicWidth \ 2), m_BorderWidth)
   DisplayXY.Y = IIf(m_PicHeight < DisplayAreaHeight, (ScaleHeight \ 2) - (m_PicHeight \ 2), m_BorderWidth)

'  where we are in the overall image.
   If CurrentXY.Right = 0 And CurrentXY.Bottom = 0 Then
      CurrentXY.Left = 0    ' the x coordinate of the top left corner of the image always is 0.
      CurrentXY.Top = 0     ' the y coordinate of the top left corner of the image is always 0.
'     the x coordinate of the bottom right corner.
      CurrentXY.Right = IIf(m_PicWidth < DisplayAreaWidth, m_PicWidth - 1, DisplayAreaWidth)
'     the y coordinate of the bottom right corner.
      CurrentXY.Bottom = IIf(m_PicHeight < DisplayAreaHeight, m_PicHeight - 1, DisplayAreaHeight)
   End If

End Sub

Private Sub CalculateBackgroundGradient()

'*************************************************************************
'* generates gradient array/structure info and paints to virtual DC.     *
'*************************************************************************

'  generate and store the background gradient.
   CalculateGradient ScaleWidth, ScaleHeight, TranslateColor(ActiveBGColor1), TranslateColor(ActiveBGColor2), m_BackAngle, m_BackMiddleOut, uBIH_BG, lBits_BG()

'  paint the gradient onto the virtual DC bitmap.
   Call StretchDIBits(VirtualDC_Image, 0, 0, ScaleWidth, ScaleHeight, 0, 0, ScaleWidth, ScaleHeight, lBits_BG(0), uBIH_BG, DIB_RGB_COLORS, vbSrcCopy)

End Sub

Private Sub InitializeBorderGraphics()

'*************************************************************************
'* master routine for creating trapezoidal border segments.              *
'*************************************************************************

'  initialize border segment regions.
   DeleteBorderSegmentObjects    ' make sure they haven't already been created.
   BorderSegment(TOP_SEGMENT) = CreateDiagRectRegion(ScaleWidth, m_BorderWidth, 1, 1)
   BorderSegment(BOTTOM_SEGMENT) = CreateDiagRectRegion(ScaleWidth, m_BorderWidth, -1, -1)
   BorderSegment(RIGHT_SEGMENT) = CreateDiagRectRegion(m_BorderWidth, ScaleHeight, -1, -1)
   BorderSegment(LEFT_SEGMENT) = CreateDiagRectRegion(m_BorderWidth, ScaleHeight, 1, 1)

   InitializeBorderGradients

End Sub

Private Sub InitializeBorderGradients()

'*************************************************************************
'* creates virtual bitmaps and gradients for border segments.            *
'*************************************************************************

'  create the horizontal border segment virtual DC.
   CreateVirtualDC hdc, VirtualDC_SegH, mMemoryBitmap_SegH, mOriginalBitmap_SegH, ScaleWidth + 1, m_BorderWidth
'  create the vertical border segment virtual DC.
   CreateVirtualDC hdc, VirtualDC_SegV, mMemoryBitmap_SegV, mOriginalBitmap_SegV, m_BorderWidth, ScaleHeight
'  calculate the border gradients.
   CalculateBorderGradients

End Sub

Private Sub CalculateBorderGradients()

'*************************************************************************
'* calculates border gradients, accounting for middle-out status.        *
'*************************************************************************

'  calculate the primary horizontal segment gradient.
   CalculateGradient ScaleWidth, m_BorderWidth + 1, TranslateColor(ActiveBorderColor1), TranslateColor(ActiveBorderColor2), 90, m_BorderMiddleOut, SegH1uBIH, SegH1lBits()

'  if gradients are not middle-out, calculate the secondary horizontal segment gradient.
   If Not m_BorderMiddleOut Then CalculateGradient ScaleWidth, m_BorderWidth + 1, TranslateColor(ActiveBorderColor2), TranslateColor(ActiveBorderColor1), 90, m_BorderMiddleOut, SegH2uBIH, SegH2lBits()

'  calculate the primary vertical segment gradient.
   CalculateGradient m_BorderWidth + 1, ScaleHeight, TranslateColor(ActiveBorderColor1), TranslateColor(ActiveBorderColor2), 180, m_BorderMiddleOut, SegV1uBIH, SegV1lBits()

'  if gradients are not middle-out, calculate the secondary vertical segment gradient.
   If Not m_BorderMiddleOut Then CalculateGradient m_BorderWidth + 1, ScaleHeight, TranslateColor(ActiveBorderColor2), TranslateColor(ActiveBorderColor1), 180, m_BorderMiddleOut, SegV2uBIH, SegV2lBits()

End Sub

Private Sub GetActiveBorderColors()

'*************************************************************************
'* determines border colors based on current state of control.           *
'*************************************************************************

   ActiveBorderColor1 = IIf(m_Enabled = True, IIf(HasFocus = True, m_FocusBorderColor1, m_BorderColor1), m_DisBorderBGColor1)
   ActiveBorderColor2 = IIf(m_Enabled = True, IIf(HasFocus = True, m_FocusBorderColor2, m_BorderColor2), m_DisBorderBGColor2)

   CalculateBorderGradients

End Sub

Private Sub GetActiveBackgroundColors()

'*************************************************************************
'* determines background colors based on enabled status of control.      *
'*************************************************************************

   ActiveBGColor1 = IIf(m_Enabled = True, m_BackColor1, m_DisBorderBGColor1)
   ActiveBGColor2 = IIf(m_Enabled = True, m_BackColor2, m_DisBorderBGColor2)

'  recalculate the background gradient.
   CalculateGradient ScaleWidth, ScaleHeight, TranslateColor(ActiveBGColor1), TranslateColor(ActiveBGColor2), m_BackAngle, m_BackMiddleOut, uBIH_BG, lBits_BG()

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Properties >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_InitProperties()

'*************************************************************************
'* initialize properties to default values.                              *
'*************************************************************************

   Set m_Picture = LoadPicture("")
   m_BackAngle = 60
   m_BackColor1 = &H0
   m_BackColor2 = &HA0A0A0
   m_BackMiddleOut = True
   m_BorderColor1 = &H0
   m_BorderColor2 = &HE0E0E0
   m_BorderMiddleOut = True
   m_BorderWidth = 10
   m_DisBorderBGColor1 = &H808080
   m_DisBorderBGColor2 = &HE0E0E0
   m_Enabled = True
   m_FocusBorderColor1 = &H0
   m_FocusBorderColor2 = &HA0A0A0
   m_KeepAspectRatio = True
   m_NavModeEnabled = True
   m_NavRectColor = &HFFFF&
   m_PicHeight = 0
   m_PictureMode = 0
   m_PicWidth = 0

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*************************************************************************
'* read properties in the property bag.                                  *
'*************************************************************************

   With PropBag
      Set m_Picture = .ReadProperty("Picture", Nothing)
      m_BackAngle = .ReadProperty("BackAngle", 60)
      m_BackColor1 = .ReadProperty("BackColor1", &H0)
      m_BackColor2 = .ReadProperty("BackColor2", &HA0A0A0)
      m_BackMiddleOut = .ReadProperty("BackMiddleOut", True)
      m_BorderColor1 = .ReadProperty("BorderColor1", &H0)
      m_BorderColor2 = .ReadProperty("BorderColor2", &HE0E0E0)
      m_BorderMiddleOut = .ReadProperty("BorderMiddleOut", True)
      m_BorderWidth = .ReadProperty("BorderWidth", 10)
      m_DisBorderBGColor1 = .ReadProperty("DisBorderBGColor1", &H808080)
      m_DisBorderBGColor2 = .ReadProperty("DisBorderBGColor2", &HE0E0E0)
      m_Enabled = .ReadProperty("Enabled", True)
      m_FocusBorderColor1 = .ReadProperty("FocusBorderColor1", &H0)
      m_FocusBorderColor2 = .ReadProperty("FocusBorderColor2", &HA0A0A0)
      m_KeepAspectRatio = .ReadProperty("KeepAspectRatio", True)
      m_NavModeEnabled = .ReadProperty("NavModeEnabled", True)
      m_NavRectColor = .ReadProperty("NavRectColor", &HFFFF&)
      m_PicHeight = .ReadProperty("PicHeight", 0)
      m_PictureMode = .ReadProperty("PictureMode", 0)
      m_PicWidth = .ReadProperty("PicWidth", 0)
   End With

   InitializeControl

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'*************************************************************************
'* write the properties in the property bag.                             *
'*************************************************************************

   With PropBag
      .WriteProperty "BackAngle", m_BackAngle, 60
      .WriteProperty "BackColor1", m_BackColor1, &H0
      .WriteProperty "BackColor2", m_BackColor2, &HA0A0A0
      .WriteProperty "BackMiddleOut", m_BackMiddleOut, True
      .WriteProperty "BorderColor1", m_BorderColor1, &H0
      .WriteProperty "BorderColor2", m_BorderColor2, &HE0E0E0
      .WriteProperty "BorderMiddleOut", m_BorderMiddleOut, True
      .WriteProperty "BorderWidth", m_BorderWidth, 10
      .WriteProperty "DisBorderBGColor1", m_DisBorderBGColor1, &H0
      .WriteProperty "DisBorderBGColor2", m_DisBorderBGColor2, &HA0A0A0
      .WriteProperty "Enabled", m_Enabled, True
      .WriteProperty "FocusBorderColor1", m_FocusBorderColor1, &H0
      .WriteProperty "FocusBorderColor2", m_FocusBorderColor2, &HA0A0A0
      .WriteProperty "KeepAspectRatio", m_KeepAspectRatio, True
      .WriteProperty "NavModeEnabled", m_NavModeEnabled, True
      .WriteProperty "NavRectColor", m_NavRectColor, &HFFFF&
      .WriteProperty "PicHeight", m_PicHeight, 0
      .WriteProperty "Picture", m_Picture, Nothing
      .WriteProperty "PictureMode", m_PictureMode, 0
      .WriteProperty "PicWidth", m_PicWidth, 0
   End With

End Sub

Public Property Get BackAngle() As Single
Attribute BackAngle.VB_Description = "The angle, in degrees, of the background gradient."
Attribute BackAngle.VB_ProcData.VB_Invoke_Property = ";BG Graphics"
   BackAngle = m_BackAngle
End Property

Public Property Let BackAngle(ByVal New_BackAngle As Single)
'  do some bounds checking.
   If New_BackAngle > 360 Then New_BackAngle = 360
   If New_BackAngle < 0 Then New_BackAngle = 0
   m_BackAngle = New_BackAngle
   CalculateBackgroundGradient
   RedrawControl
   PropertyChanged "BackAngle"
End Property

Public Property Get BackColor1() As OLE_COLOR
Attribute BackColor1.VB_Description = "The first color of the control background gradient."
Attribute BackColor1.VB_ProcData.VB_Invoke_Property = ";BG Graphics"
   BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
   m_BackColor1 = New_BackColor1
   CalculateBackgroundGradient
   RedrawControl
   PropertyChanged "BackColor1"
End Property

Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "The second color of the control background gradient."
Attribute BackColor2.VB_ProcData.VB_Invoke_Property = ";BG Graphics"
   BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
   m_BackColor2 = New_BackColor2
   CalculateBackgroundGradient
   RedrawControl
   PropertyChanged "BackColor2"
End Property

Public Property Get BackMiddleOut() As Boolean
Attribute BackMiddleOut.VB_Description = "If True, the control background is drawn in middle-out mode (Color1>Color2>Color1)."
Attribute BackMiddleOut.VB_ProcData.VB_Invoke_Property = ";BG Graphics"
   BackMiddleOut = m_BackMiddleOut
End Property

Public Property Let BackMiddleOut(ByVal New_BackMiddleOut As Boolean)
   m_BackMiddleOut = New_BackMiddleOut
   CalculateBackgroundGradient
   RedrawControl
   PropertyChanged "BackMiddleOut"
End Property

Public Property Get BorderColor1() As OLE_COLOR
Attribute BorderColor1.VB_Description = "The first color of the control's border. "
Attribute BorderColor1.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BorderColor1 = m_BorderColor1
End Property

Public Property Let BorderColor1(ByVal New_BorderColor1 As OLE_COLOR)
   m_BorderColor1 = New_BorderColor1
   GetActiveBorderColors
   RedrawControl
   PropertyChanged "BorderColor1"
End Property

Public Property Get BorderColor2() As OLE_COLOR
Attribute BorderColor2.VB_Description = "The second color of the control's border."
   BorderColor2 = m_BorderColor2
End Property

Public Property Let BorderColor2(ByVal New_BorderColor2 As OLE_COLOR)
   m_BorderColor2 = New_BorderColor2
   GetActiveBorderColors
   RedrawControl
   PropertyChanged "BorderColor2"
End Property

Public Property Get BorderMiddleOut() As Boolean
Attribute BorderMiddleOut.VB_Description = "If True, the control border is drawn in middle-out mode (Color1>Color2>Color1)."
   BorderMiddleOut = m_BorderMiddleOut
End Property

Public Property Let BorderMiddleOut(ByVal New_BorderMiddleOut As Boolean)
   m_BorderMiddleOut = New_BorderMiddleOut
   CalculateBorderGradients
   RedrawControl
   PropertyChanged "BorderMiddleOut"
End Property

Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "The width, in pixels, of the control's border."
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = ";Main Graphics"
   BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
   m_BorderWidth = New_BorderWidth
   InitializeBorderGraphics
   GetImageInfo
   RedrawControl
   PropertyChanged "BorderWidth"
End Property

Public Property Get DisBorderBGColor1() As OLE_COLOR
Attribute DisBorderBGColor1.VB_Description = "First gradient color of the border AND background when the control is disabled."
   DisBorderBGColor1 = m_DisBorderBGColor1
End Property

Public Property Get ColorAtCursor() As Long
    ColorAtCursor = m_ColorAtCursor
End Property

Public Property Let DisBorderBGColor1(ByVal New_DisBorderBGColor1 As OLE_COLOR)
   m_DisBorderBGColor1 = New_DisBorderBGColor1
   PropertyChanged "DisBorderBGColor1"
End Property

Public Property Get DisBorderBGColor2() As OLE_COLOR
Attribute DisBorderBGColor2.VB_Description = "Second gradient color of the border AND background when the control is disabled."
   DisBorderBGColor2 = m_DisBorderBGColor2
End Property

Public Property Let DisBorderBGColor2(ByVal New_DisBorderBGColor2 As OLE_COLOR)
   m_DisBorderBGColor2 = New_DisBorderBGColor2
   PropertyChanged "DisBorderBGColor2"
End Property

Public Property Get DispX1() As Long
Attribute DispX1.VB_Description = "The X coordinate of the top left corner of the display area where the image is displayed."
   DispX1 = m_DispX1
End Property

Public Property Get DispY1() As Long
Attribute DispY1.VB_Description = "The Y coordinate of the top left corner of the display area where the image is displayed."
   DispY1 = m_DispY1
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "If True, the control is active and can be used normally."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   If Not m_Enabled Then UserControl.MousePointer = vbDefault ' restore normal pointer if disabled.
   GetActiveBorderColors
   GetActiveBackgroundColors
   RedrawControl
   PropertyChanged "Enabled"
End Property

Public Property Get FocusBorderColor1() As OLE_COLOR
Attribute FocusBorderColor1.VB_Description = "The first gradient color of the border when the control has the focus."
   FocusBorderColor1 = m_FocusBorderColor1
End Property

Public Property Let FocusBorderColor1(ByVal New_FocusBorderColor1 As OLE_COLOR)
   m_FocusBorderColor1 = New_FocusBorderColor1
   PropertyChanged "FocusBorderColor1"
End Property

Public Property Get FocusBorderColor2() As OLE_COLOR
Attribute FocusBorderColor2.VB_Description = "The second gradient color of the border when the control has the focus."
   FocusBorderColor2 = m_FocusBorderColor2
End Property

Public Property Let FocusBorderColor2(ByVal New_FocusBorderColor2 As OLE_COLOR)
   m_FocusBorderColor2 = New_FocusBorderColor2
   PropertyChanged "FocusBorderColor2"
End Property

Public Property Get hdc() As Long
   hdc = UserControl.hdc
End Property

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get KeepAspectRatio() As Boolean
Attribute KeepAspectRatio.VB_Description = "If True, proper width and height proportions are maintained when stretching the image to be fully visible in the display area.  Also applies to ClickNavigation display."
   KeepAspectRatio = m_KeepAspectRatio
End Property

Public Property Let KeepAspectRatio(ByVal New_KeepAspectRatio As Boolean)
   m_KeepAspectRatio = New_KeepAspectRatio
   ProcessImage
   RedrawControl
   PropertyChanged "KeepAspectRatio"
End Property

Public Property Get NavModeEnabled() As Boolean
Attribute NavModeEnabled.VB_Description = "If True, keyboard, mouse and software navigation of large images is allowed.  If False, navigation via code is still permitted."
   NavModeEnabled = m_NavModeEnabled
End Property

Public Property Let NavModeEnabled(ByVal New_NavModeEnabled As Boolean)
   m_NavModeEnabled = New_NavModeEnabled
   PropertyChanged "NavModeEnabled"
End Property

Public Property Get NavRectColor() As OLE_COLOR
Attribute NavRectColor.VB_Description = "The color of the navigation rectangle that is displayed in ClickNavigation mode."
   NavRectColor = m_NavRectColor
End Property

Public Property Let NavRectColor(ByVal New_NavRectColor As OLE_COLOR)
   m_NavRectColor = New_NavRectColor
   PropertyChanged "NavRectColor"
End Property

Public Property Get PicHeight() As Long
Attribute PicHeight.VB_Description = "The height, in pixels, of the image currently being displayed."
   PicHeight = m_PicHeight
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "The image currently being displayed."
   Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
   Set m_Picture = New_Picture
   ProcessImage
   RedrawControl
   PropertyChanged "Picture"
End Property

Public Property Get PictureMode() As MPV_PicModeOptions
Attribute PictureMode.VB_Description = "How to display the image (normal size or stretched)."
   PictureMode = m_PictureMode
End Property

Public Property Let PictureMode(ByVal New_PictureMode As MPV_PicModeOptions)
   m_PictureMode = New_PictureMode
   RedrawControl
   PropertyChanged "PictureMode"
End Property

Public Property Get PicWidth() As Long
Attribute PicWidth.VB_Description = "The width, in pixels, of the image currently being displayed."
   PicWidth = m_PicWidth
End Property

Public Property Get PicX() As Long
Attribute PicX.VB_Description = "Returns the X coordinate of where in the image (NOT the display area coordinate) the mouse pointer is.  If not over the image, returns -1."
   PicX = m_PicX
End Property

Public Property Get PicY() As Long
Attribute PicY.VB_Description = "Returns the Y coordinate of where in the image (NOT the display area coordinate) the mouse pointer is.  If not over the image, returns -1."
   PicY = m_PicY
End Property
