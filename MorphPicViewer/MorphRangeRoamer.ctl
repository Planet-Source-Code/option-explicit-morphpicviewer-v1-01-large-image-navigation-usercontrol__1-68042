VERSION 5.00
Begin VB.UserControl MorphRangeRoamer 
   AutoRedraw      =   -1  'True
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   ScaleHeight     =   54
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   30
   ToolboxBitmap   =   "MorphRangeRoamer.ctx":0000
End
Attribute VB_Name = "MorphRangeRoamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************
'* MorphRangeRoamer 1.02 - --->VB6<--- UpDown/Slider hybrid usercontrol. *
'* Author: Matthew R. Usner, Dec. 2006 for www.planet-source-code.com.   *
'* Last update 18 Feb 2007 - Added .UD_SwapDirections property.          *
'* The most up-to-date version of this control can always be found at:   *
'* www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=67526&lngWId=1       *
'* Copyright ©2006 - 2007, Matthew R. Usner.  All rights reserved.       *
'*************************************************************************
'* MorphRangeRoamer is my attempt to overcome the inadequacies of the    *
'* humble UpDown and Slider controls.  To a reasonable extent, I believe *
'* I succeeded.  I like the concept of the UpDown control but rarely use *
'* it.  Why?  Because it's realistically only useful for selecting among *
'* a small range of values, otherwise you're stuck clicking down on one  *
'* of the direction buttons forever waiting for the correct value.  Auto-*
'* matic value acceleration via the .Increment property helps, but not a *
'* great deal.  Usually I end up overshooting my target by a gazillion   *
'* and have to go back. The Slider, on the other hand, is good for huge  *
'* ranges, but only if it takes up a lot of screen space.  Tweaking it   *
'* to narrow it down to a specific value is a pain. MorphRangeRoamer is  *
'* an UpDown/Slider hybrid that tries to improve the efficiency of both  *
'* controls by having them seamlessly work together.  Both small and     *
'* large ranges of values are handled easily. At first, MorphRangeRoamer *
'* acts much like a typical VB UpDown.  Click the up or down buttons to  *
'* increase or decrease the .Value property. Click and hold on one of    *
'* the direction buttons and the .Value will increment or decrement at   *
'* definable rates of speed.  You can set properties that control the    *
'* size of the value increment when the Ctrl, Shift, or both keys are    *
'* held down while clicking the UpDown buttons.  For example, the default*
'* increment could be 1, the increment while holding Ctrl could be 10,   *
'* while holding Shift 100, and Ctrl-Shift 1000.  This helps you navigate*
'* through large ranges much more efficiently than a regular UpDown. For *
'* larger ranges, after a definable time interval clicking and holding   *
'* down an UpDown button, a Slider-esque "RangeWindow" appears.  By keep-*
'* ing the mouse button held down, and moving the mouse right or left in *
'* the progress bar, you can instantly move anywhere in the value range. *
'* The progress bar and LED display in the RangeWindow let you keep      *
'* track of where you are in the value range you have defined in the     *
'* appropriate properties.  When you're close to (or at) the value you   *
'* want, release the mouse button.  The RangeWindow disappears and you   *
'* can click the up or down button a few times (using Shift and/or Ctrl  *
'* if necessary) until precisely the right value is chosen.  In this     *
'* fashion, you can use MorphRangeRoamer to very quickly navigate        *
'* through much larger value ranges than the intrinsic UpDown is capable *
'* of efficiently doing, with more precision than is convenient with the *
'* Slider.  There are some limitations; I am not trying to imply this    *
'* control is suitable for ranges in the billions, it's not.  Pixels     *
'* (RangeWindow operation) and ease of keyboard use (UpDown) prevent     *
'* that.  And as this is an early version, there is no BuddyControl      *
'* capability.  I may or may not add it; for now you'll have to emulate  *
'* BuddyControl functionality in your project's form code (which is what *
'* many coders do anyway).  But when you try this you'll see how much it *
'* outperforms the VB UpDown and Slider controls in many respects!       *
'*************************************************************************
'*    CONSTRUCTIVE FEEDBACK ALWAYS WELCOME, VOTES ALWAYS APPRECIATED!    *
'*************************************************************************
'* Legal:  Redistribution of this code, whole or in part, as source code *
'* or in binary form, alone or as part of a larger distribution or prod- *
'* uct, is forbidden for any commercial or for-profit use without the    *
'* author's explicit written permission.                                 *
'*                                                                       *
'* Non-commercial redistribution of this code, as source code or in      *
'* binary form, with or without modification, is permitted provided that *
'* the following conditions are met:                                     *
'*                                                                       *
'* Redistributions of source code must include this list of conditions,  *
'* and the following acknowledgment:                                     *
'*                                                                       *
'* This VB6 usercontrol was developed by Matthew R. Usner.               *
'* Source code, written in Visual Basic 6.0, is freely available for     *
'* non-commercial, non-profit use.                                       *
'*                                                                       *
'* Redistributions in binary form, as part of a larger project, must     *
'* include the above acknowledgment in the end-user documentation.       *
'* Alternatively, the above acknowledgment may appear in the software    *
'* itself, if and where such third-party acknowledgments normally appear.*
'*************************************************************************
'* Credits and Thanks:                                                   *
'* Carles P.V., for the gradient generation code and CreateWindowEx tip. *
'* LaVolpe, for the border segment region generation code.               *
'*************************************************************************

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32.dll" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

' enum tied to .UD_Orientation property.
Public Enum MRR_Orientation
   Vertical
   Horizontal
End Enum

' enum tied to .Theme property.
Public Enum MRR_ThemeOptions
   [None] = 0
   [Cyan Eyed] = 1
   [Gunmetal Grey] = 2
   [Blue Moon] = 3
   [Red Rum] = 4
   [Green With Envy] = 5
   [Purple People Eater] = 6
   [Golden Goose] = 7
   [Penny Wise] = 8
End Enum

' ********** updown graphics declares. *********

' declares for gradient painting and bitmap tiling.
Private Type BITMAPINFOHEADER
   biSize                                     As Long
   biWidth                                    As Long
   biHeight                                   As Long
   biPlanes                                   As Integer
   biBitCount                                 As Integer
   biCompression                              As Long
   biSizeImage                                As Long
   biXPelsPerMeter                            As Long
   biYPelsPerMeter                            As Long
   biClrUsed                                  As Long
   biClrImportant                             As Long
End Type

Private Type POINTAPI
   x                                          As Long
   y                                          As Long
End Type
Private CursorPos As POINTAPI                                      ' RangeWindow cursor XY for progressbar.

Private Const DIB_RGB_COLORS                  As Long = 0          ' used in gradient generation and transfer.

'  gradient generation constants.
Private Const PI                              As Single = 3.14159265358979
Private Const TO_DEG                          As Single = 180 / PI
Private Const TO_RAD                          As Single = PI / 180
Private Const INT_ROT                         As Long = 1000

'  used to define various graphics areas and updown component locations.
Private Type RECT
   Left                                       As Long
   Top                                        As Long
   Right                                      As Long
   Bottom                                     As Long
End Type

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< UpDown Button declares >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' 4 virtual bitmaps created - 2 for left/top button, 2 for right/bottom button.  For each
' button one bitmap holds graphics for button in up state, the other for button in down state.
' ******************* top/left button *******************************

Private ButtonWidth                           As Long              ' width, in pixels, of UpDown button.
Private ButtonHeight                          As Long              ' height, in pixels, of UpDown button.

' coordinates of UpDown buttons.
Private Type ButtonCoordinates
   X1                                         As Long
   Y1                                         As Long
   X2                                         As Long
   Y2                                         As Long
End Type
Private ButtonCoords(1 To 2)                  As ButtonCoordinates

' mouse location constants.
Private Const LEFT_OR_TOP_BUTTON              As Long = 1
Private Const RIGHT_OR_BOTTOM_BUTTON          As Long = 2
Private Const MOUSE_NOT_IN_BUTTON             As Long = 0
Private Const MOUSE_IN_LEFT_OR_TOP_BUTTON     As Long = 1
Private Const MOUSE_IN_RIGHT_OR_BOTTOM_BUTTON As Long = 2
Private MouseLocation                         As Long              ' assigned one of the above constants.

' declares for UpDown 'up' button virtual bitmap.
Private UD_VirtualDC_LT_ButtonUp              As Long              ' handle of the created DC.
Private UD_mMemoryBitmap_LT_ButtonUp          As Long              ' handle of the created bitmap.
Private UD_mOriginalBitmap_LT_ButtonUp        As Long              ' used in destroying virtual DC.

' declares for UpDown 'down' button virtual bitmap.
Private UD_VirtualDC_LT_ButtonDown            As Long              ' handle of the created DC.
Private UD_mMemoryBitmap_LT_ButtonDown        As Long              ' handle of the created bitmap.
Private UD_mOriginalBitmap_LT_ButtonDown      As Long              ' used in destroying virtual DC.

' ******************* bottom/right button *******************************
' declares for UpDown 'up' button virtual bitmap.
Private UD_VirtualDC_RB_ButtonUp              As Long              ' handle of the created DC.
Private UD_mMemoryBitmap_RB_ButtonUp          As Long              ' handle of the created bitmap.
Private UD_mOriginalBitmap_RB_ButtonUp        As Long              ' used in destroying virtual DC.

' declares for UpDown 'down' button virtual bitmap.
Private UD_VirtualDC_RB_ButtonDown            As Long              ' handle of the created DC.
Private UD_mMemoryBitmap_RB_ButtonDown        As Long              ' handle of the created bitmap.
Private UD_mOriginalBitmap_RB_ButtonDown      As Long              ' used in destroying virtual DC.

'  gradient information for updown button faces.
Private UD_ButtonUp_uBIH                      As BITMAPINFOHEADER
Private UD_ButtonUp_lBits()                   As Long
Private UD_ButtonDown_LT_uBIH                 As BITMAPINFOHEADER
Private UD_ButtonDown_LT_lBits()              As Long
Private UD_ButtonDown_RB_uBIH                 As BITMAPINFOHEADER
Private UD_ButtonDown_RB_lBits()              As Long
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

' constants defining the four border segments.
Private Const TOP_SEGMENT                     As Long = 0
Private Const RIGHT_SEGMENT                   As Long = 1
Private Const BOTTOM_SEGMENT                  As Long = 2
Private Const LEFT_SEGMENT                    As Long = 3

' ******************** UpDown border declares ************************
'  gradient information for UpDown horizontal and vertical border segments.
Private UD_SegV1uBIH                          As BITMAPINFOHEADER
Private UD_SegV1lBits()                       As Long
Private UD_SegV2uBIH                          As BITMAPINFOHEADER
Private UD_SegV2lBits()                       As Long
Private UD_SegH1uBIH                          As BITMAPINFOHEADER
Private UD_SegH1lBits()                       As Long
Private UD_SegH2uBIH                          As BITMAPINFOHEADER
Private UD_SegH2lBits()                       As Long

' holds region pointers for UpDown border segments.
Private UD_BorderSegment(0 To 3)              As Long

' declares for horizontal border segment virtual bitmap.
Private UD_VirtualDC_SegH                     As Long              ' handle of the created DC.
Private UD_mMemoryBitmap_SegH                 As Long              ' handle of the created bitmap.
Private UD_mOriginalBitmap_SegH               As Long              ' used in destroying virtual DC.

' declares for vertical border segment virtual bitmap.
Private UD_VirtualDC_SegV                     As Long              ' handle of the created DC.
Private UD_mMemoryBitmap_SegV                 As Long              ' handle of the created bitmap.
Private UD_mOriginalBitmap_SegV               As Long              ' used in destroying virtual DC.

Private CtrlKeyDown                           As Boolean           ' global "ctrl key being pressed" flag.
Private ShiftKeyDown                          As Boolean           ' global "shift key is being pressed" flag.
Private ShiftAndCtrlDown                      As Boolean           ' both Shift and Ctrl keys are down.

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<< RangeWindow declares >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private RangeWindowWidth                      As Long              ' width, in pixels, of the RangeWindow.
Private RangeWindowHeight                     As Long              ' height, in pixels, of the RangeWindow.
Private RW_BG_Width                           As Long              ' RangeWindow background bitmap width.
Private RW_BG_Height                          As Long              ' RangeWindow background bitmap height.
Private RW_X1                                 As Long              ' left coord of RangeWindow progressbar.
Private RW_X2                                 As Long              ' right coord of RangeWindow progressbar.
Private RW_Y1                                 As Long              ' top coordinate of RangeWindow progress dragbar.
Private RW_Y2                                 As Long              ' bottom coordinate of RangeWindow progress dragbar.

' declares that are used to make sure RangeWindow is always fully visible on screen (Y coordinate).
Private ScreenWorkArea                        As RECT              ' screen area without taskbar.
Private Const SPI_GETWORKAREA                 As Long = 48&        ' used in SystemParametersInfo API call.
Private TaskBarHeight                         As Long              ' for adjustment of RangeWindow Y coord.

' RangeWindow LED display declares.
Private LEDSegment(0 To 1)                    As Long              ' region pointers for LED digit segments.
Private Const VERTICAL_LED_SEGMENT            As Long = 0          ' vertical LED segment region index.
Private Const HORIZONTAL_LED_SEGMENT          As Long = 1          ' vertical LED segment region index.
Private Const SEGMENT_LIT                     As String = "1"      ' segment lit constant.
Private Const SEGMENT_UNLIT                   As String = "0"      ' segment unlit constant.
Private Const SegmentWidth                    As Long = 2          ' LED digit segment width, in pixels.
Private Const SegmentHeight                   As Long = 7          ' LED digit segment height, in pixels.
Private LEDLitColorBrush                      As Long              ' color brush for lit LED segments.
Private LEDBurnInColorBrush                   As Long              ' color brush for 'burned in' LED segments.
Private DisplayPattern()                      As String            ' LED digit segment display patterns.
Private Const InterSegmentGap                 As Long = 1          ' number of pixels between LED digit segments.
Private DigitHeight                           As Long              ' pixel height of LED digit.
Private DigitWidth                            As Long              ' pixel width of LED digit.
Private DigitXPos(0 To 9)                     As Long              ' X coordinate of each value LED digit.
Private Const InterDigitGap                   As Long = 6          ' number of pixels between each LED digit.
Private PreviousValue                         As String            ' helps in displaying only changed LED digits.
Private Const MAX_DIGITS                      As Long = 10         ' # of possible LED digits (including minus sign).

' declares for RangeWindow virtual bitmap.
Private RW_VirtualDC                          As Long              ' handle of the created DC.
Private RW_mMemoryBitmap                      As Long              ' handle of the created bitmap.
Private RW_mOriginalBitmap                    As Long              ' used in destroying virtual DC.

' declares for RangeWindow LED display background virtual bitmap.
Private RW_BG_VirtualDC                       As Long              ' handle of the created DC.
Private RW_BG_mMemoryBitmap                   As Long              ' handle of the created bitmap.
Private RW_BG_mOriginalBitmap                 As Long              ' used in destroying virtual DC.

' gradient information for RangeWindow progress bar meter.
Private RW_Meter_uBIH                         As BITMAPINFOHEADER
Private RW_Meter_lBits()                      As Long

' gradient information for RangeWindow display background.
Private RW_BG_uBIH                            As BITMAPINFOHEADER
Private RW_BG_lBits()                         As Long

' RangeWindow CreateWindowEx constants.
Private Const SS_CUSTOMDRAW                   As Long = &HD
Private Const SW_SHOWNORMAL                   As Long = 1
Private Const WS_POPUP                        As Long = &H80000000
Private Const WS_EX_TOOLWINDOW                As Long = &H80&

' used in SetWindowPos API to ignore resizing and repositioning when setting RangeWindow ZOrder to topmost.
Private Const HWND_TOP                        As Long = 0
Private Const SWP_NOMOVE                      As Long = &H2
Private Const SWP_NOSIZE                      As Long = &H1

Private RangeWindowPopped                     As Boolean           ' "RangeWindow activated" flag.
Private RWhWnd                                As Long              ' RangeWindow virtual window handle.
Private RWhDC                                 As Long              ' RangeWindow virtual window DC.

' ************************* RangeWindow border declares *************************
' gradient information for RangeWindow horizontal and vertical border segments.
Private RW_SegV1uBIH                          As BITMAPINFOHEADER
Private RW_SegV1lBits()                       As Long
Private RW_SegV2uBIH                          As BITMAPINFOHEADER
Private RW_SegV2lBits()                       As Long
Private RW_SegH1uBIH                          As BITMAPINFOHEADER
Private RW_SegH1lBits()                       As Long
Private RW_SegH2uBIH                          As BITMAPINFOHEADER
Private RW_SegH2lBits()                       As Long

' holds region pointers for RangeWindow border segments.
Private RW_BorderSegment(0 To 3)              As Long

' declares for horizontal border segment virtual bitmap.
Private RW_VirtualDC_SegH                     As Long              ' handle of the created DC.
Private RW_mMemoryBitmap_SegH                 As Long              ' handle of the created bitmap.
Private RW_mOriginalBitmap_SegH               As Long              ' used in destroying virtual DC.

' declares for vertical border segment virtual bitmap.
Private RW_VirtualDC_SegV                     As Long              ' handle of the created DC.
Private RW_mMemoryBitmap_SegV                 As Long              ' handle of the created bitmap.
Private RW_mOriginalBitmap_SegV               As Long              ' used in destroying virtual DC.
'********************************************************************************

' default property values.  They correspond to the "Gunmetal Grey" theme.
Private Const m_def_Enabled = True
Private Const m_def_RW_BackAngle = 90
Private Const m_def_RW_BackColor1 = 0
Private Const m_def_RW_BackColor2 = &H404040
Private Const m_def_RW_BackMiddleOut = True
Private Const m_def_RW_BorderColor1 = 0
Private Const m_def_RW_BorderColor2 = &H808080
Private Const m_def_RW_BorderMiddleOut = True
Private Const m_def_RW_BorderWidth = 8
Private Const m_def_RW_GenerateEvent = True
Private Const m_def_RW_LED_BurnInColor = &H404040
Private Const m_def_RW_LED_DigitColor = &HE0E0E0
Private Const m_def_RW_LED_ShowBurnIn = True
Private Const m_def_RW_PBarColor1 = &H0
Private Const m_def_RW_PBarColor2 = &HE0E0E0
Private Const m_def_RW_PopInterval = 3000
Private Const m_def_RW_ShowLED = True
Private Const m_def_Theme = 2
Private Const m_def_UD_ArrowColor = &HFFFFFF
Private Const m_def_UD_AutoIncrement = False
Private Const m_def_UD_BorderColor1 = &H0
Private Const m_def_UD_BorderColor2 = &HC0C0C0
Private Const m_def_UD_BorderMiddleOut = True
Private Const m_def_UD_BorderWidth = 8
Private Const m_def_UD_ButtonColor1 = &H0
Private Const m_def_UD_ButtonColor2 = &HC0C0C0
Private Const m_def_UD_ButtonDownAngle = 90
Private Const m_def_UD_ButtonDownMidOut = False
Private Const m_def_UD_ButtonUpAngle = 90
Private Const m_def_UD_ButtonUpMidOut = True
Private Const m_def_UD_DisArrowColor = &H909090
Private Const m_def_UD_DisBorderColor1 = &H808080
Private Const m_def_UD_DisBorderColor2 = &HE0E0E0
Private Const m_def_UD_DisButtonColor1 = &H808080
Private Const m_def_UD_DisButtonColor2 = &HE0E0E0
Private Const m_def_UD_FocusBorderColor1 = &H0
Private Const m_def_UD_FocusBorderColor2 = &H808080
Private Const m_def_UD_IncrementInterval = 250
Private Const m_def_UD_Orientation = [Vertical]
Private Const m_def_UD_ScrollDelay = 1000
Private Const m_def_UD_SwapDirections = False
Private Const m_def_Value = 0
Private Const m_def_ValueIncrCtrl = 10
Private Const m_def_ValueIncrement = 1
Private Const m_def_ValueIncrShift = 100
Private Const m_def_ValueIncrShiftCtrl = 1000
Private Const m_def_ValueMax = 100
Private Const m_def_ValueMin = 1
Private Const m_def_Wrap = False

' property variables.
Private m_Enabled                             As Boolean           ' master control enabled flag.
Private m_RW_BackAngle                        As Single            ' RangeWindow background gradient angle.
Private m_RW_BackColor1                       As OLE_COLOR         ' RangeWindow background gradient color 1.
Private m_RW_BackColor2                       As OLE_COLOR         ' RangeWindow background gradient color 2.
Private m_RW_BorderColor1                     As OLE_COLOR         ' RangeWindow border gradient color 1.
Private m_RW_BorderColor2                     As OLE_COLOR         ' RangeWindow border gradient color 2.
Private m_RW_BackMiddleOut                    As Boolean           ' RangeWindow background middle-out status.
Private m_RW_BorderMiddleOut                  As Boolean           ' RangeWindow border middle-out status.
Private m_RW_BorderWidth                      As Long              ' width, in pixels, of RangeWindow border.
Private m_RW_GenerateEvent                    As Boolean           ' Change event thrown in MouseMove? (RangeWindow only).
Private m_RW_LED_BurnInColor                  As OLE_COLOR         ' simulated LED digit burn-in color.
Private m_RW_LED_DigitColor                   As OLE_COLOR         ' LED digit segment color.
Private m_RW_LED_ShowBurnIn                   As Boolean           ' show simulated LED burn-in digits flag.
Private m_RW_PBarColor1                       As OLE_COLOR         ' gradient color 1 of RangeWindow progressbar.
Private m_RW_PBarColor2                       As OLE_COLOR         ' gradient color 2 of RangeWindow progressbar.
Private m_RW_PopInterval                      As Long              ' time (ms) before slider appears.
Private m_RW_ShowLED                          As Boolean           ' if True, RangeWindow LED display is shown.
Private m_Theme                               As MRR_ThemeOptions  ' color scheme for control.
Private m_UD_ArrowColor                       As OLE_COLOR         ' color of button arrows when buttons are up.
Private m_UD_AutoIncrement                    As Boolean           ' sets auto-calculation of various increments.
Private m_UD_BorderColor1                     As OLE_COLOR         ' first UpDown border gradient color.
Private m_UD_BorderColor2                     As OLE_COLOR         ' second UpDown border gradient color.
Private m_UD_BorderMiddleOut                  As Boolean           ' UpDown gradient border middle-out status.
Private m_UD_BorderWidth                      As Long              ' UpDown gradient border width, in pixels.
Private m_UD_ButtonColor1                     As OLE_COLOR         ' first gradient color of UD buttons (up position).
Private m_UD_ButtonColor2                     As OLE_COLOR         ' second gradient color of UD buttons (up position).
Private m_UD_ButtonDownAngle                  As Single            ' button down gradient angle.
Private m_UD_ButtonDownMidOut                 As Boolean           ' button down middle-out gradient status.
Private m_UD_ButtonUpAngle                    As Single            ' button up gradient angle.
Private m_UD_ButtonUpMidOut                   As Boolean           ' button up middle-out gradient status.
Private m_UD_DisArrowColor                    As OLE_COLOR         ' arrow color when control is disabled.
Private m_UD_DisBorderColor1                  As OLE_COLOR         ' border color 1 when control is disabled.
Private m_UD_DisBorderColor2                  As OLE_COLOR         ' border color 2 when control is disabled.
Private m_UD_DisButtonColor1                  As OLE_COLOR         ' button color 1 when control is disabled.
Private m_UD_DisButtonColor2                  As OLE_COLOR         ' button color 2 when control is disabled.
Private m_UD_FocusBorderColor1                As OLE_COLOR         ' UpDown border color 1 when control has focus.
Private m_UD_FocusBorderColor2                As OLE_COLOR         ' UpDown border color 2 when control has focus.
Private m_UD_IncrementInterval                As Long              ' delay, in ms, between .Value auto-increments.
Private m_UD_Orientation                      As MRR_Orientation   ' buttons in vertical or horizontal format.
Private m_UD_ScrollDelay                      As Long              ' delay (ms) before value scrolling begins.
Private m_UD_SwapDirections                   As Boolean           ' swaps increment/decrement if desired.
Private m_Value                               As Long              ' current UpDown value.
Private m_ValueIncrCtrl                       As Long              ' increment when Ctrl key is pressed.
Private m_ValueIncrement                      As Long              ' default increment (should be 1, usually).
Private m_ValueIncrShift                      As Long              ' increment when Shift key is pressed.
Private m_ValueIncrShiftCtrl                  As Long              ' increment when Shift and Ctrl keys are pressed.
Private m_ValueMin                            As Long              ' minimum value in UpDown range.
Private m_ValueMax                            As Long              ' maximum value in UpDown range.
Private m_Wrap                                As Boolean           ' value wraps back to min or max value flag.

' variables that display the enabled/disabled UpDown colors.
Private Active_UD_ArrowColor                  As OLE_COLOR         ' current UpDown button arrow color.
Private Active_UD_BorderColor1                As OLE_COLOR         ' current UpDown gradient border color 1.
Private Active_UD_BorderColor2                As OLE_COLOR         ' current UpDown gradient border color 2.
Private Active_UD_ButtonColor1                As OLE_COLOR         ' current UpDown gradient button color 1.
Private Active_UD_ButtonColor2                As OLE_COLOR         ' current UpDown gradient button color 2.
Private Active_UD_ButtonUpAngle               As Single            ' UpAngle used for both enabled/disabled.
Private Active_UD_ButtonUpMidOut              As Boolean           ' UpMidOut used for both enabled/disabled.

'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event Change()

' mouse control variables.
Private MouseX                                As Single            ' global mouse X position variable.
Private MouseY                                As Single            ' global mouse Y position variable.
Private Const MOUSEEVENTF_LEFTDOWN            As Long = &H2        ' for generating a mousedown to replace double click.
Private MouseIsDown                           As Boolean           ' left mouse button down flag.
Private OriginalMouseDownButton               As Long              ' which UpDown button mouse was clicked down in.
Private MoveRelease                           As Boolean           ' flag to unclick button when mouse moved out.
Private RightMouseButtonClicked               As Boolean           ' flag to help ignore right clicking.

Private Sub InitializeControlGraphics()

'*************************************************************************
'* master procedure for initializing all MorphRangeRoamer graphics.      *
'*************************************************************************

   InitializeUpDownGraphics
   InitializeRangeWindowGraphics

End Sub

Private Sub InitializeUpDownGraphics()

'*************************************************************************
'* master procedure for initializing all UpDown graphics.                *
'*************************************************************************

'  initialize appropriate display colors.
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If

   InitializeUpDownButtonGraphics
   InitializeUpDownBorderGraphics

End Sub

Private Sub RedrawControl()

'*************************************************************************
'* displays the UpDown portion of the control.                           *
'*************************************************************************

   DisplayButton LEFT_OR_TOP_BUTTON
   DisplayButton RIGHT_OR_BOTTOM_BUTTON
   DisplayUpDownBorder

End Sub

Private Sub InitializeUpDownButtonGraphics()

'*************************************************************************
'* initializes UpDown button graphics.                                   *
'*************************************************************************

   CalculateUpDownButtonDimensions
   CalculateUpDownButtonCoordinates
   CreateUpDownButtonVirtualBitmaps
   GenerateUpDownButtonGradients
   TransferButtonGradientsToVirtualBitmaps
   TransferArrowsToVirtualBitmaps

End Sub

Private Sub DisplayUpDownBorder()

'*************************************************************************
'* paints the UpDown gradient border.                                    *
'*************************************************************************

'  if the borderwidth is greater than 1 pixel, use the gradient border.
   If m_UD_BorderWidth > 1 Then

      DisplayBorderSegment UD_VirtualDC_SegV, hdc, ScaleWidth, ScaleHeight, m_UD_BorderWidth, _
                           UD_BorderSegment(), LEFT_SEGMENT, 0, 0, UD_SegV1uBIH, UD_SegV1lBits(), _
                           m_UD_BorderMiddleOut

      If m_UD_BorderMiddleOut Then
         DisplayBorderSegment UD_VirtualDC_SegV, hdc, ScaleWidth, ScaleHeight, m_UD_BorderWidth, _
                              UD_BorderSegment(), RIGHT_SEGMENT, ScaleWidth - m_UD_BorderWidth, 0, _
                              UD_SegV1uBIH, UD_SegV1lBits(), m_UD_BorderMiddleOut
      Else
         DisplayBorderSegment UD_VirtualDC_SegV, hdc, ScaleWidth, ScaleHeight, m_UD_BorderWidth, _
                              UD_BorderSegment(), RIGHT_SEGMENT, ScaleWidth - m_UD_BorderWidth, 0, _
                              UD_SegV2uBIH, UD_SegV2lBits(), m_UD_BorderMiddleOut
      End If

      DisplayBorderSegment UD_VirtualDC_SegH, hdc, ScaleWidth, ScaleHeight, m_UD_BorderWidth, _
                           UD_BorderSegment(), TOP_SEGMENT, 0, 0, UD_SegH1uBIH, UD_SegH1lBits(), _
                           m_UD_BorderMiddleOut

      If m_UD_BorderMiddleOut Then
         DisplayBorderSegment UD_VirtualDC_SegH, hdc, ScaleWidth, ScaleHeight, m_UD_BorderWidth, _
                              UD_BorderSegment(), BOTTOM_SEGMENT, -1, ScaleHeight - m_UD_BorderWidth, _
                              UD_SegH1uBIH, UD_SegH1lBits(), m_UD_BorderMiddleOut
      Else
         DisplayBorderSegment UD_VirtualDC_SegH, hdc, ScaleWidth, ScaleHeight, m_UD_BorderWidth, _
                              UD_BorderSegment(), BOTTOM_SEGMENT, -1, ScaleHeight - m_UD_BorderWidth, _
                              UD_SegH2uBIH, UD_SegH2lBits(), m_UD_BorderMiddleOut
      End If
   
   End If

End Sub

Private Sub DisplayRangeWindowBorder()

'*************************************************************************
'* paints the RangeWindow gradient border.                               *
'*************************************************************************

'  if the borderwidth is greater than 1 pixel, use the gradient border.
   If m_RW_BorderWidth > 1 Then

      DisplayBorderSegment RW_VirtualDC_SegV, RW_VirtualDC, RangeWindowWidth, RangeWindowHeight, _
                           m_RW_BorderWidth, RW_BorderSegment(), LEFT_SEGMENT, 0, 0, _
                           RW_SegV1uBIH, RW_SegV1lBits(), m_RW_BorderMiddleOut

      If m_RW_BorderMiddleOut Then
         DisplayBorderSegment RW_VirtualDC_SegV, RW_VirtualDC, RangeWindowWidth, RangeWindowHeight, _
                              m_RW_BorderWidth, RW_BorderSegment(), RIGHT_SEGMENT, RangeWindowWidth - m_RW_BorderWidth, _
                              0, RW_SegV1uBIH, RW_SegV1lBits(), m_RW_BorderMiddleOut
      Else
         DisplayBorderSegment RW_VirtualDC_SegV, RW_VirtualDC, RangeWindowWidth, RangeWindowHeight, _
                              m_RW_BorderWidth, RW_BorderSegment(), RIGHT_SEGMENT, RangeWindowWidth - m_RW_BorderWidth, _
                              0, RW_SegV2uBIH, RW_SegV2lBits(), m_RW_BorderMiddleOut
      End If

      DisplayBorderSegment RW_VirtualDC_SegH, RW_VirtualDC, RangeWindowWidth, RangeWindowHeight, _
                           m_RW_BorderWidth, RW_BorderSegment(), TOP_SEGMENT, 0, 0, _
                           RW_SegH1uBIH, RW_SegH1lBits(), m_RW_BorderMiddleOut

      If m_RW_BorderMiddleOut Then
         DisplayBorderSegment RW_VirtualDC_SegH, RW_VirtualDC, RangeWindowWidth, RangeWindowHeight, _
                              m_RW_BorderWidth, RW_BorderSegment(), BOTTOM_SEGMENT, -1, _
                              RangeWindowHeight - m_RW_BorderWidth, RW_SegH1uBIH, RW_SegH1lBits(), m_RW_BorderMiddleOut
      Else
         DisplayBorderSegment RW_VirtualDC_SegH, RW_VirtualDC, RangeWindowWidth, RangeWindowHeight, _
                              m_RW_BorderWidth, RW_BorderSegment(), BOTTOM_SEGMENT, -1, RangeWindowHeight - m_RW_BorderWidth, _
                              RW_SegH2uBIH, RW_SegH2lBits(), m_RW_BorderMiddleOut
      End If

   End If

End Sub

Private Sub DisplayBorderSegment(ByVal SourceDC As Long, ByVal TargetDC As Long, ByVal TargetWidth As Long, ByVal TargetHeight As Long, _
                                 ByVal BorderWidth As Long, ByRef SegArray() As Long, ByVal SegmentNdx As Long, _
                                 ByVal StartX As Long, ByVal StartY As Long, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long, _
                                 ByVal bMOut As Boolean)

'*************************************************************************
'* displays one border segment.  Border segment gradients are displayed  *
'* to virtual bitmaps on the fly so that correct gradient orientation    *
'* is maintained if the .MiddleOut property is set to False.             *
'*************************************************************************

'  position the border segment region in the correct location.
   MoveRegionToXY SegArray(SegmentNdx), StartX, StartY

   Select Case SegmentNdx

      Case LEFT_SEGMENT
         PaintVerticalBorderGradient SourceDC, BorderWidth, TargetHeight, uBIH, lBits()
         BlitToRegion SourceDC, TargetDC, BorderWidth, TargetHeight, SegArray(SegmentNdx), StartX, StartY

      Case RIGHT_SEGMENT
         If bMOut Then
            PaintVerticalBorderGradient SourceDC, BorderWidth, TargetHeight, uBIH, lBits()
         Else
            PaintVerticalBorderGradient SourceDC, BorderWidth, TargetHeight, uBIH, lBits()
         End If
         BlitToRegion SourceDC, TargetDC, BorderWidth, TargetHeight, SegArray(SegmentNdx), StartX, StartY

      Case TOP_SEGMENT
         PaintHorizontalBorderGradient SourceDC, BorderWidth, TargetWidth, uBIH, lBits()
         BlitToRegion SourceDC, TargetDC, TargetWidth, BorderWidth, SegArray(SegmentNdx), StartX, StartY

      Case BOTTOM_SEGMENT
         If bMOut Then
            PaintHorizontalBorderGradient SourceDC, BorderWidth, TargetWidth, uBIH, lBits()
         Else
            PaintHorizontalBorderGradient SourceDC, BorderWidth, TargetWidth, uBIH, lBits()
         End If
         BlitToRegion SourceDC, TargetDC, TargetWidth, BorderWidth, SegArray(SegmentNdx), StartX, StartY

   End Select

End Sub

Private Sub PaintHorizontalBorderGradient(ByVal TargetDC As Long, ByVal BorderWidth As Long, ByVal TargetWidth As Long, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

'*************************************************************************
'* paints appropriate horizontal gradient to horizontal virtual bitmap.  *
'*************************************************************************

   Call StretchDIBits(TargetDC, 0, 0, TargetWidth, BorderWidth, 0, 1, TargetWidth, BorderWidth - 1, _
                      lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)

End Sub

Private Sub PaintVerticalBorderGradient(ByVal TargetDC As Long, ByVal BorderWidth As Long, ByVal TargetHeight As Long, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

'*************************************************************************
'* paints appropriate vertical gradient to vertical virtual bitmap.      *
'*************************************************************************

   Call StretchDIBits(TargetDC, 0, 0, BorderWidth, TargetHeight, 1, 0, BorderWidth - 1, TargetHeight, _
                      lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)

End Sub

Private Sub MoveRegionToXY(ByVal Rgn As Long, ByVal x As Long, ByVal y As Long)

'*************************************************************************
'* moves the supplied region to absolute X,Y coordinates.                *
'*************************************************************************

   Dim R As RECT    ' holds current X and Y coordinates of region.

'  get the current X,Y coordinates of the region.
   GetRgnBox Rgn, R

'  shift the region to 0,0 then to X,Y.
   OffsetRgn Rgn, -R.Left + x, -R.Top + y

End Sub

Private Sub BlitToRegion(ByVal SourceDC As Long, DestDC As Long, lWidth As Long, lHeight As Long, Region As Long, ByVal XPos As Long, ByVal YPos As Long)

'*************************************************************************
'* blits the contents of a source DC to a non-rectangular region in a    *
'* destination DC.  A clipping region is selected in the destination DC, *
'* then the source DC is blitted to that location.  Technique is used in *
'* this control to blit to the trapezoid-shaped border regions.  Thanks  *
'* to LaVolpe for his help in tweaking this routine.                     *
'*************************************************************************

   Dim R              As Long    ' bitblt function call return.
   Dim ClippingRegion As Long    ' clipping region for bitblt.

'  move the region to the desired position.
   MoveRegionToXY Region, XPos, YPos

'  select a clipping region consisting of the segment parameter.
   ClippingRegion = SelectClipRgn(DestDC, Region)

'  blit the virtual bitmap to the control or form.  Since the clipping region has been
'  selected, only that region-shaped portion of the background will actually be drawn.
   R = BitBlt(DestDC, XPos, YPos, lWidth, lHeight, SourceDC, 0, 0, vbSrcCopy)

'  remove the clipping region constraint from the control.
   SelectClipRgn DestDC, ByVal 0&

'  reset the region coordinates to 0,0.
   MoveRegionToXY Region, 0, 0

End Sub

Private Sub DisplayButton(ByVal WhichButton As Long, Optional ByVal ButtonDown As Boolean = False)

'*************************************************************************
'* displays button with appropriate graphics for button and click state. *
'*************************************************************************

   If WhichButton = LEFT_OR_TOP_BUTTON Then
      If ButtonDown Then
         BitBlt hdc, ButtonCoords(LEFT_OR_TOP_BUTTON).X1, ButtonCoords(LEFT_OR_TOP_BUTTON).Y1, _
                ButtonWidth, ButtonHeight, UD_VirtualDC_LT_ButtonDown, 0, 0, vbSrcCopy
      Else
         BitBlt hdc, ButtonCoords(LEFT_OR_TOP_BUTTON).X1, ButtonCoords(LEFT_OR_TOP_BUTTON).Y1, _
                ButtonWidth, ButtonHeight, UD_VirtualDC_LT_ButtonUp, 0, 0, vbSrcCopy
      End If
   Else
      If ButtonDown Then
         BitBlt hdc, ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).X1, ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).Y1, _
                ButtonWidth, ButtonHeight, UD_VirtualDC_RB_ButtonDown, 0, 0, vbSrcCopy
      Else
         BitBlt hdc, ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).X1, ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).Y1, _
                ButtonWidth, ButtonHeight, UD_VirtualDC_RB_ButtonUp, 0, 0, vbSrcCopy
      End If
   End If

End Sub

Private Sub CalculateUpDownButtonDimensions()

'*************************************************************************
'* calculates the pixel height and width of the UpDown buttons.          *
'*************************************************************************

   Select Case m_UD_Orientation
      Case [Vertical]
        ButtonWidth = ScaleWidth - 2 * m_UD_BorderWidth
        ButtonHeight = (ScaleHeight - (2 * m_UD_BorderWidth)) \ 2
      Case [Horizontal]
         ButtonWidth = (ScaleWidth - (2 * m_UD_BorderWidth)) \ 2
         ButtonHeight = ScaleHeight - 2 * m_UD_BorderWidth
   End Select

End Sub

Private Sub CalculateUpDownButtonCoordinates()

'*************************************************************************
'* determines top left / bottom right XY coordinates for updown buttons. *
'* coordinates are in pixels and account for possible border.            *
'*************************************************************************

'  top or left button has same coordinates within control regardless of UpDown button orientation.
   ButtonCoords(LEFT_OR_TOP_BUTTON).X1 = m_UD_BorderWidth
   ButtonCoords(LEFT_OR_TOP_BUTTON).Y1 = m_UD_BorderWidth
   ButtonCoords(LEFT_OR_TOP_BUTTON).X2 = ButtonCoords(LEFT_OR_TOP_BUTTON).X1 + ButtonWidth - 1
   ButtonCoords(LEFT_OR_TOP_BUTTON).Y2 = ButtonCoords(LEFT_OR_TOP_BUTTON).Y1 + ButtonHeight - 1

'  determine bottom or right button coordinates based on UpDown button orientation.
   Select Case m_UD_Orientation
      Case [Vertical]
         ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).X1 = m_UD_BorderWidth
         ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).Y1 = ButtonCoords(LEFT_OR_TOP_BUTTON).Y2 + 1
         ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).X2 = ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).X1 + ButtonWidth - 1
         ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).Y2 = ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).Y1 + ButtonHeight - 1
      Case [Horizontal]
         ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).X1 = ButtonCoords(LEFT_OR_TOP_BUTTON).X2 + 1
         ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).Y1 = m_UD_BorderWidth
         ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).X2 = ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).X1 + ButtonWidth - 1
         ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).Y2 = ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).Y1 + ButtonHeight - 1
   End Select

End Sub

Private Sub CreateUpDownButtonVirtualBitmaps()

'*************************************************************************
'* creates the virtual bitmaps that hold button up and down gradients.   *
'*************************************************************************

   CreateTopOrLeftButtonBitmaps
   CreateBottomOrRightButtonBitmaps

End Sub

Private Sub CreateTopOrLeftButtonBitmaps()

'*************************************************************************
'* creates the virtual bitmaps for top/left up and down button graphics. *
'*************************************************************************

'  'up button'
   CreateVirtualDC hdc, UD_VirtualDC_LT_ButtonUp, UD_mMemoryBitmap_LT_ButtonUp, UD_mOriginalBitmap_LT_ButtonUp, ButtonWidth, ButtonHeight

'  'down button'
   CreateVirtualDC hdc, UD_VirtualDC_LT_ButtonDown, UD_mMemoryBitmap_LT_ButtonDown, UD_mOriginalBitmap_LT_ButtonDown, ButtonWidth, ButtonHeight

End Sub

Private Sub CreateBottomOrRightButtonBitmaps()

'*************************************************************************
'* creates virtual bitmaps for bottom/right up and down button graphics. *
'*************************************************************************

'  'up button'
   CreateVirtualDC hdc, UD_VirtualDC_RB_ButtonUp, UD_mMemoryBitmap_RB_ButtonUp, UD_mOriginalBitmap_RB_ButtonUp, ButtonWidth, ButtonHeight

'  'down button'
   CreateVirtualDC hdc, UD_VirtualDC_RB_ButtonDown, UD_mMemoryBitmap_RB_ButtonDown, UD_mOriginalBitmap_RB_ButtonDown, ButtonWidth, ButtonHeight

End Sub

Private Sub GenerateUpDownButtonGradients()

'*************************************************************************
'* initializes UpDown button gradients.                                  *
'*************************************************************************

'  calculate the 'button up' gradient.
   CalculateGradient ButtonWidth, ButtonHeight, TranslateColor(Active_UD_ButtonColor1), TranslateColor(Active_UD_ButtonColor2), _
                     Active_UD_ButtonUpAngle, Active_UD_ButtonUpMidOut, UD_ButtonUp_uBIH, UD_ButtonUp_lBits()

'  Note: I use a different 'button down' gradient for each button so that the buttons appear to be
'  tilting in opposite directions when clicked.  The assumption here is that most people will use a
'  middle-out gradient in 'button up' mode and no middle-out in 'button down' mode, to give the
'  visual impression that the button is tilting when clicked.

'  calculate the 'button down' gradient for left or top button.
   CalculateGradient ButtonWidth, ButtonHeight, TranslateColor(Active_UD_ButtonColor2), TranslateColor(Active_UD_ButtonColor1), _
                     UD_ButtonDownAngle, UD_ButtonDownMidOut, UD_ButtonDown_LT_uBIH, UD_ButtonDown_LT_lBits()

'  calculate the 'button down' gradient for right or bottom button.
   CalculateGradient ButtonWidth, ButtonHeight, TranslateColor(Active_UD_ButtonColor1), TranslateColor(Active_UD_ButtonColor2), _
                     UD_ButtonDownAngle, UD_ButtonDownMidOut, UD_ButtonDown_RB_uBIH, UD_ButtonDown_RB_lBits()

End Sub

Private Sub TransferButtonGradientsToVirtualBitmaps()

'*************************************************************************
'* paints gradient information onto UpDown button virtual bitmaps.       *
'*************************************************************************

'  top or left 'button up'.
   Call StretchDIBits(UD_VirtualDC_LT_ButtonUp, 0, 0, ButtonWidth, ButtonHeight, 0, 0, ButtonWidth, _
                      ButtonHeight, UD_ButtonUp_lBits(0), UD_ButtonUp_uBIH, DIB_RGB_COLORS, vbSrcCopy)

'  top or left 'button down'.
   Call StretchDIBits(UD_VirtualDC_LT_ButtonDown, 0, 0, ButtonWidth, ButtonHeight, 0, 0, ButtonWidth, _
                      ButtonHeight, UD_ButtonDown_LT_lBits(0), UD_ButtonDown_LT_uBIH, DIB_RGB_COLORS, vbSrcCopy)

'  bottom or right 'button up'.
   Call StretchDIBits(UD_VirtualDC_RB_ButtonUp, 0, 0, ButtonWidth, ButtonHeight, 0, 0, ButtonWidth, ButtonHeight, _
                      UD_ButtonUp_lBits(0), UD_ButtonUp_uBIH, DIB_RGB_COLORS, vbSrcCopy)

'  bottom or right 'button down'.
   Call StretchDIBits(UD_VirtualDC_RB_ButtonDown, 0, 0, ButtonWidth, ButtonHeight, 0, 0, ButtonWidth, ButtonHeight, _
                      UD_ButtonDown_RB_lBits(0), UD_ButtonDown_RB_uBIH, DIB_RGB_COLORS, vbSrcCopy)

'  since the generated gradients are painted to virtual bitmaps for the duration
'  of the control's existence (or until Theme is changed), free up some resources.
   Erase UD_ButtonDown_RB_lBits
   Erase UD_ButtonUp_lBits
   Erase UD_ButtonDown_LT_lBits

End Sub

Private Sub TransferArrowsToVirtualBitmaps()

'*************************************************************************
'* draws the UpDown button arrows on the four button virtual bitmaps.    *
'*************************************************************************

   Dim XOffset As Long    ' how many pixels from the left the arrow starts.
   Dim YOffset As Long    ' how many pixels from the top the arrow starts.

   Select Case m_UD_Orientation
      Case [Vertical]
         XOffset = (ButtonWidth \ 2) - 4
         YOffset = (ButtonHeight \ 2) - 2
         DrawUpArrow UD_VirtualDC_LT_ButtonUp, XOffset, YOffset           ' up arrow unclicked.
         DrawUpArrow UD_VirtualDC_LT_ButtonDown, XOffset, YOffset - 1     ' up arrow clicked.
         DrawDownArrow UD_VirtualDC_RB_ButtonUp, XOffset, YOffset         ' down arrow unclicked.
         DrawDownArrow UD_VirtualDC_RB_ButtonDown, XOffset, YOffset + 1   ' down arrow clicked.
      Case [Horizontal]
         XOffset = (ButtonWidth \ 2) - 3
         YOffset = (ButtonHeight \ 2) - 4
         DrawLeftArrow UD_VirtualDC_LT_ButtonUp, XOffset, YOffset         ' left arrow unclicked.
         DrawLeftArrow UD_VirtualDC_LT_ButtonDown, XOffset - 1, YOffset   ' left arrow clicked.
         DrawRightArrow UD_VirtualDC_RB_ButtonUp, XOffset + 1, YOffset    ' right arrow unclicked.
         DrawRightArrow UD_VirtualDC_RB_ButtonDown, XOffset + 3, YOffset  ' right arrow clicked.
   End Select

End Sub

Private Sub DrawLeftArrow(ByVal vDC As Long, ByVal XOffset As Long, YOffset As Long)

'*************************************************************************
'* draws left arrow to appropriate UpDown button virtual bitmap.         *
'*************************************************************************

   Dim I      As Long    ' loop variable.
   Dim AColor As Long    ' translated arrow color.

   AColor = TranslateColor(Active_UD_ArrowColor)

   SetPixelV vDC, 4 + XOffset, YOffset, AColor
   For I = 3 To 4: SetPixelV vDC, I + XOffset, 1 + YOffset, AColor: Next I
   For I = 2 To 4: SetPixelV vDC, I + XOffset, 2 + YOffset, AColor: Next I
   For I = 1 To 4: SetPixelV vDC, I + XOffset, 3 + YOffset, AColor: Next I
   For I = 0 To 4: SetPixelV vDC, I + XOffset, 4 + YOffset, AColor: Next I
   For I = 1 To 4: SetPixelV vDC, I + XOffset, 5 + YOffset, AColor: Next I
   For I = 2 To 4: SetPixelV vDC, I + XOffset, 6 + YOffset, AColor: Next I
   For I = 3 To 4: SetPixelV vDC, I + XOffset, 7 + YOffset, AColor: Next I
   SetPixelV vDC, 4 + XOffset, 8 + YOffset, AColor

End Sub

Private Sub DrawRightArrow(ByVal vDC As Long, ByVal XOffset As Long, YOffset As Long)

'*************************************************************************
'* draws right arrow to appropriate UpDown button virtual bitmap.        *
'*************************************************************************

   Dim I      As Long    ' loop variable.
   Dim AColor As Long    ' translated arrow color.

   AColor = TranslateColor(Active_UD_ArrowColor)

   SetPixelV vDC, XOffset, YOffset, AColor
   For I = 0 To 1: SetPixelV vDC, I + XOffset, 1 + YOffset, AColor: Next I
   For I = 0 To 2: SetPixelV vDC, I + XOffset, 2 + YOffset, AColor: Next I
   For I = 0 To 3: SetPixelV vDC, I + XOffset, 3 + YOffset, AColor: Next I
   For I = 0 To 4: SetPixelV vDC, I + XOffset, 4 + YOffset, AColor: Next I
   For I = 0 To 3: SetPixelV vDC, I + XOffset, 5 + YOffset, AColor: Next I
   For I = 0 To 2: SetPixelV vDC, I + XOffset, 6 + YOffset, AColor: Next I
   For I = 0 To 1: SetPixelV vDC, I + XOffset, 7 + YOffset, AColor: Next I
   SetPixelV vDC, XOffset, 8 + YOffset, AColor

End Sub

Private Sub DrawUpArrow(ByVal vDC As Long, ByVal XOffset As Long, YOffset As Long)

'*************************************************************************
'* draws up arrow to appropriate UpDown button virtual bitmap.           *
'*************************************************************************

   Dim I      As Long    ' loop variable.
   Dim AColor As Long    ' translated arrow color.

   AColor = TranslateColor(Active_UD_ArrowColor)

   SetPixelV vDC, 4 + XOffset, YOffset, AColor
   For I = 3 To 5: SetPixelV vDC, I + XOffset, 1 + YOffset, AColor: Next I
   For I = 2 To 6: SetPixelV vDC, I + XOffset, 2 + YOffset, AColor: Next I
   For I = 1 To 7: SetPixelV vDC, I + XOffset, 3 + YOffset, AColor: Next I
   For I = 0 To 8: SetPixelV vDC, I + XOffset, 4 + YOffset, AColor: Next I

End Sub

Private Sub DrawDownArrow(ByVal vDC As Long, ByVal XOffset As Long, YOffset As Long)

'*************************************************************************
'* draws down arrow to appropriate UpDown button virtual bitmap.         *
'*************************************************************************

   Dim I      As Long    ' loop variable.
   Dim AColor As Long    ' translated arrow color.

   AColor = TranslateColor(Active_UD_ArrowColor)

   For I = 0 To 8: SetPixelV vDC, I + XOffset, 0 + YOffset, AColor: Next I
   For I = 1 To 7: SetPixelV vDC, I + XOffset, 1 + YOffset, AColor: Next I
   For I = 2 To 6: SetPixelV vDC, I + XOffset, 2 + YOffset, AColor: Next I
   For I = 3 To 5: SetPixelV vDC, I + XOffset, 3 + YOffset, AColor: Next I
   SetPixelV vDC, 4 + XOffset, 4 + YOffset, AColor

End Sub

Private Sub InitializeUpDownBorderGraphics()

'*************************************************************************
'* creates the four border segments for the UpDown control.              *
'*************************************************************************

'  create the horizontal border segment virtual DC.
   CreateVirtualDC hdc, UD_VirtualDC_SegH, UD_mMemoryBitmap_SegH, UD_mOriginalBitmap_SegH, ScaleWidth + 1, m_UD_BorderWidth

'  create the vertical border segment virtual DC.
   CreateVirtualDC hdc, UD_VirtualDC_SegV, UD_mMemoryBitmap_SegV, UD_mOriginalBitmap_SegV, m_UD_BorderWidth, ScaleHeight

'  calculate the primary horizontal segment gradient.
   CalculateGradient ScaleWidth, m_UD_BorderWidth + 1, TranslateColor(Active_UD_BorderColor1), _
                     TranslateColor(Active_UD_BorderColor2), 90, m_UD_BorderMiddleOut, UD_SegH1uBIH, UD_SegH1lBits()

'  if gradients are not middle-out, calculate the secondary horizontal segment gradient.
   If Not m_UD_BorderMiddleOut Then
      CalculateGradient ScaleWidth, m_UD_BorderWidth + 1, TranslateColor(Active_UD_BorderColor2), TranslateColor(Active_UD_BorderColor1), _
                        90, m_UD_BorderMiddleOut, UD_SegH2uBIH, UD_SegH2lBits()
   End If

'  calculate the primary vertical segment gradient.
   CalculateGradient m_UD_BorderWidth + 1, ScaleHeight, TranslateColor(Active_UD_BorderColor1), TranslateColor(Active_UD_BorderColor2), _
                     180, m_UD_BorderMiddleOut, UD_SegV1uBIH, UD_SegV1lBits()

'  if gradients are not middle-out, calculate the secondary vertical segment gradient.
   If Not m_UD_BorderMiddleOut Then
      CalculateGradient m_UD_BorderWidth + 1, ScaleHeight, TranslateColor(Active_UD_BorderColor2), TranslateColor(Active_UD_BorderColor1), _
                        180, m_UD_BorderMiddleOut, UD_SegV2uBIH, UD_SegV2lBits()
   End If

'  create the four border segments.
   CreateUpDownBorderSegments

End Sub

Private Sub CreateUpDownBorderSegments()

'*************************************************************************
'* creates the vertical and horizontal trapezoidal border regions.       *
'*************************************************************************

   DeleteUpDownBorderSegmentObjects    ' make sure the segments don't already exist.

   UD_BorderSegment(TOP_SEGMENT) = CreateDiagRectRegion(ScaleWidth, m_UD_BorderWidth, 1, 1)
   UD_BorderSegment(BOTTOM_SEGMENT) = CreateDiagRectRegion(ScaleWidth, m_UD_BorderWidth, -1, -1)
   UD_BorderSegment(RIGHT_SEGMENT) = CreateDiagRectRegion(m_UD_BorderWidth, ScaleHeight, -1, -1)
   UD_BorderSegment(LEFT_SEGMENT) = CreateDiagRectRegion(m_UD_BorderWidth, ScaleHeight, 1, 1)

End Sub

Private Sub DeleteUpDownBorderSegmentObjects()

'*************************************************************************
'* destroys the border segment objects if they exist, to save memory.    *
'*************************************************************************

   If UD_BorderSegment(TOP_SEGMENT) Then
      DeleteObject UD_BorderSegment(TOP_SEGMENT)
   End If

   If UD_BorderSegment(RIGHT_SEGMENT) Then
      DeleteObject UD_BorderSegment(RIGHT_SEGMENT)
   End If

   If UD_BorderSegment(BOTTOM_SEGMENT) Then
      DeleteObject UD_BorderSegment(BOTTOM_SEGMENT)
   End If

   If UD_BorderSegment(LEFT_SEGMENT) Then
      DeleteObject UD_BorderSegment(LEFT_SEGMENT)
   End If

End Sub

Private Sub DeleteUpDownBorderVirtualDCs()

'**************************************************************************
'* destroys the updown borders' virtual bitmaps upon control termination. *
'**************************************************************************

   DestroyVirtualDC UD_VirtualDC_SegH, UD_mMemoryBitmap_SegH, UD_mOriginalBitmap_SegH
   DestroyVirtualDC UD_VirtualDC_SegV, UD_mMemoryBitmap_SegV, UD_mOriginalBitmap_SegV

End Sub

Private Sub DeleteUpDownButtonVirtualDCs()

'**************************************************************************
'* destroys the updown buttons' virtual bitmaps upon control termination. *
'**************************************************************************

   DestroyVirtualDC UD_VirtualDC_LT_ButtonUp, UD_mMemoryBitmap_LT_ButtonUp, UD_mOriginalBitmap_LT_ButtonUp
   DestroyVirtualDC UD_VirtualDC_LT_ButtonDown, UD_mMemoryBitmap_LT_ButtonDown, UD_mOriginalBitmap_LT_ButtonDown
   DestroyVirtualDC UD_VirtualDC_RB_ButtonUp, UD_mMemoryBitmap_RB_ButtonUp, UD_mOriginalBitmap_RB_ButtonUp
   DestroyVirtualDC UD_VirtualDC_RB_ButtonDown, UD_mMemoryBitmap_RB_ButtonDown, UD_mOriginalBitmap_RB_ButtonDown

End Sub

Private Sub DeleteRangeWindowVirtualDCs()

'**************************************************************************
'* destroys the RangeWindow virtual bitmaps upon control termination.     *
'**************************************************************************

   DestroyVirtualDC RW_VirtualDC, RW_mMemoryBitmap, RW_mOriginalBitmap
   DestroyVirtualDC RW_BG_VirtualDC, RW_BG_mMemoryBitmap, RW_BG_mOriginalBitmap

End Sub

Private Sub DeleteRangeWindowBorderVirtualDCs()

'**************************************************************************
'* destroys the updown borders' virtual bitmaps upon control termination. *
'**************************************************************************

   DestroyVirtualDC RW_VirtualDC_SegH, RW_mMemoryBitmap_SegH, RW_mOriginalBitmap_SegH
   DestroyVirtualDC RW_VirtualDC_SegV, RW_mMemoryBitmap_SegV, RW_mOriginalBitmap_SegV

End Sub

Private Function CreateDiagRectRegion(ByVal cx As Long, ByVal cy As Long, SideAStyle As Integer, SideBStyle As Integer) As Long

'**************************************************************************
'* Author: LaVolpe                                                        *
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
         tpts(0).x = cy - 1
         tpts(1).x = -1
      ElseIf SideAStyle > 0 Then
         tpts(1).x = cy
      End If
      tpts(1).y = cy

      tpts(2).x = cx + Abs(SideBStyle < 0)
      If SideBStyle > 0 Then tpts(2).x = tpts(2).x - cy
      tpts(2).y = cy

      tpts(3).x = cx + Abs(SideBStyle < 0)
      If SideBStyle < 0 Then tpts(3).x = tpts(3).x - cy

   Else

'     absolute minimum width & height of a trapezoid
      If Abs(SideAStyle + SideBStyle) = 2 Then ' has 2 opposing slanted sides
         If cy < cx * 2 Then cx = cy \ 2
      End If

      If SideAStyle < 0 Then
         tpts(0).y = cx - 1
         tpts(3).y = -1
      ElseIf SideAStyle > 0 Then
         tpts(3).y = cx - 1
         tpts(0).y = -1
      End If

      tpts(1).y = cy
      If SideBStyle < 0 Then tpts(1).y = tpts(1).y - cx
      tpts(2).x = cx

      tpts(2).y = cy
      If SideBStyle > 0 Then tpts(2).y = tpts(2).y - cx
      tpts(3).x = cx

   End If

   tpts(4) = tpts(0)

   CreateDiagRectRegion = CreatePolygonRgn(tpts(0), UBound(tpts) + 1, 2)

End Function

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* converts color long COLORREF for api coloring purposes.               *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If

End Function

Private Sub CalculateGradient(Width As Long, Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, _
                              ByVal Angle As Single, ByVal bMOut As Boolean, ByRef uBIH As BITMAPINFOHEADER, ByRef lBits() As Long)

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
   Dim I         As Long, j As Long, k As Long
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
         For I = 0 To iEnd
            lGrad2(I) = b1 + (dB * I) \ iEnd + 256 * (G1 + (dG * I) \ iEnd) + 65536 * (R1 + (dR * I) \ iEnd)
         Next I
      End If

'     'if' block added by Matthew R. Usner - accounts for possible MiddleOut gradient draw.
      If bMOut Then
         k = 0
         For I = 0 To iEnd Step 2
            lGrad(k) = lGrad2(I)
            k = k + 1
         Next I
         For I = iEnd - 1 To 1 Step -2
            lGrad(k) = lGrad2(I)
            k = k + 1
         Next I
      Else
         For I = 0 To iEnd
            lGrad(I) = lGrad2(I)
         Next I
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
               For I = 0 To iEnd
                  lBits(I + Offset) = lGrad((I * luSin + jIn) \ INT_ROT)
               Next I
               jIn = jIn + luCos
               Offset = Offset + Scan
            Next j

         Case 1, 3
            luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
            luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
            Offset = jEnd * Scan
            jIn = 0
            For j = 0 To jEnd
               For I = 0 To iEnd
                  lBits(I + Offset) = lGrad((I * luSin + jIn) \ INT_ROT)
               Next I
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

Private Sub GetEnabledDisplayProperties()

'*************************************************************************
'* applies enabled graphics properties to the active display properties. *
'*************************************************************************

   Active_UD_ArrowColor = m_UD_ArrowColor
   Active_UD_BorderColor1 = m_UD_BorderColor1
   Active_UD_BorderColor2 = m_UD_BorderColor2
   Active_UD_ButtonColor1 = m_UD_ButtonColor1
   Active_UD_ButtonColor2 = m_UD_ButtonColor2
   Active_UD_ButtonUpAngle = m_UD_ButtonUpAngle     ' UpAngle used for both enabled/disabled.
   Active_UD_ButtonUpMidOut = m_UD_ButtonUpMidOut   ' UpMidOut used for both enabled/disabled.

End Sub

Private Sub GetDisabledDisplayProperties()

'*************************************************************************
'* applies disabled graphics properties to active display properties.    *
'*************************************************************************

   Active_UD_ArrowColor = m_UD_DisArrowColor
   Active_UD_BorderColor1 = m_UD_DisBorderColor1
   Active_UD_BorderColor2 = m_UD_DisBorderColor2
   Active_UD_ButtonColor1 = m_UD_DisButtonColor1
   Active_UD_ButtonColor2 = m_UD_DisButtonColor2
   Active_UD_ButtonUpAngle = m_UD_ButtonUpAngle     ' UpAngle used for both enabled/disabled.
   Active_UD_ButtonUpMidOut = m_UD_ButtonUpMidOut   ' UpMidOut used for both enabled/disabled.

End Sub

Private Sub InitializeRangeWindowGraphics()

'*************************************************************************
'* master procedure for initializing all RangeWindow graphics.           *
'*************************************************************************

   RangeWindowWidth = 180
   RangeWindowHeight = IIf(m_RW_ShowLED = True, 60, 40)

'  create a virtual bitmap to contain RangeWindow graphics for blitting when popped.
   CreateVirtualDC hdc, RW_VirtualDC, RW_mMemoryBitmap, RW_mOriginalBitmap, RangeWindowWidth, RangeWindowHeight
   
'  create the border graphics.
   InitializeRangeWindowBorderGraphics

'  copy the border graphics to the RangeWindow virtual bitmap.
   DisplayRangeWindowBorder

'  create LED display background graphics.
   InitializeRangeWindowBackgroundGraphics

'  create LED digit segment regions.
   CreateLEDSegmentRegions

'  create the color brush used to paint the segment regions.
   CreateLEDColorBrush

'  set up the segment display patterns for LED digits 0-9, burn-in and negative sign.
   DisplayPattern() = Split("1111110,0110000,1101101,1111001,0110011,1011011,1011111,1110000,1111111,1111011,0000000,0000001", ",")

'  calculate the x coordinate of each digit in the LED display.
   MapLEDDigits

'  generate gradient for bar meter.
   CalculateGradient RW_BG_Width, RW_BG_Height, TranslateColor(m_RW_PBarColor1), TranslateColor(m_RW_PBarColor2), _
                     90, True, RW_Meter_uBIH, RW_Meter_lBits()

'  initialize variable used in displaying changed LED digits.
   PreviousValue = String(MAX_DIGITS, "@")

End Sub

Private Sub CreateLEDSegmentRegions()

'*************************************************************************
'* creates the vertical and horizontal rectangular LCD segment regions.  *
'*************************************************************************

'  a safety net to make sure segments have not already been created.
   DeleteLEDSegmentRegions

'  create the segment regions.
   LEDSegment(VERTICAL_LED_SEGMENT) = CreateRectRgn(0, 0, SegmentWidth, SegmentHeight)
   LEDSegment(HORIZONTAL_LED_SEGMENT) = CreateRectRgn(0, 0, SegmentHeight, SegmentWidth)

End Sub

Private Sub CreateLEDColorBrush()

'*************************************************************************
'*  generates the color brush to fill the lit segment objects with.      *
'*************************************************************************

'  a safety net to make sure brushes have not already been created.
   If LEDLitColorBrush Then
      DeleteObject LEDLitColorBrush
   End If
   If LEDBurnInColorBrush Then
      DeleteObject LEDBurnInColorBrush
   End If

'  create the brushes.
   LEDLitColorBrush = CreateSolidBrush(TranslateColor(m_RW_LED_DigitColor))
   LEDBurnInColorBrush = CreateSolidBrush(TranslateColor(m_RW_LED_BurnInColor))

End Sub

Private Sub DeleteLEDSegmentRegions()

'*************************************************************************
'* destroys the LED segment regions if they have been created.           *
'*************************************************************************

   If LEDSegment(VERTICAL_LED_SEGMENT) Then
      DeleteObject LEDSegment(VERTICAL_LED_SEGMENT)
   End If
   If LEDSegment(HORIZONTAL_LED_SEGMENT) Then
      DeleteObject LEDSegment(HORIZONTAL_LED_SEGMENT)
   End If

End Sub

Private Sub MapLEDDigits()

'*************************************************************************
'* maps the starting x position in RangeWindow of each LED digit.        *
'*************************************************************************

   Dim I As Long    ' loop variable.

   DigitWidth = SegmentHeight + (2 * SegmentWidth) + (2 * InterSegmentGap) - 3
   DigitHeight = (2 * SegmentHeight) + (3 * SegmentWidth) + (4 * InterSegmentGap) - 4

'  calculate and store the x offset for each display digit.
   DigitXPos(0) = m_RW_BorderWidth + 4
   For I = 1 To MAX_DIGITS - 1
      DigitXPos(I) = DigitXPos(I - 1) + DigitWidth + InterDigitGap
   Next I

End Sub

Private Sub DisplayValue()

'*************************************************************************
'* displays the .Value in the RangeWindow LED display.                   *
'*************************************************************************

   Dim CurrentDigit As String    ' the current digit to be displayed.
   Dim I As Long                 ' loop variable.
   Dim sVal As String            ' right-justified version of value to be displayed.

'  obtain the string equivalent of the value and add leading spaces if necessary.
   sVal = CStr(m_Value)
   sVal = Right(String(MAX_DIGITS, " ") & sVal, MAX_DIGITS)

   For I = 1 To MAX_DIGITS
      CurrentDigit = Mid(sVal, I, 1)
'     only draw a digit if it is different from previous value's digit to eliminate flicker.
      If CurrentDigit <> Mid(PreviousValue, I, 1) Then
'        blit appropriate part of background to erase old digit.
         BitBlt RWhDC, DigitXPos(I - 1), m_RW_BorderWidth + 1, DigitWidth, DigitHeight, RW_BG_VirtualDC, DigitXPos(I - 1), 0, vbSrcCopy
'        display the current LED digit.
         DisplayRectangularSegmentDigit CurrentDigit, LEDSegment(), DigitXPos(I - 1), m_RW_BorderWidth + 1, _
                                        SegmentHeight, SegmentWidth, InterSegmentGap
      End If
   Next I

'  since only changed digits are updated in the LED display, update
'  the previous value variable to prepare for display of next value.
   PreviousValue = sVal

End Sub

Private Sub DisplayRectangularSegmentDigit(ByVal strDigit As String, ByRef LED() As Long, _
                                           ByVal OffsetX As Long, ByVal OffsetY As Long, _
                                           ByVal SegmentHeight As Long, ByVal SegmentWidth As Long, ByVal SegmentGap As Long)

'*************************************************************************
'* displays a rectangular-segment display digit according to pattern.    *
'*************************************************************************

   Dim Digit               As Long    ' the display pattern index of the current digit to draw.
'  used to avoid unnecessary recalculations of segment gap multiples.
   Dim DoubleSegmentGap    As Long
   Dim TripleSegmentGap    As Long
   Dim DoubleSegmentWidth  As Long

   DoubleSegmentGap = 2 * SegmentGap
   TripleSegmentGap = 3 * SegmentGap
   DoubleSegmentWidth = 2 * SegmentWidth

'  get the appropriate segment display pattern for the digit.
   If strDigit = " " Then
'     determine index for 'burn-in', or ignore (-1) if not showing burned in digits.
      Digit = IIf(m_RW_LED_ShowBurnIn = True, 10, -1)
      If Digit = -1 Then
         Exit Sub
      End If
   Else
      Digit = InStr("0123456789 -", strDigit) - 1
   End If

'  segment 1 (top)
   DisplaySegment LED(), HORIZONTAL_LED_SEGMENT, OffsetX + SegmentWidth + SegmentGap - 1, _
                  OffsetY, Mid(DisplayPattern(Digit), 1, 1)

'  segment 2 (top right)
   DisplaySegment LED(), VERTICAL_LED_SEGMENT, OffsetX + SegmentWidth + SegmentHeight + DoubleSegmentGap - 2, _
                  OffsetY + SegmentWidth + SegmentGap - 1, Mid(DisplayPattern(Digit), 2, 1)

'  segment 3 (bottom right)
   DisplaySegment LED(), VERTICAL_LED_SEGMENT, OffsetX + SegmentWidth + SegmentHeight + DoubleSegmentGap - 2, _
                  OffsetY + SegmentHeight + DoubleSegmentWidth + TripleSegmentGap - 3, Mid(DisplayPattern(Digit), 3, 1)

'  segment 4 (bottom)
   DisplaySegment LED(), HORIZONTAL_LED_SEGMENT, OffsetX + SegmentWidth + SegmentGap - 1, _
                  OffsetY + (2 * SegmentHeight) + DoubleSegmentWidth + 4 * SegmentGap - 4, _
                  Mid(DisplayPattern(Digit), 4, 1)

'  segment 5 (bottom left)
   DisplaySegment LED(), VERTICAL_LED_SEGMENT, OffsetX, _
                  OffsetY + SegmentHeight + DoubleSegmentWidth + TripleSegmentGap - 3, _
                  Mid(DisplayPattern(Digit), 5, 1)

'  segment 6 (top left)
   DisplaySegment LED(), VERTICAL_LED_SEGMENT, OffsetX, OffsetY + SegmentWidth + SegmentGap - 1, _
                  Mid(DisplayPattern(Digit), 6, 1)

'  segment 7 (center)
   DisplaySegment LED(), HORIZONTAL_LED_SEGMENT, OffsetX + SegmentWidth + SegmentGap - 1, _
                  OffsetY + SegmentHeight + SegmentWidth + DoubleSegmentGap - 2, _
                  Mid(DisplayPattern(Digit), 7, 1)

End Sub

Private Sub DisplaySegment(LED() As Long, ByVal SegmentNdx As Long, ByVal StartX As Long, ByVal StartY As Long, ByVal LitStatus As String)

'*************************************************************************
'* displays one segment of an LCD digit according to its fill style.     *
'*************************************************************************

'  position the segment region in the correct location.
   OffsetRgn LED(SegmentNdx), StartX, StartY

   If LitStatus = SEGMENT_UNLIT And m_RW_LED_ShowBurnIn Then
'     if segment is unlit but burn-in mode is active, display as unlit according to fill mode.
      FillRgn RWhDC, LED(SegmentNdx), LEDBurnInColorBrush
   Else
      If LitStatus = SEGMENT_LIT Then
'        display segment using lit color brush.
         FillRgn RWhDC, LED(SegmentNdx), LEDLitColorBrush
      End If
   End If

'  reset the region location to (0, 0) to prepare for the next segment draw.
   OffsetRgn LED(SegmentNdx), -StartX, -StartY

End Sub

Private Sub InitializeRangeWindowBackgroundGraphics()

'*************************************************************************
'* creates virtual bitmap and gradient for LED display background.       *
'*************************************************************************

'  calculate the width and height of the RangeWindow background bitmap.
   RW_BG_Width = RangeWindowWidth - 2 * m_RW_BorderWidth
   If m_RW_ShowLED Then
      RW_BG_Height = (RangeWindowHeight - 2 * m_RW_BorderWidth) \ 2
   Else
      RW_BG_Height = RangeWindowHeight - 2 * m_RW_BorderWidth
   End If

'  create virtual bitmap to hold LED background gradient.
   CreateVirtualDC hdc, RW_BG_VirtualDC, RW_BG_mMemoryBitmap, RW_BG_mOriginalBitmap, RW_BG_Width, RW_BG_Height

'  generate gradient information for the virtual bitmap.
   CalculateGradient RW_BG_Width, RW_BG_Height, TranslateColor(m_RW_BackColor1), TranslateColor(m_RW_BackColor2), _
                     m_RW_BackAngle, m_RW_BackMiddleOut, RW_BG_uBIH, RW_BG_lBits()

'  transfer the gradient to the range window background virtual bitmap.
   Call StretchDIBits(RW_BG_VirtualDC, 0, 0, RW_BG_Width, RW_BG_Height, 0, 0, _
                      RW_BG_Width, RW_BG_Height, RW_BG_lBits(0), RW_BG_uBIH, DIB_RGB_COLORS, vbSrcCopy)

End Sub

Private Sub InitializeRangeWindowBorderGraphics()

'*************************************************************************
'* creates region segments and gradients for RangeWindow borders.        *
'*************************************************************************

'  create the horizontal border segment virtual DC.
   CreateVirtualDC RW_VirtualDC, RW_VirtualDC_SegH, RW_mMemoryBitmap_SegH, RW_mOriginalBitmap_SegH, RangeWindowWidth + 1, m_RW_BorderWidth

'  create the vertical border segment virtual DC.
   CreateVirtualDC RW_VirtualDC, RW_VirtualDC_SegV, RW_mMemoryBitmap_SegV, RW_mOriginalBitmap_SegV, m_RW_BorderWidth, RangeWindowHeight

'  calculate the primary horizontal segment gradient.
   CalculateGradient RangeWindowWidth, m_RW_BorderWidth + 1, TranslateColor(m_RW_BorderColor1), TranslateColor(m_RW_BorderColor2), _
                     90, m_RW_BorderMiddleOut, RW_SegH1uBIH, RW_SegH1lBits()

'  if gradients are not middle-out, calculate the secondary horizontal segment gradient.
   If Not m_RW_BorderMiddleOut Then
      CalculateGradient RangeWindowWidth, m_RW_BorderWidth + 1, TranslateColor(m_RW_BorderColor2), TranslateColor(m_RW_BorderColor1), _
                        90, m_RW_BorderMiddleOut, RW_SegH2uBIH, RW_SegH2lBits()
   End If

'  calculate the primary vertical segment gradient.
   CalculateGradient m_RW_BorderWidth + 1, RangeWindowHeight, TranslateColor(m_RW_BorderColor1), TranslateColor(m_RW_BorderColor2), _
                     180, m_RW_BorderMiddleOut, RW_SegV1uBIH, RW_SegV1lBits()

'  if gradients are not middle-out, calculate the secondary vertical segment gradient.
   If Not m_RW_BorderMiddleOut Then    ' use same middle-out style as UpDown main control.
      CalculateGradient m_RW_BorderWidth + 1, RangeWindowHeight, TranslateColor(m_RW_BorderColor2), TranslateColor(m_RW_BorderColor1), _
                        180, m_RW_BorderMiddleOut, RW_SegV2uBIH, RW_SegV2lBits()
   End If

'  create the four border segments.
   CreateRangeWindowBorderSegments

End Sub

Private Sub CreateRangeWindowBorderSegments()

'*************************************************************************
'* creates the vertical and horizontal trapezoidal border segments.      *
'*************************************************************************

   DeleteRangeWindowBorderSegmentObjects    ' make sure the segments don't already exist.

   RW_BorderSegment(TOP_SEGMENT) = CreateDiagRectRegion(RangeWindowWidth, m_RW_BorderWidth, 1, 1)
   RW_BorderSegment(BOTTOM_SEGMENT) = CreateDiagRectRegion(RangeWindowWidth, m_RW_BorderWidth, -1, -1)
   RW_BorderSegment(RIGHT_SEGMENT) = CreateDiagRectRegion(m_RW_BorderWidth, RangeWindowHeight, -1, -1)
   RW_BorderSegment(LEFT_SEGMENT) = CreateDiagRectRegion(m_RW_BorderWidth, RangeWindowHeight, 1, 1)

End Sub

Private Sub DeleteRangeWindowBorderSegmentObjects()

'*************************************************************************
'* destroys the border segment objects if they exist, to save memory.    *
'*************************************************************************

   If RW_BorderSegment(TOP_SEGMENT) Then DeleteObject RW_BorderSegment(TOP_SEGMENT)
   If RW_BorderSegment(RIGHT_SEGMENT) Then DeleteObject RW_BorderSegment(RIGHT_SEGMENT)
   If RW_BorderSegment(BOTTOM_SEGMENT) Then DeleteObject RW_BorderSegment(BOTTOM_SEGMENT)
   If RW_BorderSegment(LEFT_SEGMENT) Then DeleteObject RW_BorderSegment(LEFT_SEGMENT)

End Sub

'******************** Virtual DC Code **********************
Private Sub CreateVirtualDC(TargetDC As Long, vDC As Long, mMB As Long, mOB As Long, ByVal vWidth As Long, ByVal vHeight As Long)

'*************************************************************************
'* creates virtual bitmaps for background and cells.                     *
'*************************************************************************

   If vDC <> 0 Then DestroyVirtualDC vDC, mMB, mOB

'  create a memory device context to use.
   vDC = CreateCompatibleDC(TargetDC)

'  define it as a bitmap so that drawing can be performed to the virtual DC.
   mMB = CreateCompatibleBitmap(TargetDC, vWidth, vHeight)
   mOB = SelectObject(vDC, mMB)

End Sub

Private Function IsCreated(ByVal vDC As Long) As Boolean

'*************************************************************************
'* checks the handle of a virtual DC and returns if it exists.           *
'*************************************************************************

   IsCreated = (vDC <> 0)

End Function

Private Sub DestroyVirtualDC(ByRef vDC As Long, ByVal mMB As Long, ByVal mOB As Long)

'*************************************************************************
'* eliminates a virtual dc bitmap on control's termination.              *
'*************************************************************************

   If vDC = 0 Then Exit Sub

   Call SelectObject(vDC, mOB)
   Call DeleteObject(mMB)
   Call DeleteDC(vDC)
   vDC = 0

End Sub
'********************************************************************

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<< Events >>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_DblClick()

'*************************************************************************
'* lets doubleclick be treated as two rapid clicks for UpDown buttons.   *
'*************************************************************************

'  a DblClick can be generated by right-double-clicking; we don't want that happening in this control.
   If RightMouseButtonClicked Then
      RightMouseButtonClicked = False
      Exit Sub
   End If

'  generate a MouseDown to take the place of the doubleclick.
   mouse_event MOUSEEVENTF_LEFTDOWN, MouseX, MouseY, 0, 0

End Sub

Private Sub UserControl_GotFocus()

'*************************************************************************
'* changes UpDown border colors when control receives the focus.         *
'*************************************************************************

   If m_Enabled Then
      Active_UD_BorderColor1 = m_UD_FocusBorderColor1
      Active_UD_BorderColor2 = m_UD_FocusBorderColor2
      InitializeUpDownBorderGraphics
      DisplayUpDownBorder
      UserControl.Refresh
   End If

End Sub

Private Sub UserControl_Initialize()

'*************************************************************************
'* first step in life of control.                                        *
'*************************************************************************

   RangeWindowWidth = 180
   RangeWindowHeight = 60

End Sub

Private Sub UserControl_LostFocus()

'*************************************************************************
'* changes UpDown border colors when control loses the focus.            *
'*************************************************************************

   If m_Enabled Then
      Active_UD_BorderColor1 = m_UD_BorderColor1
      Active_UD_BorderColor2 = m_UD_BorderColor2
      InitializeUpDownBorderGraphics
      DisplayUpDownBorder
      UserControl.Refresh
   End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* checks for Ctrl and/or Shift keys pressed for increment acceleration. *
'*************************************************************************

   ShiftKeyDown = (Shift And vbShiftMask) > 0
   CtrlKeyDown = (Shift And vbCtrlMask) > 0
   ShiftAndCtrlDown = ShiftKeyDown And CtrlKeyDown

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* checks for Ctrl and/or Shift keys pressed for increment acceleration. *
'*************************************************************************

   ShiftKeyDown = (Shift And vbShiftMask) > 0
   CtrlKeyDown = (Shift And vbCtrlMask) > 0
   ShiftAndCtrlDown = ShiftKeyDown And CtrlKeyDown

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'*************************************************************************
'* controls clicking of the UpDownButtons.                               *
'*************************************************************************

   If Not m_Enabled Then
      Exit Sub
   End If

   If Button = vbRightButton Then
      RightMouseButtonClicked = True  ' to trap right button doubleclicks.
      RaiseEvent MouseDown(Button, Shift, x, y)
      Exit Sub
   End If

   If MouseLocation = MOUSE_NOT_IN_BUTTON Then
      RaiseEvent MouseDown(Button, Shift, x, y)
      Exit Sub
   End If

   MouseIsDown = True    ' mouse has been clicked on a button.

'  keep track of which UpDown button was clicked down on.  This is so the button can be "unclicked"
'  if mouse moves out of that button while the left mouse button is still held down.
   OriginalMouseDownButton = MouseLocation

   DisplayButton MouseLocation, True
   UserControl.Refresh

   IncrementUpDownValue

   RaiseEvent MouseDown(Button, Shift, x, y)

'  check for and process possible continous value scrolling (and RangeWindow popping).
   ProcessContinousScroll

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'*************************************************************************
'* controls release-clicking of the UpDownButtons.                       *
'*************************************************************************

'  ignore right button mouse up.
   If Button = vbRightButton Then
      RaiseEvent MouseUp(Button, Shift, x, y)
      Exit Sub
   End If

   MouseIsDown = False

   If MouseLocation = MOUSE_NOT_IN_BUTTON Then
      RaiseEvent MouseUp(Button, Shift, x, y)
      ProcessRangeWindowMouseUp
      Exit Sub
   End If

'  reset mousedown UpDown button index variable since mouse button has been released.
   OriginalMouseDownButton = 0
   MoveRelease = False

   DisplayButton MouseLocation, False
   UserControl.Refresh

   ProcessRangeWindowMouseUp
   RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'*************************************************************************
'* keeps track of mouse pointer location every time mouse is moved.      *
'*************************************************************************

   Dim PctIntoRange As Single   ' how far (proportionally) into the value range cursor is in progressbar.

'  store mouse coordinates in control-global variables and determine what part of control mouse is in.
   MouseX = x
   MouseY = y
   MouseLocation = GetMouseLocation

'  if an UpDown button was clicked down and mouse pointer dragged out of control a MouseUp event does
'  not register (when mouse is outside control).  This keeps the graphics in sync by resetting the
'  appropriate variables if no button is down when MouseMove is triggered.
   If Button = 0 Then
      MoveRelease = False
      OriginalMouseDownButton = 0
   End If

'  see if the mouse pointer was dragged back into the same button it was dragged
'  out of.  If this is the case, redraw that button in "down" fashion if RangeWindow not active.
   If MoveRelease And OriginalMouseDownButton = MouseLocation And Not RangeWindowPopped Then
      MoveRelease = False
      DisplayButton OriginalMouseDownButton, True
      UserControl.Refresh
   End If

'  see if mouse was clicked over a button then moved out of that button.  If it was, 'unclick' that button.
   If OriginalMouseDownButton <> 0 And OriginalMouseDownButton <> MouseLocation Then
      MoveRelease = True
      DisplayButton OriginalMouseDownButton, False
      UserControl.Refresh
   End If

'  if the RangeWindow is active, calculate the distance left or right the mouse
'  has traveled, and update .Value and RangeWindow display accordingly.
   If RangeWindowPopped Then
      If MouseInRWProgressBar Then
'        find out how far the mouse is, percentagewise, in the range of X coordinates.
         PctIntoRange = Round((CursorPos.x - RW_X1) / (RW_X2 - RW_X1), 2)
         m_Value = m_ValueMin + ((m_ValueMax - m_ValueMin + 1) * PctIntoRange)
         If m_Value < m_ValueMin Then
            m_Value = m_ValueMin
         ElseIf m_Value > m_ValueMax Then
            m_Value = m_ValueMax
         End If
         UpdateRangeWindowDisplay
      Else
'        this If block accounts for when mouse is dragged very quickly to the left or right of
'        the progress bar.  MouseMove events are not always caught when mouse moves very quickly.
'        This doesn't solve the problem entirely, but does help considerably.
         If CursorPos.y >= RW_Y1 And CursorPos.y <= RW_Y2 Then
            If CursorPos.x > RW_X2 And CursorPos.x <= RW_X2 + m_RW_BorderWidth - 1 And m_Value <> m_ValueMax Then
               m_Value = m_ValueMax
               UpdateRangeWindowDisplay
            ElseIf CursorPos.x < RW_X1 And CursorPos.x >= RW_X1 - m_RW_BorderWidth And m_Value <> m_ValueMin Then
               m_Value = m_ValueMin
               UpdateRangeWindowDisplay
            End If
         End If
      End If
   End If

   RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub UserControl_Resize()

'*************************************************************************
'* redraws the control when it's resized in design mode.                 *
'*************************************************************************

   If Ambient.UserMode Then    ' don't execute in runtime.
'      Exit Sub
   End If

'  make sure control's height or width is always even so UpDown buttons fill control.
   If m_UD_Orientation = [Vertical] And ScaleHeight Mod 2 <> 0 Then
      ScaleHeight = ScaleHeight - 1
   ElseIf m_UD_Orientation = [Horizontal] And ScaleWidth Mod 2 <> 0 Then
      ScaleWidth = ScaleWidth - 1
   End If

   InitializeControlGraphics
   RedrawControl

End Sub

Private Sub UserControl_Show()
   
'*************************************************************************
'* performs initial control paint.                                       *
'*************************************************************************

   RedrawControl
   UserControl.Refresh

End Sub

Private Sub UserControl_Terminate()

'*************************************************************************
'* triggers on control termination.                                      *
'*************************************************************************

'  destroy UpDown related objects.
   DeleteUpDownButtonVirtualDCs
   DeleteUpDownBorderSegmentObjects
   DeleteUpDownBorderVirtualDCs

'  destroy RangeWindow-related objects.
   DeleteRangeWindowVirtualDCs
   DeleteRangeWindowBorderVirtualDCs
   DeleteLEDSegmentRegions
   If LEDLitColorBrush Then
      DeleteObject LEDLitColorBrush
   End If
   If LEDBurnInColorBrush Then
      DeleteObject LEDBurnInColorBrush
   End If

'  make sure the range window is destroyed also, if it is active.
   If RangeWindowPopped Then
      DeleteDC RWhDC
      DestroyWindow RWhWnd
   End If

End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<< Miscellaneous Procedures and Functions >>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UpdateRangeWindowDisplay()

'*************************************************************************
'* displays the range window value and progress bar.                     *
'*************************************************************************

   DisplayProgressBar
   If m_RW_ShowLED Then
      DisplayValue
   End If

'  generate Change event if necessary.
   If m_RW_GenerateEvent Then
      RaiseEvent Change
   End If

End Sub

Private Sub ProcessRangeWindowMouseUp()

'*************************************************************************
'* destroys RangeWindow on MouseUp and generates a Change event if the   *
'* RW_GenerateEvent property (for Change event generation in MouseMove)  *
'* has been set to False.  That way, even if no Change events are        *
'* generated in the MouseMove routine, one is still thrown on MouseUp.   *
'*************************************************************************

   If RangeWindowPopped Then
      DeleteDC RWhDC
      DestroyWindow RWhWnd
      RangeWindowPopped = False
      If Not m_RW_GenerateEvent Then
         RaiseEvent Change
      End If
      PreviousValue = String(MAX_DIGITS, "@")    ' re-init variable used in displaying changed LED digits.
   End If

End Sub

Private Function GetMouseLocation() As Long

'*************************************************************************
'* determines if mouse pointer is over one of the UpDown buttons.        *
'*************************************************************************

   If MouseX >= ButtonCoords(LEFT_OR_TOP_BUTTON).X1 And MouseX <= ButtonCoords(LEFT_OR_TOP_BUTTON).X2 And _
      MouseY >= ButtonCoords(LEFT_OR_TOP_BUTTON).Y1 And MouseY <= ButtonCoords(LEFT_OR_TOP_BUTTON).Y2 Then
         GetMouseLocation = MOUSE_IN_LEFT_OR_TOP_BUTTON
   ElseIf MouseX >= ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).X1 And MouseX <= ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).X2 And _
          MouseY >= ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).Y1 And MouseY <= ButtonCoords(RIGHT_OR_BOTTOM_BUTTON).Y2 Then
             GetMouseLocation = MOUSE_IN_RIGHT_OR_BOTTOM_BUTTON
   Else
      GetMouseLocation = MOUSE_NOT_IN_BUTTON
   End If

End Function

Private Sub ProcessContinousScroll()

'*************************************************************************
'* handles continous value scrolling when an UpDown button is held down. *
'* Also pops the RangeWindow if enough time has elapsed.                 *
'*************************************************************************

   Dim OriginalTickCount    As Long   ' comparison tick count for initial value scroll delay.
   Dim OriginalTickCountVal As Long   ' comparison tick count for calculating value increment elapsed time.
   Dim OriginalTickCountRW  As Long   ' comparison tick count for calculating RangeWindow elapsed time.
   Dim CurrentTickCount     As Long   ' current time.

   OriginalTickCountVal = GetTickCount
   OriginalTickCountRW = OriginalTickCountVal
   CurrentTickCount = OriginalTickCountVal

'  create a preliminary delay before values start scrolling.  This gives
'  the user time to unclick the mouse button and prevent unwanted scrolling.
   OriginalTickCount = GetTickCount
   CurrentTickCount = OriginalTickCount
   While MouseIsDown And CurrentTickCount - OriginalTickCount < m_UD_ScrollDelay
      CurrentTickCount = GetTickCount
      DoEvents    ' allow a MouseUp event if it happens.
   Wend

'  if the initial delay has passed without a MouseUp event, start the value scrolling.
   While MouseIsDown And MouseLocation <> MOUSE_NOT_IN_BUTTON And Not RangeWindowPopped
      CurrentTickCount = GetTickCount
      DoEvents ' a MouseUp event will change 'MouseIsDown' to False and terminate the loop.
'     check to see if the time limit for incrementing the value has been exceeded.
      If CurrentTickCount - OriginalTickCountVal >= m_UD_IncrementInterval Then
         IncrementUpDownValue
         OriginalTickCountVal = GetTickCount
      End If
'     check to see if the RangeWindow needs to be popped.
      If CurrentTickCount - OriginalTickCountRW >= m_RW_PopInterval Then
         RangeWindowPopped = True
         DisplayRangeWindow
      End If
   Wend

End Sub

Private Sub DisplayRangeWindow()

'*************************************************************************
'* displays the RangeWindow and sets ZOrder to topmost.                  *
'*************************************************************************

'  set key detect variables in case any are pressed.  They will
'  be reset if any are still pressed when RangeWindow disappears.
   ShiftKeyDown = False
   CtrlKeyDown = False
   ShiftAndCtrlDown = False

'  create the RangeWindow virtual window.
   CreateRangeWindow

'  show the RangeWindow.
   ShowWindow RWhWnd, SW_SHOWNORMAL

'  set RangeWindow to be at the top of the ZOrder so it does not appear under any other windows.
   SetWindowPos RWhWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

'  transfer the contents of the RangeWindow virtual bitmap to the RangeWindow.
   BitBlt RWhDC, 0, 0, RangeWindowWidth, RangeWindowHeight, RW_VirtualDC, 0, 0, vbSrcCopy

'  copy the background virtual bitmap to the top part of the RangeWindow.
   BitBlt RWhDC, m_RW_BorderWidth, m_RW_BorderWidth, (RangeWindowWidth - 2 * m_RW_BorderWidth), _
          (RangeWindowHeight - 2 * m_RW_BorderWidth) \ 2, RW_BG_VirtualDC, 0, 0, vbSrcCopy

'  display the progress bar and LED value.
   If m_RW_ShowLED Then
      DisplayValue
   End If
   DisplayProgressBar

End Sub

Private Sub DisplayProgressBar()
'*************************************************************************
'* displays the appropriate length of the progress bar gradient.         *
'*************************************************************************

   Dim CurrentMeterWidth As Long    ' current width of progressbar.
   Dim PBarY             As Long
   Dim PBarHeight        As Long

   If m_RW_ShowLED Then
      PBarY = RangeWindowHeight \ 2
      PBarHeight = (RangeWindowHeight - 2 * m_RW_BorderWidth) \ 2
   Else
      PBarY = m_RW_BorderWidth
      PBarHeight = RW_BG_Height
   End If

'  determine progress bar width by calculating how far the .Value is in the range.
   CurrentMeterWidth = RW_BG_Width * (Abs(m_Value - m_ValueMin + 1) / Abs(m_ValueMax - m_ValueMin + 1))

'  display the gradient progress meter bar.
   Call StretchDIBits(RWhDC, m_RW_BorderWidth, PBarY, CurrentMeterWidth, RW_BG_Height, 0, 0, _
                      CurrentMeterWidth, RW_BG_Height, RW_Meter_lBits(0), RW_Meter_uBIH, DIB_RGB_COLORS, vbSrcCopy)

'  display the part of the background virtual bitmap that is unaffected by progress bar.
'  Note: by displaying the progress bar, then displaying the part of the background not
'  covered by the progress bar, flicker that would occur (when entire background is drawn
'  then progress bar superimposed on top) is eliminated.
   BitBlt RWhDC, m_RW_BorderWidth + CurrentMeterWidth - 1, PBarY, (RW_BG_Width - CurrentMeterWidth + 1), _
          PBarHeight, RW_BG_VirtualDC, 0, 0, vbSrcCopy

End Sub

Private Sub CreateRangeWindow()

'*************************************************************************
'* creates the virtual window for the RangeWindow, and gets hWnd / hDC.  *
'*************************************************************************

   Dim x As Long    ' x coordinate of top left RangeWindow corner.
   Dim y As Long    ' y coordinate of top left RangeWindow corner.

'  get the absolute screen coordinates of the cursor.  This provides
'  the X,Y coordinates for the top left corner of the RangeWindow.
   GetCursorPos CursorPos

   If UserControl.Parent.ScaleMode = vbTwips Then
      x = (UserControl.Extender.Left / Screen.TwipsPerPixelX)
      y = (UserControl.Extender.Top / Screen.TwipsPerPixelY)
   Else
      x = UserControl.Extender.Left
      y = UserControl.Extender.Top
   End If

'  if necessary, adjust the RangeWindow XY location so it stays fully on screen.
   If CursorPos.x + RangeWindowWidth - 1 > Screen.Width / Screen.TwipsPerPixelX Then
      CursorPos.x = ((Screen.Width / Screen.TwipsPerPixelX) - RangeWindowWidth - 1)
   End If
   If CursorPos.y + RangeWindowHeight - 1 > Screen.Height / Screen.TwipsPerPixelY - TaskBarHeight Then
      CursorPos.y = ((Screen.Height / Screen.TwipsPerPixelY - TaskBarHeight) - RangeWindowHeight - 1)
   End If

'  this defines the screen coordinate rectangle for the progress bar for mouse drag value changing.
   RW_X1 = CursorPos.x + m_RW_BorderWidth - 1
   RW_X2 = RW_X1 + RW_BG_Width - 1
   If m_RW_ShowLED Then
      RW_Y1 = CursorPos.y + (RangeWindowHeight \ 2) - 1
   Else
      RW_Y1 = CursorPos.y + m_RW_BorderWidth - 1
   End If
   RW_Y2 = RW_Y1 + RW_BG_Height - 1

'  create the virtual window on which RangeWindow graphics will be painted.
   RWhWnd = CreateWindowEx(WS_EX_TOOLWINDOW, "Static", "", WS_POPUP Or SS_CUSTOMDRAW, x, y, _
                           RangeWindowWidth, RangeWindowHeight, UserControl.Parent.hwnd, 0, App.hInstance, ByVal 0)
   RWhDC = GetDC(RWhWnd)

'  move the RangeWindow to the screen coordinates specified by the GetCursorPos call.
   MoveWindow RWhWnd, CursorPos.x, CursorPos.y, RangeWindowWidth, RangeWindowHeight, False

End Sub

Private Sub IncrementUpDownValue()

'*************************************************************************
'* increments or decrements the .Value property, accounting for the      *
'* .Wrap and .ValueIncrement properties and generates a Change event.    *
'*************************************************************************

   Dim Increment As Long    ' value increment.

'  determine the correct increment based on Shift and Ctrl key status.
   If ShiftAndCtrlDown Then
      Increment = m_ValueIncrShiftCtrl
   ElseIf CtrlKeyDown Then
      Increment = m_ValueIncrCtrl
   ElseIf ShiftKeyDown Then
      Increment = m_ValueIncrShift
   Else
      Increment = m_ValueIncrement
   End If

   If (m_UD_Orientation = [Vertical] And OriginalMouseDownButton = LEFT_OR_TOP_BUTTON) Or _
      (m_UD_Orientation = [Horizontal] And OriginalMouseDownButton = RIGHT_OR_BOTTOM_BUTTON) Then
         If Not m_UD_SwapDirections Then
            IncrementValue Increment
         Else
            DecrementValue Increment
         End If
   Else
         If Not m_UD_SwapDirections Then
            DecrementValue Increment
         Else
            IncrementValue Increment
         End If
   End If

   RaiseEvent Change

End Sub

Private Sub IncrementValue(ByVal Increment As Long)

'*************************************************************************
'* increments control .Value property, accounting for .Wrap property.    *
'*************************************************************************

   m_Value = m_Value + Increment
   If m_Value > m_ValueMax Then
      m_Value = IIf(m_Wrap = True, m_ValueMin, m_ValueMax)
   End If

End Sub

Private Sub DecrementValue(ByVal Increment As Long)

'*************************************************************************
'* decrements control .Value property, accounting for .Wrap property.    *
'*************************************************************************

   m_Value = m_Value - Increment
   If m_Value < m_ValueMin Then
      m_Value = IIf(m_Wrap = True, m_ValueMax, m_ValueMin)
   End If

End Sub

Private Sub CalculateIncrements()

'*************************************************************************
'* if the .AutoIncrement property is set to True, this routine will det- *
'* ermine the various increments based on the size of the value range.   *
'*************************************************************************

   Dim Range As Long    ' the number of possible values in the value range.

   Range = m_ValueMax - m_ValueMin + 1

   m_ValueIncrement = 1                  ' regular increment always 1.

   m_ValueIncrCtrl = 0.01 * Range        ' ctrl-click = 1% of range.
   If m_ValueIncrCtrl < 1 Then
      m_ValueIncrCtrl = 1
   End If

   m_ValueIncrShift = 0.05 * Range       ' shift-click = 5% of range.
   If m_ValueIncrShift < 1 Then
      m_ValueIncrShift = 1
   End If

   m_ValueIncrShiftCtrl = 0.1 * Range    ' ctrl-shift-click = 10% of range.
   If m_ValueIncrShiftCtrl < 1 Then
      m_ValueIncrShiftCtrl = 1
   End If

End Sub

Private Function MouseInRWProgressBar() As Boolean

'*************************************************************************
'* determines if mouse is in same XY coords as RangeWindow progress bar. *
'* Value can only be changed by dragging mouse in progress bar section   *
'* of the RangeWindow.                                                   *
'*************************************************************************

   GetCursorPos CursorPos    ' get absolute screen coordinates of cursor.
   If CursorPos.x >= RW_X1 And CursorPos.x <= RW_X2 And CursorPos.y >= RW_Y1 And CursorPos.y <= RW_Y2 Then
      MouseInRWProgressBar = True
   End If

End Function

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<< Property Routines >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_InitProperties()

'*************************************************************************
'* default properties for usercontrol.                                   *
'*************************************************************************

   m_Enabled = m_def_Enabled
   m_RW_BackAngle = m_def_RW_BackAngle
   m_RW_BackColor1 = m_def_RW_BackColor1
   m_RW_BackColor2 = m_def_RW_BackColor2
   m_RW_BorderColor1 = m_def_RW_BorderColor1
   m_RW_BorderColor2 = m_def_RW_BorderColor2
   m_RW_BackMiddleOut = m_def_RW_BackMiddleOut
   m_RW_BorderMiddleOut = m_def_RW_BorderMiddleOut
   m_RW_BorderWidth = m_def_RW_BorderWidth
   m_RW_GenerateEvent = m_def_RW_GenerateEvent
   m_RW_LED_BurnInColor = m_def_RW_LED_BurnInColor
   m_RW_LED_DigitColor = m_def_RW_LED_DigitColor
   m_RW_LED_ShowBurnIn = m_def_RW_LED_ShowBurnIn
   m_RW_PBarColor1 = m_def_RW_PBarColor1
   m_RW_PBarColor2 = m_def_RW_PBarColor2
   m_RW_PopInterval = m_def_RW_PopInterval
   m_RW_ShowLED = m_def_RW_ShowLED
   m_Theme = m_def_Theme
   m_UD_ArrowColor = m_def_UD_ArrowColor
   m_UD_AutoIncrement = m_def_UD_AutoIncrement
   m_UD_BorderMiddleOut = m_def_UD_BorderMiddleOut
   m_UD_BorderWidth = m_def_UD_BorderWidth
   m_UD_ButtonColor1 = m_def_UD_ButtonColor1
   m_UD_ButtonColor2 = m_def_UD_ButtonColor2
   m_UD_BorderColor1 = m_def_UD_BorderColor1
   m_UD_BorderColor2 = m_def_UD_BorderColor2
   m_UD_ButtonDownAngle = m_def_UD_ButtonDownAngle
   m_UD_ButtonDownMidOut = m_def_UD_ButtonDownMidOut
   m_UD_ButtonUpAngle = m_def_UD_ButtonUpAngle
   m_UD_ButtonUpMidOut = m_def_UD_ButtonUpMidOut
   m_UD_DisArrowColor = m_def_UD_DisArrowColor
   m_UD_DisBorderColor1 = m_def_UD_DisBorderColor1
   m_UD_DisBorderColor2 = m_def_UD_DisBorderColor2
   m_UD_DisButtonColor1 = m_def_UD_DisButtonColor1
   m_UD_DisButtonColor2 = m_def_UD_DisButtonColor2
   m_UD_FocusBorderColor1 = m_def_UD_FocusBorderColor1
   m_UD_FocusBorderColor2 = m_def_UD_FocusBorderColor2
   m_UD_IncrementInterval = m_def_UD_IncrementInterval
   m_UD_Orientation = m_def_UD_Orientation
   m_UD_ScrollDelay = m_def_UD_ScrollDelay
   m_UD_SwapDirections = m_def_UD_SwapDirections
   m_Value = m_def_Value
   m_ValueIncrCtrl = m_def_ValueIncrCtrl
   m_ValueIncrement = m_def_ValueIncrement
   m_ValueIncrShift = m_def_ValueIncrShift
   m_ValueIncrShiftCtrl = m_def_ValueIncrShiftCtrl
   m_ValueMax = m_def_ValueMax
   m_ValueMin = m_def_ValueMin
   m_Wrap = m_def_Wrap

'  initialize appropriate display colors.
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*************************************************************************
'* read the stored properties in the PropertyBag.                        *
'*************************************************************************

   With PropBag
      m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
      m_RW_BackAngle = .ReadProperty("RW_BackAngle", m_def_RW_BackAngle)
      m_RW_BackColor1 = .ReadProperty("RW_BackColor1", m_def_RW_BackColor1)
      m_RW_BackColor2 = .ReadProperty("RW_BackColor2", m_def_RW_BackColor2)
      m_RW_BackMiddleOut = .ReadProperty("RW_BackMiddleOut", m_def_RW_BackMiddleOut)
      m_RW_BorderColor1 = .ReadProperty("RW_BorderColor1", m_def_RW_BorderColor1)
      m_RW_BorderColor2 = .ReadProperty("RW_BorderColor2", m_def_RW_BorderColor2)
      m_RW_BorderMiddleOut = .ReadProperty("RW_BorderMiddleOut", m_def_RW_BorderMiddleOut)
      m_RW_BorderWidth = .ReadProperty("RW_BorderWidth", m_def_RW_BorderWidth)
      m_RW_GenerateEvent = .ReadProperty("RW_GenerateEvent", m_def_RW_GenerateEvent)
      m_RW_LED_BurnInColor = .ReadProperty("RW_LED_BurnInColor", m_def_RW_LED_BurnInColor)
      m_RW_LED_DigitColor = .ReadProperty("RW_LED_DigitColor", m_def_RW_LED_DigitColor)
      m_RW_LED_ShowBurnIn = .ReadProperty("RW_LED_ShowBurnIn", m_def_RW_LED_ShowBurnIn)
      m_RW_PBarColor1 = .ReadProperty("RW_PBarColor1", m_def_RW_PBarColor1)
      m_RW_PBarColor2 = .ReadProperty("RW_PBarColor2", m_def_RW_PBarColor2)
      m_RW_PopInterval = .ReadProperty("RW_PopInterval", m_def_RW_PopInterval)
      m_RW_ShowLED = .ReadProperty("RW_ShowLED", m_def_RW_ShowLED)
      m_Theme = .ReadProperty("Theme", m_def_Theme)
      m_UD_ArrowColor = .ReadProperty("UD_ArrowColor", m_def_UD_ArrowColor)
      m_UD_AutoIncrement = .ReadProperty("UD_AutoIncrement", m_def_UD_AutoIncrement)
      m_UD_BorderColor1 = .ReadProperty("UD_BorderColor1", m_def_UD_BorderColor1)
      m_UD_BorderColor2 = .ReadProperty("UD_BorderColor2", m_def_UD_BorderColor2)
      m_UD_BorderMiddleOut = .ReadProperty("UD_BorderMiddleOut", m_def_UD_BorderMiddleOut)
      m_UD_BorderWidth = .ReadProperty("UD_BorderWidth", m_def_UD_BorderWidth)
      m_UD_ButtonColor1 = .ReadProperty("UD_ButtonColor1", m_def_UD_ButtonColor1)
      m_UD_ButtonColor2 = .ReadProperty("UD_ButtonColor2", m_def_UD_ButtonColor2)
      m_UD_ButtonDownAngle = .ReadProperty("UD_ButtonDownAngle", m_def_UD_ButtonDownAngle)
      m_UD_ButtonDownMidOut = .ReadProperty("UD_ButtonDownMidOut", m_def_UD_ButtonDownMidOut)
      m_UD_ButtonUpAngle = .ReadProperty("UD_ButtonUpAngle", m_def_UD_ButtonUpAngle)
      m_UD_ButtonUpMidOut = .ReadProperty("UD_ButtonUpMidOut", m_def_UD_ButtonUpMidOut)
      m_UD_DisArrowColor = .ReadProperty("UD_DisArrowColor", m_def_UD_DisArrowColor)
      m_UD_DisBorderColor1 = .ReadProperty("UD_DisBorderColor1", m_def_UD_DisBorderColor1)
      m_UD_DisBorderColor2 = .ReadProperty("UD_DisBorderColor2", m_def_UD_DisBorderColor2)
      m_UD_DisButtonColor1 = .ReadProperty("UD_DisButtonColor1", m_def_UD_DisButtonColor1)
      m_UD_DisButtonColor2 = .ReadProperty("UD_DisButtonColor2", m_def_UD_DisButtonColor2)
      m_UD_FocusBorderColor1 = .ReadProperty("UD_FocusBorderColor1", m_def_UD_FocusBorderColor1)
      m_UD_FocusBorderColor2 = .ReadProperty("UD_FocusBorderColor2", m_def_UD_FocusBorderColor2)
      m_UD_IncrementInterval = .ReadProperty("UD_IncrementInterval", m_def_UD_IncrementInterval)
      m_UD_Orientation = .ReadProperty("UD_Orientation", m_def_UD_Orientation)
      m_UD_ScrollDelay = .ReadProperty("UD_ScrollDelay", m_def_UD_ScrollDelay)
      m_UD_SwapDirections = .ReadProperty("UD_SwapDirections", m_def_UD_SwapDirections)
      m_Value = .ReadProperty("Value", m_def_Value)
      m_ValueIncrCtrl = .ReadProperty("ValueIncrCtrl", m_def_ValueIncrCtrl)
      m_ValueIncrement = .ReadProperty("ValueIncrement", m_def_ValueIncrement)
      m_ValueIncrShift = .ReadProperty("ValueIncrShift", m_def_ValueIncrShift)
      m_ValueIncrShiftCtrl = .ReadProperty("ValueIncrShiftCtrl", m_def_ValueIncrShiftCtrl)
      m_ValueMin = .ReadProperty("ValueMin", m_def_ValueMin)
      m_ValueMax = .ReadProperty("ValueMax", m_def_ValueMax)
      m_Wrap = .ReadProperty("Wrap", m_def_Wrap)
   End With

'   RangeWindowWidth = 180
'   RangeWindowHeight = IIf(m_RW_ShowLED = True, 60, 40)

'  determine the taskbar height (in pixels).  This is used when positioning
'  the RangeWindow so that it is always fully visible on the screen.
   Call SystemParametersInfo(SPI_GETWORKAREA, 0&, ScreenWorkArea, 0&)
   TaskBarHeight = (Screen.Height / Screen.TwipsPerPixelY) - ScreenWorkArea.Bottom

'  if the .UD_AutoIncrement property is True, generate the appropriate increments.
   If m_UD_AutoIncrement Then
      CalculateIncrements
   End If

   InitializeControlGraphics

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'*************************************************************************
'* writes property values to PropertyBag.                                *
'*************************************************************************

   With PropBag
      .WriteProperty "Enabled", m_Enabled, m_def_Enabled
      .WriteProperty "RW_BackAngle", m_RW_BackAngle, m_def_RW_BackAngle
      .WriteProperty "RW_BackColor1", m_RW_BackColor1, m_def_RW_BackColor1
      .WriteProperty "RW_BackColor2", m_RW_BackColor2, m_def_RW_BackColor2
      .WriteProperty "RW_BackMiddleOut", m_RW_BackMiddleOut, m_def_RW_BackMiddleOut
      .WriteProperty "RW_BorderMiddleOut", m_RW_BorderMiddleOut, m_def_RW_BorderMiddleOut
      .WriteProperty "RW_BorderColor1", m_RW_BorderColor1, m_def_RW_BorderColor1
      .WriteProperty "RW_BorderColor2", m_RW_BorderColor2, m_def_RW_BorderColor2
      .WriteProperty "RW_BorderWidth", m_RW_BorderWidth, m_def_RW_BorderWidth
      .WriteProperty "RW_GenerateEvent", m_RW_GenerateEvent, m_def_RW_GenerateEvent
      .WriteProperty "RW_LED_BurnInColor", m_RW_LED_BurnInColor, m_def_RW_LED_BurnInColor
      .WriteProperty "RW_LED_DigitColor", m_RW_LED_DigitColor, m_def_RW_LED_DigitColor
      .WriteProperty "RW_LED_ShowBurnIn", m_RW_LED_ShowBurnIn, m_def_RW_LED_ShowBurnIn
      .WriteProperty "RW_PBarColor1", m_RW_PBarColor1, m_def_RW_PBarColor1
      .WriteProperty "RW_PBarColor2", m_RW_PBarColor2, m_def_RW_PBarColor2
      .WriteProperty "RW_PopInterval", m_RW_PopInterval, m_def_RW_PopInterval
      .WriteProperty "RW_ShowLED", m_RW_ShowLED, m_def_RW_ShowLED
      .WriteProperty "Theme", m_Theme, m_def_Theme
      .WriteProperty "UD_ArrowColor", m_UD_ArrowColor, m_def_UD_ArrowColor
      .WriteProperty "UD_AutoIncrement", m_UD_AutoIncrement, m_def_UD_AutoIncrement
      .WriteProperty "UD_BorderColor1", m_UD_BorderColor1, m_def_UD_BorderColor1
      .WriteProperty "UD_BorderColor2", m_UD_BorderColor2, m_def_UD_BorderColor2
      .WriteProperty "UD_BorderMiddleOut", m_UD_BorderMiddleOut, m_def_UD_BorderMiddleOut
      .WriteProperty "UD_BorderWidth", m_UD_BorderWidth, m_def_UD_BorderWidth
      .WriteProperty "UD_ButtonColor1", m_UD_ButtonColor1, m_def_UD_ButtonColor1
      .WriteProperty "UD_ButtonColor2", m_UD_ButtonColor2, m_def_UD_ButtonColor2
      .WriteProperty "UD_ButtonDownAngle", m_UD_ButtonDownAngle, m_def_UD_ButtonDownAngle
      .WriteProperty "UD_ButtonDownMidOut", m_UD_ButtonDownMidOut, m_def_UD_ButtonDownMidOut
      .WriteProperty "UD_ButtonUpAngle", m_UD_ButtonUpAngle, m_def_UD_ButtonUpAngle
      .WriteProperty "UD_ButtonUpMidOut", m_UD_ButtonUpMidOut, m_def_UD_ButtonUpMidOut
      .WriteProperty "UD_DisArrowColor", m_UD_DisArrowColor, m_def_UD_DisArrowColor
      .WriteProperty "UD_DisBorderColor1", m_UD_DisBorderColor1, m_def_UD_DisBorderColor1
      .WriteProperty "UD_DisBorderColor2", m_UD_DisBorderColor2, m_def_UD_DisBorderColor2
      .WriteProperty "UD_DisButtonColor1", m_UD_DisButtonColor1, m_def_UD_DisButtonColor1
      .WriteProperty "UD_DisButtonColor2", m_UD_DisButtonColor2, m_def_UD_DisButtonColor2
      .WriteProperty "UD_FocusBorderColor1", m_UD_FocusBorderColor1, m_def_UD_FocusBorderColor1
      .WriteProperty "UD_FocusBorderColor2", m_UD_FocusBorderColor2, m_def_UD_FocusBorderColor2
      .WriteProperty "UD_IncrementInterval", m_UD_IncrementInterval, m_def_UD_IncrementInterval
      .WriteProperty "UD_Orientation", m_UD_Orientation, m_def_UD_Orientation
      .WriteProperty "UD_ScrollDelay", m_UD_ScrollDelay, m_def_UD_ScrollDelay
      .WriteProperty "UD_SwapDirections", m_UD_SwapDirections, m_def_UD_SwapDirections
      .WriteProperty "Value", m_Value, m_def_Value
      .WriteProperty "ValueIncrCtrl", m_ValueIncrCtrl, m_def_ValueIncrCtrl
      .WriteProperty "ValueIncrement", m_ValueIncrement, m_def_ValueIncrement
      .WriteProperty "ValueIncrShift", m_ValueIncrShift, m_def_ValueIncrShift
      .WriteProperty "ValueIncrShiftCtrl", m_ValueIncrShiftCtrl, m_def_ValueIncrShiftCtrl
      .WriteProperty "ValueMax", m_ValueMax, m_def_ValueMax
      .WriteProperty "ValueMin", m_ValueMin, m_def_ValueMin
      .WriteProperty "Wrap", m_Wrap, m_def_Wrap
   End With

End Sub

Public Property Get Enabled() As Boolean
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal new_Enabled As Boolean)
   m_Enabled = new_Enabled
   If m_Enabled Then
      GetEnabledDisplayProperties
   Else
      GetDisabledDisplayProperties
   End If
   InitializeUpDownGraphics
   RedrawControl
   UserControl.Refresh
   PropertyChanged "Enabled"
End Property

Public Property Get hdc() As Long
   hdc = UserControl.hdc
End Property

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property

Public Property Get RW_BackAngle() As Single
Attribute RW_BackAngle.VB_Description = "RangeWindow background gradient angle, in degrees."
Attribute RW_BackAngle.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_BackAngle = m_RW_BackAngle
End Property

Public Property Let RW_BackAngle(ByVal New_RW_BackAngle As Single)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_BackAngle = New_RW_BackAngle
   PropertyChanged "RW_BackAngle"
End Property

Public Property Get RW_BackColor1() As OLE_COLOR
Attribute RW_BackColor1.VB_Description = "RangeWindow background first gradient color."
Attribute RW_BackColor1.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_BackColor1 = m_RW_BackColor1
End Property

Public Property Let RW_BackColor1(ByVal New_RW_BackColor1 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_BackColor1 = New_RW_BackColor1
   PropertyChanged "RW_BackColor1"
End Property

Public Property Get RW_BackColor2() As OLE_COLOR
Attribute RW_BackColor2.VB_Description = "RangeWindow background second gradient color."
Attribute RW_BackColor2.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_BackColor2 = m_RW_BackColor2
End Property

Public Property Let RW_BackColor2(ByVal New_RW_BackColor2 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_BackColor2 = New_RW_BackColor2
   PropertyChanged "RW_BackColor2"
End Property

Public Property Get RW_BackMiddleOut() As Boolean
Attribute RW_BackMiddleOut.VB_Description = "RangeWindow background gradient middle-out status."
Attribute RW_BackMiddleOut.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_BackMiddleOut = m_RW_BackMiddleOut
End Property

Public Property Let RW_BackMiddleOut(ByVal New_RW_BackMiddleOut As Boolean)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_BackMiddleOut = New_RW_BackMiddleOut
   PropertyChanged "RW_BackMiddleOut"
End Property

Public Property Get RW_BorderColor1() As OLE_COLOR
Attribute RW_BorderColor1.VB_Description = "RangeWindow border first gradient color."
Attribute RW_BorderColor1.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_BorderColor1 = m_RW_BorderColor1
End Property

Public Property Let RW_BorderColor1(ByVal New_RW_BorderColor1 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_BorderColor1 = New_RW_BorderColor1
   PropertyChanged "RW_BorderColor1"
End Property

Public Property Get RW_BorderColor2() As OLE_COLOR
Attribute RW_BorderColor2.VB_Description = "RangeWindow border second gradient color."
Attribute RW_BorderColor2.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_BorderColor2 = m_RW_BorderColor2
End Property

Public Property Let RW_BorderColor2(ByVal New_RW_BorderColor2 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_BorderColor2 = New_RW_BorderColor2
   PropertyChanged "RW_BorderColor2"
End Property

Public Property Get RW_BorderMiddleOut() As Boolean
Attribute RW_BorderMiddleOut.VB_Description = "RangeWindow border gradient middle-out status."
Attribute RW_BorderMiddleOut.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
'   If Ambient.UserMode Then Err.Raise 393    ' not available at all at runtime.
   RW_BorderMiddleOut = m_RW_BorderMiddleOut
End Property

Public Property Let RW_BorderMiddleOut(ByVal New_RW_BorderMiddleOut As Boolean)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_BorderMiddleOut = New_RW_BorderMiddleOut
   PropertyChanged "RW_BorderMiddleOut"
End Property

Public Property Get RW_BorderWidth() As Long
Attribute RW_BorderWidth.VB_Description = "Width, in pixels, of the RangeWindow border."
Attribute RW_BorderWidth.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_BorderWidth = m_RW_BorderWidth
End Property

Public Property Let RW_BorderWidth(ByVal New_RW_BorderWidth As Long)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_BorderWidth = New_RW_BorderWidth
   PropertyChanged "RW_BorderWidth"
End Property

Public Property Get RW_GenerateEvent() As Boolean
Attribute RW_GenerateEvent.VB_Description = "If True, Change event is thrown when mouse drags the progress bar and a MouseMove event is triggered.  Otherwise, a Change event is only thrown when a MouseUp occurs."
Attribute RW_GenerateEvent.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_GenerateEvent = m_RW_GenerateEvent
End Property

Public Property Let RW_GenerateEvent(ByVal New_RW_GenerateEvent As Boolean)
   m_RW_GenerateEvent = New_RW_GenerateEvent
   PropertyChanged "RW_GenerateEvent"
End Property

Public Property Get RW_LED_BurnInColor() As OLE_COLOR
Attribute RW_LED_BurnInColor.VB_Description = "Color of the simulated LED 'burned in' digits."
Attribute RW_LED_BurnInColor.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_LED_BurnInColor = m_RW_LED_BurnInColor
End Property

Public Property Let RW_LED_BurnInColor(ByVal New_RW_LED_BurnInColor As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_LED_BurnInColor = New_RW_LED_BurnInColor
   PropertyChanged "RW_LED_BurnInColor"
End Property

Public Property Get RW_LED_DigitColor() As OLE_COLOR
Attribute RW_LED_DigitColor.VB_Description = "Color of the RangeWindow LED digits."
Attribute RW_LED_DigitColor.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_LED_DigitColor = m_RW_LED_DigitColor
End Property

Public Property Let RW_LED_DigitColor(ByVal New_RW_LED_DigitColor As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_LED_DigitColor = New_RW_LED_DigitColor
   PropertyChanged "RW_LED_DigitColor"
End Property

Public Property Get RW_LED_ShowBurnIn() As Boolean
Attribute RW_LED_ShowBurnIn.VB_Description = "If True, simulated LED 'burned in' digits are displayed."
Attribute RW_LED_ShowBurnIn.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_LED_ShowBurnIn = m_RW_LED_ShowBurnIn
End Property

Public Property Let RW_LED_ShowBurnIn(ByVal New_RW_LED_ShowBurnIn As Boolean)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_LED_ShowBurnIn = New_RW_LED_ShowBurnIn
   PropertyChanged "RW_LED_ShowBurnIn"
End Property

Public Property Get RW_PBarColor1() As OLE_COLOR
Attribute RW_PBarColor1.VB_Description = "First gradient color of the RangeWindow progress bar."
Attribute RW_PBarColor1.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_PBarColor1 = m_RW_PBarColor1
End Property

Public Property Let RW_PBarColor1(ByVal New_RW_PBarColor1 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_PBarColor1 = New_RW_PBarColor1
   PropertyChanged "RW_PBarColor1"
End Property

Public Property Get RW_PBarColor2() As OLE_COLOR
Attribute RW_PBarColor2.VB_Description = "Second gradient color of the RangeWindow progress bar."
Attribute RW_PBarColor2.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_PBarColor2 = m_RW_PBarColor2
End Property

Public Property Let RW_PBarColor2(ByVal New_RW_PBarColor2 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_PBarColor2 = New_RW_PBarColor2
   PropertyChanged "RW_PBarColor2"
End Property

Public Property Get RW_PopInterval() As Long
Attribute RW_PopInterval.VB_Description = "The time, in milliseconds, between an UpDown button being held down and the appearance of the RangeWindow."
Attribute RW_PopInterval.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_PopInterval = m_RW_PopInterval
End Property

Public Property Let RW_PopInterval(ByVal New_RW_PopInterval As Long)
   m_RW_PopInterval = New_RW_PopInterval
   PropertyChanged "RW_PopInterval"
End Property

Public Property Get RW_ShowLED() As Boolean
Attribute RW_ShowLED.VB_Description = "If True, the LED .Value display is shown in the RangeWindow.  If False, only the ProgressBar is shown."
Attribute RW_ShowLED.VB_ProcData.VB_Invoke_Property = ";RangeWindow"
   RW_ShowLED = m_RW_ShowLED
End Property

Public Property Let RW_ShowLED(ByVal New_RW_ShowLED As Boolean)
'  this property is useful for situations where you already have another control displaying the
'  .Value as RangeWindow progressbar is being dragged.  In this case you may feel that the LED display
'  is redundant.  If so, just set this property to False and only the progressbar will be shown.
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_RW_ShowLED = New_RW_ShowLED
   PropertyChanged "RW_ShowLED"
End Property

Public Property Get Theme() As MRR_ThemeOptions
Attribute Theme.VB_Description = "One of eight predefined color schemes for the control."
Attribute Theme.VB_ProcData.VB_Invoke_Property = ";General"
   Theme = m_Theme
End Property

Public Property Let Theme(ByVal New_Theme As MRR_ThemeOptions)

   m_Theme = New_Theme

   Select Case m_Theme

      Case [Red Rum]
         m_RW_BackAngle = 90
         m_RW_BackColor1 = &H40&
         m_RW_BackColor2 = &HC0&
         m_RW_BackMiddleOut = True
         m_RW_BorderColor1 = &H40&
         m_RW_BorderColor2 = &H8080FF
         m_RW_BorderMiddleOut = True
         m_RW_BorderWidth = 8
         m_RW_LED_BurnInColor = &H80&
         m_RW_LED_DigitColor = &H8080FF
         m_RW_LED_ShowBurnIn = True
         m_RW_PBarColor1 = &H40&
         m_RW_PBarColor2 = &H8080FF
         m_UD_ArrowColor = &HC0C0FF
         m_UD_BorderColor1 = &H40&
         m_UD_BorderColor2 = &H8080FF
         m_UD_BorderMiddleOut = True
         m_UD_BorderWidth = 8
         m_UD_ButtonColor1 = &H40&
         m_UD_ButtonColor2 = &H8080FF
         m_UD_ButtonDownAngle = 90
         m_UD_ButtonDownMidOut = False
         m_UD_ButtonUpAngle = 90
         m_UD_ButtonUpMidOut = True
         m_UD_DisArrowColor = &H909090
         m_UD_DisBorderColor1 = &H808080
         m_UD_DisBorderColor2 = &HE0E0E0
         m_UD_DisButtonColor1 = &H808080
         m_UD_DisButtonColor2 = &HE0E0E0
         m_UD_FocusBorderColor1 = &H40&
         m_UD_FocusBorderColor2 = &H4040FF
         m_UD_Orientation = [Vertical]
         m_UD_SwapDirections = False
         
      Case [Gunmetal Grey]
         m_RW_BackAngle = 90
         m_RW_BackColor1 = &H0&
         m_RW_BackColor2 = &H404040
         m_RW_BackMiddleOut = True
         m_RW_BorderColor1 = &H0
         m_RW_BorderColor2 = &H808080
         m_RW_BorderMiddleOut = True
         m_RW_BorderWidth = 8
         m_RW_LED_BurnInColor = &H404040
         m_RW_LED_DigitColor = &HE0E0E0
         m_RW_LED_ShowBurnIn = True
         m_RW_PBarColor1 = &H0
         m_RW_PBarColor2 = &HE0E0E0
         m_UD_ArrowColor = &HFFFFFF
         m_UD_BorderColor1 = &H0
         m_UD_BorderColor2 = &HC0C0C0
         m_UD_BorderMiddleOut = True
         m_UD_BorderWidth = 8
         m_UD_ButtonColor1 = &H0
         m_UD_ButtonColor2 = &HC0C0C0
         m_UD_ButtonDownAngle = 90
         m_UD_ButtonDownMidOut = False
         m_UD_ButtonUpAngle = 90
         m_UD_ButtonUpMidOut = True
         m_UD_DisArrowColor = &H909090
         m_UD_DisBorderColor1 = &H808080
         m_UD_DisBorderColor2 = &HE0E0E0
         m_UD_DisButtonColor1 = &H808080
         m_UD_DisButtonColor2 = &HE0E0E0
         m_UD_FocusBorderColor1 = &H0
         m_UD_FocusBorderColor2 = &H808080
         m_UD_Orientation = [Vertical]
         m_UD_SwapDirections = False
   
      Case [Green With Envy]
         m_RW_BackAngle = 90
         m_RW_BackColor1 = &H4000&
         m_RW_BackColor2 = &HA000&
         m_RW_BackMiddleOut = True
         m_RW_BorderColor1 = &H4000&
         m_RW_BorderColor2 = &H80FF80
         m_RW_BorderMiddleOut = True
         m_RW_BorderWidth = 8
         m_RW_LED_BurnInColor = &H8000&
         m_RW_LED_DigitColor = &H80FF80
         m_RW_LED_ShowBurnIn = True
         m_RW_PBarColor1 = &H4000&
         m_RW_PBarColor2 = &H80FF80
         m_UD_ArrowColor = &HC0FFC0
         m_UD_BorderColor1 = &H4000&
         m_UD_BorderColor2 = &H80FF80
         m_UD_BorderMiddleOut = True
         m_UD_BorderWidth = 8
         m_UD_ButtonColor1 = &H4000&
         m_UD_ButtonColor2 = &HFF00&
         m_UD_ButtonDownAngle = 90
         m_UD_ButtonDownMidOut = False
         m_UD_ButtonUpAngle = 90
         m_UD_ButtonUpMidOut = True
         m_UD_DisArrowColor = &H909090
         m_UD_DisBorderColor1 = &H808080
         m_UD_DisBorderColor2 = &HE0E0E0
         m_UD_DisButtonColor1 = &H808080
         m_UD_DisButtonColor2 = &HE0E0E0
         m_UD_FocusBorderColor1 = &H4000&
         m_UD_FocusBorderColor2 = &HC000&
         m_UD_Orientation = [Vertical]
         m_UD_SwapDirections = False
   
      Case [Purple People Eater]
         m_RW_BackAngle = 90
         m_RW_BackColor1 = &H400040
         m_RW_BackColor2 = &HC000C0
         m_RW_BackMiddleOut = True
         m_RW_BorderColor1 = &H400040
         m_RW_BorderColor2 = &HFF80FF
         m_RW_BorderMiddleOut = True
         m_RW_BorderWidth = 8
         m_RW_LED_BurnInColor = &H800080
         m_RW_LED_DigitColor = &HFF80FF
         m_RW_LED_ShowBurnIn = True
         m_RW_PBarColor1 = &H400040
         m_RW_PBarColor2 = &HFF80FF
         m_UD_ArrowColor = &H800080
         m_UD_BorderColor1 = &H400040
         m_UD_BorderColor2 = &HFF80FF
         m_UD_BorderMiddleOut = True
         m_UD_BorderWidth = 8
         m_UD_ButtonColor1 = &H400040
         m_UD_ButtonColor2 = &HFF80FF
         m_UD_ButtonDownAngle = 90
         m_UD_ButtonDownMidOut = False
         m_UD_ButtonUpAngle = 90
         m_UD_ButtonUpMidOut = True
         m_UD_DisArrowColor = &H909090
         m_UD_DisBorderColor1 = &H808080
         m_UD_DisBorderColor2 = &HE0E0E0
         m_UD_DisButtonColor1 = &H808080
         m_UD_DisButtonColor2 = &HE0E0E0
         m_UD_FocusBorderColor1 = &H400040
         m_UD_FocusBorderColor2 = &HC000C0
         m_UD_Orientation = [Vertical]
         m_UD_SwapDirections = False
   
      Case [Penny Wise]
         m_RW_BackAngle = 90
         m_RW_BackColor1 = &H404080
         m_RW_BackColor2 = &H60C0&
         m_RW_BackMiddleOut = True
         m_RW_BorderColor1 = &H404080
         m_RW_BorderColor2 = &H80C0FF
         m_RW_BorderMiddleOut = True
         m_RW_BorderWidth = 8
         m_RW_LED_BurnInColor = &H4080&
         m_RW_LED_DigitColor = &H80C0FF
         m_RW_LED_ShowBurnIn = True
         m_RW_PBarColor1 = &H404080
         m_RW_PBarColor2 = &H80C0FF
         m_UD_ArrowColor = &H4080&
         m_UD_BorderColor1 = &H404080
         m_UD_BorderColor2 = &H80C0FF
         m_UD_BorderMiddleOut = True
         m_UD_BorderWidth = 8
         m_UD_ButtonColor1 = &H404080
         m_UD_ButtonColor2 = &H80C0FF
         m_UD_ButtonDownAngle = 90
         m_UD_ButtonDownMidOut = False
         m_UD_ButtonUpAngle = 90
         m_UD_ButtonUpMidOut = True
         m_UD_DisArrowColor = &H909090
         m_UD_DisBorderColor1 = &H808080
         m_UD_DisBorderColor2 = &HE0E0E0
         m_UD_DisButtonColor1 = &H808080
         m_UD_DisButtonColor2 = &HE0E0E0
         m_UD_FocusBorderColor1 = &H404080
         m_UD_FocusBorderColor2 = &H80FF&
         m_UD_Orientation = [Vertical]
         m_UD_SwapDirections = False

      Case [Cyan Eyed]
         m_RW_BackAngle = 90
         m_RW_BackColor1 = &H404000
         m_RW_BackColor2 = &HC0C000
         m_RW_BackMiddleOut = True
         m_RW_BorderColor1 = &H404000
         m_RW_BorderColor2 = &HFFFF80
         m_RW_BorderMiddleOut = True
         m_RW_BorderWidth = 8
         m_RW_LED_BurnInColor = &H808000
         m_RW_LED_DigitColor = &HFFFF80
         m_RW_LED_ShowBurnIn = True
         m_RW_PBarColor1 = &H404000
         m_RW_PBarColor2 = &HFFFF80
         m_UD_ArrowColor = &H808000
         m_UD_BorderColor1 = &H404000
         m_UD_BorderColor2 = &HFFFF80
         m_UD_BorderMiddleOut = True
         m_UD_BorderWidth = 8
         m_UD_ButtonColor1 = &H404000
         m_UD_ButtonColor2 = &HFFFF80
         m_UD_ButtonDownAngle = 90
         m_UD_ButtonDownMidOut = False
         m_UD_ButtonUpAngle = 90
         m_UD_ButtonUpMidOut = True
         m_UD_DisArrowColor = &H909090
         m_UD_DisBorderColor1 = &H808080
         m_UD_DisBorderColor2 = &HE0E0E0
         m_UD_DisButtonColor1 = &H808080
         m_UD_DisButtonColor2 = &HE0E0E0
         m_UD_FocusBorderColor1 = &H404000
         m_UD_FocusBorderColor2 = &HC0C000
         m_UD_Orientation = [Vertical]
         m_UD_SwapDirections = False

      Case [Blue Moon]
         m_RW_BackAngle = 90
         m_RW_BackColor1 = &H400000
         m_RW_BackColor2 = &HC00000
         m_RW_BackMiddleOut = True
         m_RW_BorderColor1 = &H400000
         m_RW_BorderColor2 = &HFF8080
         m_RW_BorderMiddleOut = True
         m_RW_BorderWidth = 8
         m_RW_LED_BurnInColor = &H800000
         m_RW_LED_DigitColor = &HFF8080
         m_RW_LED_ShowBurnIn = True
         m_RW_PBarColor1 = &H400000
         m_RW_PBarColor2 = &HFF8080
         m_UD_ArrowColor = &HFFC0C0
         m_UD_BorderColor1 = &H800000
         m_UD_BorderColor2 = &HFF8080
         m_UD_BorderMiddleOut = True
         m_UD_BorderWidth = 8
         m_UD_ButtonColor1 = &H400000
         m_UD_ButtonColor2 = &HFF8080
         m_UD_ButtonDownAngle = 90
         m_UD_ButtonDownMidOut = False
         m_UD_ButtonUpAngle = 90
         m_UD_ButtonUpMidOut = True
         m_UD_DisArrowColor = &H909090
         m_UD_DisBorderColor1 = &H808080
         m_UD_DisBorderColor2 = &HE0E0E0
         m_UD_DisButtonColor1 = &H808080
         m_UD_DisButtonColor2 = &HE0E0E0
         m_UD_FocusBorderColor1 = &H400000
         m_UD_FocusBorderColor2 = &HFF0000
         m_UD_Orientation = [Vertical]
         m_UD_SwapDirections = False

      Case [Golden Goose]
         m_RW_BackAngle = 90
         m_RW_BackColor1 = &H4040&
         m_RW_BackColor2 = &H8080&
         m_RW_BackMiddleOut = True
         m_RW_BorderColor1 = &H4040&
         m_RW_BorderColor2 = &H80FFFF
         m_RW_BorderMiddleOut = True
         m_RW_BorderWidth = 8
         m_RW_LED_BurnInColor = &H8080&
         m_RW_LED_DigitColor = &H80FFFF
         m_RW_LED_ShowBurnIn = True
         m_RW_PBarColor1 = &H4040&
         m_RW_PBarColor2 = &H80FFFF
         m_UD_ArrowColor = &H8080&
         m_UD_BorderColor1 = &H4040&
         m_UD_BorderColor2 = &H80FFFF
         m_UD_BorderMiddleOut = True
         m_UD_BorderWidth = 8
         m_UD_ButtonColor1 = &H4040
         m_UD_ButtonColor2 = &H80FFFF
         m_UD_ButtonDownAngle = 90
         m_UD_ButtonDownMidOut = False
         m_UD_ButtonUpAngle = 90
         m_UD_ButtonUpMidOut = True
         m_UD_DisArrowColor = &H909090
         m_UD_DisBorderColor1 = &H808080
         m_UD_DisBorderColor2 = &HE0E0E0
         m_UD_DisButtonColor1 = &H808080
         m_UD_DisButtonColor2 = &HE0E0E0
         m_UD_FocusBorderColor1 = &H4040&
         m_UD_FocusBorderColor2 = &HC0C0&
         m_UD_Orientation = [Vertical]
         m_UD_SwapDirections = False

   End Select

   InitializeControlGraphics
   RedrawControl
   UserControl.Refresh
   PropertyChanged "Theme"

End Property

Public Property Get UD_ArrowColor() As OLE_COLOR
Attribute UD_ArrowColor.VB_Description = "Arrow color for UpDown buttons."
Attribute UD_ArrowColor.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_ArrowColor = m_UD_ArrowColor
End Property

Public Property Let UD_ArrowColor(ByVal New_UD_ArrowColor As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_ArrowColor = New_UD_ArrowColor
   PropertyChanged "UD_ArrowColor"
End Property

Public Property Get UD_AutoIncrement() As Boolean
Attribute UD_AutoIncrement.VB_Description = "If True, sets increments: regular increment=1, CtrlIncr=1% of range, ShiftIncr=5% of range, CtrlShiftIncr = 10% of range."
Attribute UD_AutoIncrement.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_AutoIncrement = m_UD_AutoIncrement
End Property

Public Property Let UD_AutoIncrement(ByVal New_UD_AutoIncrement As Boolean)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_AutoIncrement = New_UD_AutoIncrement
   PropertyChanged "UD_AutoIncrement"
End Property

Public Property Get UD_BorderColor1() As OLE_COLOR
Attribute UD_BorderColor1.VB_Description = "UpDown border first gradient color."
Attribute UD_BorderColor1.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_BorderColor1 = m_UD_BorderColor1
End Property

Public Property Let UD_BorderColor1(ByVal New_UD_BorderColor1 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_BorderColor1 = New_UD_BorderColor1
   PropertyChanged "UD_BorderColor1"
End Property

Public Property Get UD_BorderColor2() As OLE_COLOR
Attribute UD_BorderColor2.VB_Description = "UpDown border second gradient color."
Attribute UD_BorderColor2.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_BorderColor2 = m_UD_BorderColor2
End Property

Public Property Let UD_BorderColor2(ByVal New_UD_BorderColor2 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_BorderColor2 = New_UD_BorderColor2
   PropertyChanged "UD_BorderColor2"
End Property

Public Property Get UD_BorderMiddleOut() As Boolean
Attribute UD_BorderMiddleOut.VB_Description = "UpDown gradient border middle-out status."
Attribute UD_BorderMiddleOut.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_BorderMiddleOut = m_UD_BorderMiddleOut
End Property

Public Property Let UD_BorderMiddleOut(ByVal New_UD_BorderMiddleOut As Boolean)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_BorderMiddleOut = New_UD_BorderMiddleOut
   PropertyChanged "UD_BorderMiddleOut"
End Property

Public Property Get UD_BorderWidth() As Long
Attribute UD_BorderWidth.VB_Description = "Width, in pixels, of UpDown border."
Attribute UD_BorderWidth.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_BorderWidth = m_UD_BorderWidth
End Property

Public Property Let UD_BorderWidth(ByVal New_UD_BorderWidth As Long)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_BorderWidth = New_UD_BorderWidth
   PropertyChanged "UD_BorderWidth"
End Property

Public Property Get UD_ButtonColor1() As OLE_COLOR
Attribute UD_ButtonColor1.VB_Description = "First gradient color of UpDown buttons."
Attribute UD_ButtonColor1.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_ButtonColor1 = m_UD_ButtonColor1
End Property

Public Property Let UD_ButtonColor1(ByVal New_UD_ButtonColor1 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_ButtonColor1 = New_UD_ButtonColor1
   PropertyChanged "UD_ButtonColor1"
End Property

Public Property Get UD_ButtonColor2() As OLE_COLOR
Attribute UD_ButtonColor2.VB_Description = "Second gradient color of UpDown buttons."
Attribute UD_ButtonColor2.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_ButtonColor2 = m_UD_ButtonColor2
End Property

Public Property Let UD_ButtonColor2(ByVal New_UD_ButtonColor2 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_ButtonColor2 = New_UD_ButtonColor2
   PropertyChanged "UD_ButtonColor2"
End Property

Public Property Get UD_ButtonDownAngle() As Single
Attribute UD_ButtonDownAngle.VB_Description = "Angle in degrees of UpDown buttons when they are clicked down."
Attribute UD_ButtonDownAngle.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_ButtonDownAngle = m_UD_ButtonDownAngle
End Property

Public Property Let UD_ButtonDownAngle(ByVal New_UD_ButtonDownAngle As Single)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_ButtonDownAngle = New_UD_ButtonDownAngle
   PropertyChanged "UD_ButtonDownAngle"
End Property

Public Property Get UD_ButtonDownMidOut() As Boolean
Attribute UD_ButtonDownMidOut.VB_Description = "Gradient middle-out status of UpDown button when button is clicked down."
Attribute UD_ButtonDownMidOut.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_ButtonDownMidOut = m_UD_ButtonDownMidOut
End Property

Public Property Let UD_ButtonDownMidOut(ByVal New_UD_ButtonDownMidOut As Boolean)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_ButtonDownMidOut = New_UD_ButtonDownMidOut
   PropertyChanged "UD_ButtonDownMidOut"
End Property

Public Property Get UD_ButtonUpAngle() As Single
Attribute UD_ButtonUpAngle.VB_Description = "Gradient angle in degrees when button is in its unclicked state."
Attribute UD_ButtonUpAngle.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_ButtonUpAngle = m_UD_ButtonUpAngle
End Property

Public Property Let UD_ButtonUpAngle(ByVal New_UD_ButtonUpAngle As Single)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_ButtonUpAngle = New_UD_ButtonUpAngle
   PropertyChanged "UD_ButtonUpAngle"
End Property

Public Property Get UD_ButtonUpMidOut() As Boolean
Attribute UD_ButtonUpMidOut.VB_Description = "Gradient middle-out status when button is in its unclicked state."
Attribute UD_ButtonUpMidOut.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_ButtonUpMidOut = m_UD_ButtonUpMidOut
End Property

Public Property Let UD_ButtonUpMidOut(ByVal New_UD_ButtonUpMidOut As Boolean)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_ButtonUpMidOut = New_UD_ButtonUpMidOut
   PropertyChanged "UD_ButtonUpMidOut"
End Property

Public Property Get UD_DisArrowColor() As OLE_COLOR
Attribute UD_DisArrowColor.VB_Description = "Button arrow color when control is disabled."
Attribute UD_DisArrowColor.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_DisArrowColor = m_UD_DisArrowColor
End Property

Public Property Let UD_DisArrowColor(ByVal New_UD_DisArrowColor As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_DisArrowColor = New_UD_DisArrowColor
   PropertyChanged "UD_DisArrowColor"
End Property

Public Property Get UD_DisBorderColor1() As OLE_COLOR
Attribute UD_DisBorderColor1.VB_Description = "UpDown border first gradient color when control is disabled."
Attribute UD_DisBorderColor1.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_DisBorderColor1 = m_UD_DisBorderColor1
End Property

Public Property Let UD_DisBorderColor1(ByVal New_UD_DisBorderColor1 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_DisBorderColor1 = New_UD_DisBorderColor1
   PropertyChanged "UD_DisBorderColor1"
End Property

Public Property Get UD_DisBorderColor2() As OLE_COLOR
Attribute UD_DisBorderColor2.VB_Description = "UpDown border second gradient color when control is disabled."
Attribute UD_DisBorderColor2.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_DisBorderColor2 = m_UD_DisBorderColor2
End Property

Public Property Let UD_DisBorderColor2(ByVal New_UD_DisBorderColor2 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_DisBorderColor2 = New_UD_DisBorderColor2
   PropertyChanged "UD_DisBorderColor2"
End Property

Public Property Get UD_DisButtonColor1() As OLE_COLOR
Attribute UD_DisButtonColor1.VB_Description = "UpDown button first gradient color when control is disabled."
Attribute UD_DisButtonColor1.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_DisButtonColor1 = m_UD_DisButtonColor1
End Property

Public Property Let UD_DisButtonColor1(ByVal New_UD_DisButtonColor1 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_DisButtonColor1 = New_UD_DisButtonColor1
   PropertyChanged "UD_DisButtonColor1"
End Property

Public Property Get UD_DisButtonColor2() As OLE_COLOR
Attribute UD_DisButtonColor2.VB_Description = "UpDown button second gradient color when control is disabled."
Attribute UD_DisButtonColor2.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_DisButtonColor2 = m_UD_DisButtonColor2
End Property

Public Property Let UD_DisButtonColor2(ByVal New_UD_DisButtonColor2 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_DisButtonColor2 = New_UD_DisButtonColor2
   PropertyChanged "UD_DisButtonColor2"
End Property

Public Property Get UD_FocusBorderColor1() As OLE_COLOR
Attribute UD_FocusBorderColor1.VB_Description = "UpDown border first gradient color when control has the focus."
Attribute UD_FocusBorderColor1.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_FocusBorderColor1 = m_UD_FocusBorderColor1
End Property

Public Property Let UD_FocusBorderColor1(ByVal New_UD_FocusBorderColor1 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_FocusBorderColor1 = New_UD_FocusBorderColor1
   PropertyChanged "UD_FocusBorderColor1"
End Property

Public Property Get UD_FocusBorderColor2() As OLE_COLOR
Attribute UD_FocusBorderColor2.VB_Description = "UpDown border second gradient color when control has the focus."
Attribute UD_FocusBorderColor2.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_FocusBorderColor2 = m_UD_FocusBorderColor2
End Property

Public Property Let UD_FocusBorderColor2(ByVal New_UD_FocusBorderColor2 As OLE_COLOR)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_FocusBorderColor2 = New_UD_FocusBorderColor2
   PropertyChanged "UD_FocusBorderColor2"
End Property

Public Property Get UD_IncrementInterval() As Long
Attribute UD_IncrementInterval.VB_Description = "Time, in milliseconds, between value changes when UpDown button is held down."
Attribute UD_IncrementInterval.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_IncrementInterval = m_UD_IncrementInterval
End Property

Public Property Let UD_IncrementInterval(ByVal New_UD_IncrementInterval As Long)
   m_UD_IncrementInterval = New_UD_IncrementInterval
   PropertyChanged "UD_IncrementInterval"
End Property

Public Property Get UD_Orientation() As MRR_Orientation
Attribute UD_Orientation.VB_Description = "Orientation of UpDown - vertical or horizontal."
Attribute UD_Orientation.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_Orientation = m_UD_Orientation
End Property

Public Property Let UD_Orientation(ByVal New_UD_Orientation As MRR_Orientation)
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_Orientation = New_UD_Orientation
   InitializeUpDownGraphics
   RedrawControl
   UserControl.Refresh
   PropertyChanged "UD_Orientation"
End Property

Public Property Get UD_ScrollDelay() As Long
Attribute UD_ScrollDelay.VB_Description = "The delay, in milliseconds, between the time an UpDown button is first held down and the start of value scrolling."
Attribute UD_ScrollDelay.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_ScrollDelay = m_UD_ScrollDelay
End Property

Public Property Let UD_ScrollDelay(ByVal New_UD_ScrollDelay As Long)
   m_UD_ScrollDelay = New_UD_ScrollDelay
   PropertyChanged "UD_ScrollDelay"
End Property

Public Property Get UD_SwapDirections() As Boolean
Attribute UD_SwapDirections.VB_Description = "If True, swaps UpDown button  increments - i.e., Up decrements and Down increments the .Value property.  Useful in navigating images."
Attribute UD_SwapDirections.VB_ProcData.VB_Invoke_Property = ";UpDown"
   UD_SwapDirections = m_UD_SwapDirections
End Property

Public Property Let UD_SwapDirections(ByVal New_UD_SwapDirections As Boolean)
'  By default, in a vertical UpDown the Up button increments the .Value and the Down
'  button decrements the .Value.  Likewise, in a horizontal updown the Left button decrements and
'  the Right button increments the .Value.  However, if .UD_SwapDirections is True, the buttons
'  swap increment/decrement.  Why did I include this?  Because I was working on a demo app that
'  used a vertical and a horizontal updown to navigate through an image.  The horizontal updown
'  worked fine - the image would move to the left when I clicked the left button and right when I
'  clicked the right button.  However in the vertical updown, clicking the up button would move the
'  image DOWN and the down button would move the image UP.  That's because I used the updown to send
'  the Y coordinate of the top left corner of the portion of the image I wanted displayed.  Click
'  the UP button and the Y coordinate increases, and the picture therefore moves DOWN.  Not too
'  logical in this use of the RangeRoamer control so I added this property for these types of situations.
   If Ambient.UserMode Then Err.Raise 382    ' not available at runtime.
   m_UD_SwapDirections = New_UD_SwapDirections
   PropertyChanged "UD_SwapDirections"
End Property

Public Property Get Value() As Long
Attribute Value.VB_Description = "The value being manipulated by the UpDown and RangeWindow."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";General"
   Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
   m_Value = New_Value
   PropertyChanged "Value"
End Property

Public Property Get ValueIncrCtrl() As Long
Attribute ValueIncrCtrl.VB_Description = "The value increment when the Ctrl key is held down while clicking an UpDown button."
Attribute ValueIncrCtrl.VB_ProcData.VB_Invoke_Property = ";General"
   ValueIncrCtrl = m_ValueIncrCtrl
End Property

Public Property Let ValueIncrCtrl(ByVal New_ValueIncrCtrl As Long)
   m_ValueIncrCtrl = New_ValueIncrCtrl
   PropertyChanged "ValueIncrCtrl"
End Property

Public Property Get ValueIncrement() As Long
Attribute ValueIncrement.VB_Description = "The value increment when no keys are held down while clicking an UpDown button."
Attribute ValueIncrement.VB_ProcData.VB_Invoke_Property = ";General"
   ValueIncrement = m_ValueIncrement
End Property

Public Property Let ValueIncrement(ByVal New_ValueIncrement As Long)
   m_ValueIncrement = New_ValueIncrement
   PropertyChanged "ValueIncrement"
End Property

Public Property Get ValueIncrShift() As Long
Attribute ValueIncrShift.VB_Description = "The value increment when the Shift key is held down while clicking an UpDown button."
Attribute ValueIncrShift.VB_ProcData.VB_Invoke_Property = ";General"
   ValueIncrShift = m_ValueIncrShift
End Property

Public Property Let ValueIncrShift(ByVal New_ValueIncrShift As Long)
   m_ValueIncrShift = New_ValueIncrShift
   PropertyChanged "ValueIncrShift"
End Property

Public Property Get ValueIncrShiftCtrl() As Long
Attribute ValueIncrShiftCtrl.VB_Description = "The value increment when the Ctrl and Shift keys are held down while clicking an UpDown button."
Attribute ValueIncrShiftCtrl.VB_ProcData.VB_Invoke_Property = ";General"
   ValueIncrShiftCtrl = m_ValueIncrShiftCtrl
End Property

Public Property Let ValueIncrShiftCtrl(ByVal New_ValueIncrShiftCtrl As Long)
   m_ValueIncrShiftCtrl = New_ValueIncrShiftCtrl
   PropertyChanged "ValueIncrShiftCtrl"
End Property

Public Property Get ValueMax() As Long
Attribute ValueMax.VB_Description = "The upper value in the value range."
Attribute ValueMax.VB_ProcData.VB_Invoke_Property = ";General"
   ValueMax = m_ValueMax
End Property

Public Property Let ValueMax(ByVal New_ValueMax As Long)
   m_ValueMax = New_ValueMax
   If m_UD_AutoIncrement Then
      CalculateIncrements
   End If
   PropertyChanged "ValueMax"
End Property

Public Property Get ValueMin() As Long
Attribute ValueMin.VB_Description = "The lower value in the value range."
Attribute ValueMin.VB_ProcData.VB_Invoke_Property = ";General"
   ValueMin = m_ValueMin
End Property

Public Property Let ValueMin(ByVal New_ValueMin As Long)
   m_ValueMin = New_ValueMin
   If m_UD_AutoIncrement Then
      CalculateIncrements
   End If
   PropertyChanged "ValueMin"
End Property

Public Property Get Wrap() As Boolean
Attribute Wrap.VB_Description = "If True, values wrap around back to the beginning or end when ValueMin or ValueMax values are exceeded while clicking UpDown buttons."
Attribute Wrap.VB_ProcData.VB_Invoke_Property = ";General"
   Wrap = m_Wrap
End Property

Public Property Let Wrap(ByVal New_Wrap As Boolean)
   m_Wrap = New_Wrap
   PropertyChanged "Wrap"
End Property
