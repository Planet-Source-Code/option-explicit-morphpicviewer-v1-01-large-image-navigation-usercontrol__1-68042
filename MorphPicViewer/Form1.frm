VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MorphPicViewer - Matthew R. Usner"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin Project1.MorphContainer MorphContainer1 
      Height          =   6135
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   10821
      BackAngle       =   65
      IconSize        =   0
      BackColor2      =   12632256
      BackColor1      =   0
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      HeaderVisible   =   0   'False
      Begin Project1.MorphContainer MorphContainer2 
         Height          =   5775
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   10186
         BackAngle       =   45
         IconSize        =   0
         HeaderColor2    =   8421504
         HeaderColor1    =   0
         BackColor2      =   12632256
         BackColor1      =   0
         BorderColor     =   0
         CaptionColor    =   65535
         Caption         =   "Navigation Demo"
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Project1.MorphOptionCheck optPicStyle 
            Height          =   375
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BackColor1      =   4210752
            BackColor2      =   14737632
            Caption         =   "Normal Size"
            CheckBoxAngle   =   45
            CheckBoxColor1  =   4210752
            CheckBoxColor2  =   16777215
            CheckBoxMiddleOut=   0   'False
            CheckColor      =   0
            ControlType     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseOverActions=   3
            MOverBorderColor=   16776960
            MOverCheckBoxColor=   16776960
            ShowFocusRect   =   0   'False
            Value           =   -1
         End
         Begin Project1.MorphPicViewer mpvNavigate 
            Height          =   2655
            Left            =   120
            TabIndex        =   0
            Top             =   720
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   4683
            BorderWidth     =   8
            DisBorderBGColor1=   8421504
            DisBorderBGColor2=   14737632
            PicHeight       =   768
            Picture         =   "Form1.frx":0000
            PicWidth        =   1024
         End
         Begin Project1.MorphRangeRoamer mrrVertical 
            Height          =   720
            Left            =   600
            TabIndex        =   3
            Top             =   3480
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   1270
            UD_BorderWidth  =   4
            UD_IncrementInterval=   75
            UD_ScrollDelay  =   100
            UD_SwapDirections=   -1  'True
            ValueIncrCtrl   =   5
            ValueIncrement  =   5
            ValueIncrShift  =   5
            ValueIncrShiftCtrl=   5
         End
         Begin Project1.MorphRangeRoamer mrrHorizontal 
            Height          =   360
            Left            =   3240
            TabIndex        =   4
            Top             =   3665
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   635
            UD_BorderWidth  =   4
            UD_ButtonDownAngle=   180
            UD_ButtonUpAngle=   180
            UD_IncrementInterval=   75
            UD_Orientation  =   1
            UD_ScrollDelay  =   100
            ValueIncrCtrl   =   5
            ValueIncrement  =   5
            ValueIncrShift  =   5
            ValueIncrShiftCtrl=   5
         End
         Begin Project1.MorphOptionCheck chkKeepAspectRatio 
            Height          =   375
            Left            =   360
            TabIndex        =   6
            Top             =   5160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BackColor1      =   4210752
            BackColor2      =   14737632
            Caption         =   "Aspect Ratio"
            CheckBoxAngle   =   45
            CheckBoxColor1  =   4210752
            CheckBoxColor2  =   16777215
            CheckBoxMiddleOut=   0   'False
            CheckColor      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseOverActions=   3
            MOverBorderColor=   16776960
            MOverCheckBoxColor=   16776960
            ShowFocusRect   =   0   'False
            Value           =   1
         End
         Begin Project1.MorphOptionCheck optPicStyle 
            Height          =   375
            Index           =   1
            Left            =   2760
            TabIndex        =   7
            Top             =   4560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BackColor1      =   4210752
            BackColor2      =   14737632
            Caption         =   "Size To Fit"
            CheckBoxAngle   =   45
            CheckBoxColor1  =   4210752
            CheckBoxColor2  =   16777215
            CheckBoxMiddleOut=   0   'False
            CheckColor      =   0
            ControlType     =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseOverActions=   3
            MOverBorderColor=   16776960
            MOverCheckBoxColor=   16776960
            ShowFocusRect   =   0   'False
         End
         Begin Project1.MorphOptionCheck chkEnabled 
            Height          =   375
            Left            =   2760
            TabIndex        =   12
            Top             =   5160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BackColor1      =   4210752
            BackColor2      =   14737632
            Caption         =   "Enabled"
            CheckBoxAngle   =   45
            CheckBoxColor1  =   4210752
            CheckBoxColor2  =   16777215
            CheckBoxMiddleOut=   0   'False
            CheckColor      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseOverActions=   3
            MOverBorderColor=   16776960
            MOverCheckBoxColor=   16776960
            ShowFocusRect   =   0   'False
            Value           =   1
         End
         Begin VB.Label lblColorAtCursor 
            BackStyle       =   0  'Transparent
            Caption         =   "Color at Cursor:"
            Height          =   255
            Left            =   1080
            TabIndex        =   13
            Top             =   4200
            Width           =   2175
         End
         Begin VB.Label lblDispY1 
            BackStyle       =   0  'Transparent
            Caption         =   "Mouse X: 0"
            Height          =   255
            Left            =   1680
            TabIndex        =   11
            Top             =   3840
            Width           =   1335
         End
         Begin VB.Label lblDispX1 
            BackStyle       =   0  'Transparent
            Caption         =   "Mouse X: 0"
            Height          =   255
            Left            =   1680
            TabIndex        =   10
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label lblY 
            BackStyle       =   0  'Transparent
            Caption         =   "Mouse Y: 0"
            Height          =   255
            Left            =   2520
            TabIndex        =   9
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lblX 
            BackStyle       =   0  'Transparent
            Caption         =   "Mouse X: 0"
            Height          =   255
            Left            =   600
            TabIndex        =   8
            Top             =   480
            Width           =   1575
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkEnabled_Click()
   If chkEnabled.Value = vbChecked Then
      mpvNavigate.Enabled = True
   Else
      mpvNavigate.Enabled = False
   End If
End Sub

Private Sub chkKeepAspectRatio_Click()
   If chkKeepAspectRatio.Value = vbChecked Then
      mpvNavigate.KeepAspectRatio = True
   Else
      mpvNavigate.KeepAspectRatio = False
   End If
End Sub

Private Sub Form_Load()

   mrrVertical.ValueMin = 0
   mrrVertical.ValueMax = mpvNavigate.PicHeight - 1
   mrrVertical.Value = 0

   mrrHorizontal.ValueMin = mpvNavigate.PicY
   mrrHorizontal.ValueMax = mpvNavigate.PicWidth - 1
   mrrHorizontal.Value = mpvNavigate.PicX

End Sub


Private Sub mpvNavigate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'  update value properties w/coords of top left displayed image portion.
   mrrHorizontal.Value = mpvNavigate.DispX1
   mrrVertical.Value = mpvNavigate.DispY1

   lblX.Caption = "Mouse X: " & CStr(mpvNavigate.PicX)
   lblY.Caption = "Mouse Y: " & CStr(mpvNavigate.PicY)
   lblDispX1.Caption = "Upper Left X: " & CStr(mpvNavigate.DispX1)
   lblDispY1.Caption = "Upper Left Y: " & CStr(mpvNavigate.DispY1)
   lblColorAtCursor.Caption = "Color at Cursor: " & mpvNavigate.ColorAtCursor

End Sub

Private Sub mrrVertical_Change()
   mpvNavigate.DisplayImage 0, -1, mrrVertical.Value
   lblDispX1.Caption = "Upper Left X: " & CStr(mpvNavigate.DispX1)
   lblDispY1.Caption = "Upper Left Y: " & CStr(mpvNavigate.DispY1)
End Sub

Private Sub mrrVertical_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mrrVertical.Value = mpvNavigate.DispY1
End Sub

Private Sub mrrHorizontal_Change()
   mpvNavigate.DisplayImage 0, mrrHorizontal.Value, -1
   lblDispX1.Caption = "Upper Left X: " & CStr(mpvNavigate.DispX1)
   lblDispY1.Caption = "Upper Left Y: " & CStr(mpvNavigate.DispY1)
End Sub

Private Sub mrrHorizontal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mrrHorizontal.Value = mpvNavigate.DispX1
End Sub


Private Sub optPicStyle_Click(Index As Integer)
   Select Case Index
      Case 0
         mpvNavigate.PictureMode = mpv_Normal
      Case 1
         mpvNavigate.PictureMode = mpv_Stretch
   End Select
End Sub


