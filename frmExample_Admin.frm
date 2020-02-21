VERSION 5.00
Object = "*\AaxWidget.vbp"
Begin VB.Form frmExample_Admin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00B1B192&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   7425
   ClientTop       =   4995
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   2475
      Left            =   5490
      ScaleHeight     =   2415
      ScaleWidth      =   3525
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2805
      Width           =   3585
      Begin axWidget.axWidgetc cWidget1 
         Height          =   705
         Index           =   3
         Left            =   345
         TabIndex        =   27
         Top             =   420
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   1244
         Border          =   -1  'True
         BorderColor     =   16744576
         BorderColorOnMouseOver=   12648384
         BorderRadius    =   8
         BorderWidth     =   2
         Caption1        =   "frmExample_Admin.frx":0000
         Caption2        =   "frmExample_Admin.frx":0030
         CaptionPadding  =   5
         ChangeColorOnClick=   -1  'True
         ChangeBorderColorOnMouseOver=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         Gradient        =   -1  'True
         GradientAngle   =   225
         GradientColor1  =   15103304
         GradientColor2  =   15291533
         Caption1SizeMinus=   6
         Caption1VDistance=   18
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   5130
      Top             =   1650
   End
   Begin axWidget.axWidgetc cPoint 
      Height          =   240
      Index           =   2
      Left            =   9900
      TabIndex        =   25
      Top             =   900
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BackColor       =   65280
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderPosition  =   0
      BorderRadius    =   100
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":0060
      Caption2        =   "frmExample_Admin.frx":0082
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      PictureAlignmentH=   1
      PictureAlignmentV=   1
   End
   Begin axWidget.axWidgetc cPoint 
      Height          =   240
      Index           =   1
      Left            =   7260
      TabIndex        =   24
      Top             =   900
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BackColor       =   65280
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderPosition  =   0
      BorderRadius    =   100
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":00A2
      Caption2        =   "frmExample_Admin.frx":00C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      PictureAlignmentH=   1
      PictureAlignmentV=   1
   End
   Begin axWidget.axWidgetc cPoint 
      Height          =   240
      Index           =   0
      Left            =   4635
      TabIndex        =   23
      Top             =   900
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BackColor       =   65280
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderPosition  =   0
      BorderRadius    =   100
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":00E4
      Caption2        =   "frmExample_Admin.frx":0106
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      PictureAlignmentH=   1
      PictureAlignmentV=   1
   End
   Begin axWidget.axWidgetc cWidgeet45 
      Height          =   420
      Left            =   2685
      TabIndex        =   22
      Top             =   270
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   741
      BorderRadius    =   8
      Caption1        =   "frmExample_Admin.frx":0126
      Caption2        =   "frmExample_Admin.frx":0148
      CaptionPadding  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Gradient        =   -1  'True
      GradientAngle   =   225
      GradientColor1  =   7385702
      GradientColor2  =   13410892
   End
   Begin axWidget.axWidgetc cWidget1 
      Height          =   705
      Index           =   2
      Left            =   7950
      TabIndex        =   21
      Top             =   810
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   1244
      Border          =   -1  'True
      BorderColor     =   16744576
      BorderColorOnMouseOver=   12648384
      BorderRadius    =   8
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":01B4
      Caption2        =   "frmExample_Admin.frx":01EC
      CaptionPadding  =   5
      ChangeBorderColorOnMouseOver=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Gradient        =   -1  'True
      GradientAngle   =   225
      GradientColor1  =   15103304
      GradientColor2  =   15291533
      Caption1SizeMinus=   6
      Caption1VDistance=   18
   End
   Begin axWidget.axWidgetc cWidget1 
      Height          =   705
      Index           =   1
      Left            =   5310
      TabIndex        =   20
      Top             =   825
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   1244
      Border          =   -1  'True
      BorderColor     =   16744576
      BorderColorOnMouseOver=   12648384
      BorderRadius    =   8
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":021E
      Caption2        =   "frmExample_Admin.frx":0250
      CaptionPadding  =   5
      ChangeBorderColorOnMouseOver=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Gradient        =   -1  'True
      GradientAngle   =   225
      GradientColor1  =   15103304
      GradientColor2  =   15291533
      Caption1SizeMinus=   6
      Caption1VDistance=   18
   End
   Begin axWidget.axWidgetc cWidget5 
      Height          =   240
      Left            =   1500
      TabIndex        =   19
      Top             =   225
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BackColor       =   8251018
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderPosition  =   0
      BorderRadius    =   100
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":0280
      Caption2        =   "frmExample_Admin.frx":02A2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      PictureAlignmentH=   1
      PictureAlignmentV=   1
   End
   Begin axWidget.axWidgetc cWidget1 
      Height          =   705
      Index           =   0
      Left            =   2685
      TabIndex        =   18
      Top             =   825
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   1244
      Border          =   -1  'True
      BorderColor     =   16744576
      BorderColorOnMouseOver=   12648384
      BorderRadius    =   8
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":02C2
      Caption2        =   "frmExample_Admin.frx":02F2
      CaptionPadding  =   5
      ChangeBorderColorOnMouseOver=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Gradient        =   -1  'True
      GradientAngle   =   225
      GradientColor1  =   15103304
      GradientColor2  =   15291533
      Caption1SizeMinus=   6
      Caption1VDistance=   18
   End
   Begin axWidget.axWidgetc axWidgetc6 
      Height          =   585
      Left            =   5085
      TabIndex        =   17
      Top             =   6570
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1032
      BackColor       =   8421504
      BackColorOpacity=   80
      Border          =   -1  'True
      BorderColor     =   8421504
      BorderPosition  =   0
      BorderRadius    =   10
      BorderSmoothEdge=   -1  'True
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":0322
      Caption2        =   "frmExample_Admin.frx":0382
      CaptionPadding  =   10
      ChangeColorOnClick=   -1  'True
      ChangeBorderColorOnMouseOver=   -1  'True
      Cursor          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      GradientAngle   =   45
      GradientColor1  =   8421504
      GradientColor1Opacity=   80
      GradientColor2  =   4210752
      GradientColor2Opacity=   80
      Caption1SizeMinus=   3
      Caption1VDistance=   10
   End
   Begin axWidget.axWidgetc axWidgetc5 
      Height          =   4605
      Left            =   2535
      TabIndex        =   16
      Top             =   1845
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   8123
      BackColor       =   16777215
      Border          =   -1  'True
      BorderColor     =   8421504
      BorderRadius    =   10
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":03AC
      Caption2        =   "frmExample_Admin.frx":03F2
      CaptionPadding  =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8421504
      Caption1SizeMinus=   4
      Caption1VDistance=   150
   End
   Begin axWidget.axWidgetc axWidgetc4 
      Height          =   1590
      Left            =   2535
      TabIndex        =   15
      Top             =   120
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   2805
      BackColor       =   16777215
      Border          =   -1  'True
      BorderColor     =   8421504
      BorderRadius    =   10
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":0428
      Caption2        =   "frmExample_Admin.frx":044A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin axWidget.axWidgetc axWidgetc3 
      Height          =   1200
      Left            =   600
      TabIndex        =   14
      Top             =   180
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
      BackColor       =   14737632
      BorderRadius    =   100
      CaptionAlignmentV=   2
      Caption1        =   "frmExample_Admin.frx":046C
      Caption2        =   "frmExample_Admin.frx":048E
      ForeColorOnPress=   16711680
      ChangeColorOnClick=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12632256
      PictureAlignmentH=   1
      PictureAlignmentV=   1
   End
   Begin axWidget.axWidgetc Tick 
      Height          =   585
      Index           =   5
      Left            =   0
      TabIndex        =   13
      Top             =   4875
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   1032
      Caption1        =   "frmExample_Admin.frx":04BC
      Caption2        =   "frmExample_Admin.frx":04DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin axWidget.axWidgetc Button 
      Height          =   645
      Index           =   5
      Left            =   45
      TabIndex        =   12
      Top             =   4845
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1138
      BackColor       =   6771791
      BackColorPress  =   13409179
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderColorOnMouseOver=   14737632
      BorderSmoothEdge=   -1  'True
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":0500
      Caption2        =   "frmExample_Admin.frx":054C
      CaptionPadding  =   5
      ForeColorOnPress=   12632256
      ChangeColorOnClick=   -1  'True
      ChangeBorderColorOnMouseOver=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      GradientAngle   =   200
      GradientColor1  =   9333357
      GradientColor1Opacity=   80
      GradientColor2  =   13218225
      GradientColor2Opacity=   80
      GradientColorP1 =   13218225
      GradientColorP1Opacity=   80
      GradientColorP2 =   9333357
      GradientColorP2Opacity=   80
      Caption1SizeMinus=   3
      Caption1VDistance=   12
   End
   Begin axWidget.axWidgetc Tick 
      Height          =   585
      Index           =   4
      Left            =   0
      TabIndex        =   11
      Top             =   4215
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   1032
      Caption1        =   "frmExample_Admin.frx":0582
      Caption2        =   "frmExample_Admin.frx":05A4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin axWidget.axWidgetc Button 
      Height          =   645
      Index           =   4
      Left            =   45
      TabIndex        =   10
      Top             =   4185
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1138
      BackColor       =   6771791
      BackColorPress  =   13409179
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderColorOnMouseOver=   14737632
      BorderSmoothEdge=   -1  'True
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":05C6
      Caption2        =   "frmExample_Admin.frx":0610
      CaptionPadding  =   5
      ForeColorOnPress=   12632256
      ChangeColorOnClick=   -1  'True
      ChangeBorderColorOnMouseOver=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      GradientAngle   =   200
      GradientColor1  =   9333357
      GradientColor1Opacity=   80
      GradientColor2  =   13218225
      GradientColor2Opacity=   80
      GradientColorP1 =   13218225
      GradientColorP1Opacity=   80
      GradientColorP2 =   9333357
      Caption1SizeMinus=   3
      Caption1VDistance=   12
   End
   Begin axWidget.axWidgetc Button 
      Height          =   645
      Index           =   2
      Left            =   45
      TabIndex        =   9
      Top             =   2865
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1138
      BackColor       =   6771791
      BackColorPress  =   13409179
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderColorOnMouseOver=   14737632
      BorderSmoothEdge=   -1  'True
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":0642
      Caption2        =   "frmExample_Admin.frx":067E
      CaptionPadding  =   5
      ForeColorOnPress=   12632256
      ChangeColorOnClick=   -1  'True
      ChangeBorderColorOnMouseOver=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      GradientAngle   =   200
      GradientColor1  =   9333357
      GradientColor1Opacity=   80
      GradientColor2  =   13218225
      GradientColor2Opacity=   80
      GradientColorP1 =   13218225
      GradientColorP1Opacity=   80
      GradientColorP2 =   9333357
      Caption1SizeMinus=   3
      Caption1VDistance=   12
   End
   Begin axWidget.axWidgetc Tick 
      Height          =   585
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   2895
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   1032
      Caption1        =   "frmExample_Admin.frx":06B6
      Caption2        =   "frmExample_Admin.frx":06D8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin axWidget.axWidgetc Button 
      Height          =   645
      Index           =   3
      Left            =   45
      TabIndex        =   7
      Top             =   3525
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1138
      BackColor       =   6771791
      BackColorPress  =   13409179
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderColorOnMouseOver=   14737632
      BorderSmoothEdge=   -1  'True
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":06FA
      Caption2        =   "frmExample_Admin.frx":0746
      CaptionPadding  =   5
      ForeColorOnPress=   12632256
      ChangeColorOnClick=   -1  'True
      ChangeBorderColorOnMouseOver=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      GradientAngle   =   200
      GradientColor1  =   9333357
      GradientColor1Opacity=   80
      GradientColor2  =   13218225
      GradientColor2Opacity=   80
      GradientColorP1 =   13218225
      GradientColorP1Opacity=   80
      GradientColorP2 =   9333357
      Caption1SizeMinus=   3
      Caption1VDistance=   12
   End
   Begin axWidget.axWidgetc Tick 
      Height          =   585
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   3555
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   1032
      Caption1        =   "frmExample_Admin.frx":0784
      Caption2        =   "frmExample_Admin.frx":07A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin axWidget.axWidgetc Tick 
      Height          =   585
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   2235
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   1032
      BackColor       =   16777215
      Caption1        =   "frmExample_Admin.frx":07C8
      Caption2        =   "frmExample_Admin.frx":07EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin axWidget.axWidgetc Button 
      Height          =   645
      Index           =   1
      Left            =   45
      TabIndex        =   4
      Top             =   2205
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1138
      BackColor       =   6771791
      BackColorPress  =   13409179
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderColorOnMouseOver=   14737632
      BorderSmoothEdge=   -1  'True
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":080C
      Caption2        =   "frmExample_Admin.frx":0856
      CaptionPadding  =   5
      ForeColorOnPress=   12632256
      ChangeColorOnClick=   -1  'True
      ChangeBorderColorOnMouseOver=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      GradientAngle   =   200
      GradientColor1  =   9333357
      GradientColor1Opacity=   80
      GradientColor2  =   13218225
      GradientColor2Opacity=   80
      GradientColorP1 =   13218225
      GradientColorP1Opacity=   80
      GradientColorP2 =   9333357
      Caption1SizeMinus=   3
      Caption1VDistance=   12
   End
   Begin axWidget.axWidgetc Tick 
      Height          =   585
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1575
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   1032
      BorderColor     =   0
      Caption1        =   "frmExample_Admin.frx":0886
      Caption2        =   "frmExample_Admin.frx":08A8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin axWidget.axWidgetc Button 
      Height          =   645
      Index           =   0
      Left            =   45
      TabIndex        =   2
      Top             =   1545
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   1138
      BackColor       =   6771791
      BackColorPress  =   13409179
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderColorOnMouseOver=   14737632
      BorderSmoothEdge=   -1  'True
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":08CA
      Caption2        =   "frmExample_Admin.frx":090A
      CaptionPadding  =   5
      ForeColorOnPress=   12632256
      ChangeColorOnClick=   -1  'True
      ChangeBorderColorOnMouseOver=   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      GradientAngle   =   200
      GradientColor1  =   9333357
      GradientColor1Opacity=   80
      GradientColor2  =   13218225
      GradientColor2Opacity=   80
      GradientColorP1 =   13218225
      GradientColorP1Opacity=   80
      GradientColorP2 =   9333357
      Caption1SizeMinus=   3
      Caption1VDistance=   12
   End
   Begin axWidget.axWidgetc axWidgetc2 
      Height          =   7335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   12938
      BackColor       =   6771791
      BackColorOpacity=   80
      BorderCorner    =   3
      BorderPosition  =   0
      BorderRadius    =   100
      Caption1        =   "frmExample_Admin.frx":0938
      Caption2        =   "frmExample_Admin.frx":095A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin axWidget.axWidgetc axWidgetc1 
      Height          =   585
      Left            =   7815
      TabIndex        =   0
      Top             =   6570
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1032
      BackColor       =   8421504
      BackColorOpacity=   80
      Border          =   -1  'True
      BorderColor     =   8421504
      BorderPosition  =   0
      BorderRadius    =   10
      BorderSmoothEdge=   -1  'True
      BorderWidth     =   2
      Caption1        =   "frmExample_Admin.frx":097C
      Caption2        =   "frmExample_Admin.frx":09BE
      CaptionPadding  =   10
      ChangeColorOnClick=   -1  'True
      ChangeBorderColorOnMouseOver=   -1  'True
      Cursor          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      GradientAngle   =   45
      GradientColor1  =   8421504
      GradientColor1Opacity=   80
      GradientColor2  =   4210752
      GradientColor2Opacity=   80
      Caption1SizeMinus=   3
      Caption1VDistance=   10
   End
End
Attribute VB_Name = "frmExample_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim I As Integer, iMov As Integer

Private Sub axWidgetc1_MouseEnter()
axWidgetc1.Caption2 = "Click para Cerrar"
End Sub

Private Sub axWidgetc1_MouseLeave()
axWidgetc1.Caption2 = "CERRAR"
End Sub

Private Sub axWidgetc1_Click()
    Unload Me
End Sub

Private Sub axWidgetc6_Click()
frmExample.Show
End Sub

Private Sub Button_MouseEnter(Index As Integer)
Tick(Index).BackColor = &HFF0000
End Sub

Private Sub Button_MouseLeave(Index As Integer)
Tick(Index).BackColor = &H8000000F
End Sub

Private Sub Form_Load()
axWidgetc3.Picture = App.Path & "\img\eliteadmin_head.jpg"
I = 0
iMov = axWidgetc5.Caption1VDistance

axWidgetc5.Caption1 = "Línea 29: la clase axWidget.axMenuWidget del control axMenuWidget1 no era una clase de control cargada." & vbCrLf & _
                      "Línea 39: la clase axWidget.axWidgetc del control axWidgetc3 no era una clase de control cargada." & vbCrLf & _
                      "Línea 69: la clase axWidget.axWidgetc del control axWidgetc1 no era una clase de control cargada." & vbCrLf & _
                      "Línea 102: la clase axWidget.axWidgetc del control axWidgetc1 no era una clase de control cargada." & vbCrLf & _
                      "Línea 135: la clase axWidget.axWidgetc del control axWidgetc1 no era una clase de control cargada." & vbCrLf & _
                      "Línea 168: la clase axWidget.axWidgetc del control axWidgetc1 no era una clase de control cargada." & vbCrLf & _
                      "Línea 201: la clase axWidget.axWidgetc del control axWidgetc2 no era una clase de control cargada." & vbCrLf & _
                      "Línea 305: el nombre de la propiedad _extentx de axMenuWidget1 no es válido." & vbCrLf & _
                      "Línea 361: el nombre de la propiedad _extenty de axMenuWidget1 no es válido." & vbCrLf & _
                      "Línea 374: el nombre de la propiedad menuitems de axMenuWidget1 no es válido." & vbCrLf & _
                      "Línea 452: el nombre de la propiedad _extentx de axWidgetc3 no es válido." & vbCrLf & _
                      "Línea 461: el nombre de la propiedad _extenty de axWidgetc3 no es válido."
End Sub

Private Sub Timer1_Timer()

I = I + 1
iMov = iMov - 1

Select Case I
  Case Is = 5, 8, 12
    cPoint(0).BackColor = &HFF00&
    cWidget1(0).Caption2 = Val(cWidget1(0).Caption2) + 3589.12
  Case Is = 1, 8, 12
    cPoint(1).BackColor = &HFF&
    cWidget1(1).Caption2 = Val(cWidget1(1).Caption2) - 2758.35
  Case Is = 3, 6, 9, 12
    cPoint(2).BackColor = &HFF&
    cWidget1(2).Caption2 = Val(cWidget1(2).Caption2) - 5845.58
  Case Else
    cPoint(0).BackColor = &HFF&
    cPoint(1).BackColor = &HFF00&
    cPoint(2).BackColor = &HFF00&
End Select

axWidgetc5.Caption1VDistance = iMov

If I = 13 Then I = 0

End Sub
