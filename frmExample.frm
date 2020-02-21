VERSION 5.00
Object = "*\AaxWidget.vbp"
Begin VB.Form frmExample 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F4F0EC&
   BorderStyle     =   0  'None
   ClientHeight    =   10005
   ClientLeft      =   7425
   ClientTop       =   4995
   ClientWidth     =   15345
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
   MinButton       =   0   'False
   ScaleHeight     =   667
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1023
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtC1v 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3360
      TabIndex        =   29
      Text            =   "00"
      Top             =   2610
      Width           =   420
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C4AEA0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4080
      TabIndex        =   19
      Text            =   "Search..."
      Top             =   360
      Width           =   4455
   End
   Begin axWidget.axWidgetc karo19 
      Height          =   1050
      Left            =   13245
      TabIndex        =   28
      Top             =   0
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1852
      BackColor       =   11640208
      BackColorOpacity=   0
      BorderColor     =   6771791
      BorderPosition  =   0
      Caption1        =   "frmExample.frx":0000
      Caption2        =   "frmExample.frx":002E
      Cursor          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FontAwesome"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      PictureAlignmentH=   1
      PictureAlignmentV=   1
   End
   Begin axWidget.axWidgetc karo30 
      Height          =   570
      Left            =   12360
      TabIndex        =   27
      Top             =   8640
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1005
      BackColor       =   9617920
      BorderColor     =   10997299
      BorderPosition  =   0
      BorderRadius    =   100
      Caption1        =   "frmExample.frx":0050
      Caption2        =   "frmExample.frx":007E
      Cursor          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin axWidget.axWidgetc karo29 
      Height          =   4695
      Left            =   3720
      TabIndex        =   26
      Top             =   4920
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8281
      BackColor       =   16777215
      Border          =   -1  'True
      BorderColor     =   13287609
      BorderRadius    =   4
      BorderWidth     =   1
      Caption1        =   "frmExample.frx":00B2
      Caption2        =   "frmExample.frx":00E0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   9201489
   End
   Begin axWidget.axWidgetc karo20 
      Height          =   1935
      Left            =   3990
      TabIndex        =   23
      Top             =   2520
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      BorderRadius    =   8
      Caption1        =   "frmExample.frx":0140
      Caption2        =   "frmExample.frx":016E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Gradient        =   -1  'True
      GradientAngle   =   225
      GradientColor1  =   7108093
      GradientColor2  =   6852607
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C4AEA0&
      FillColor       =   &H00C4AEA0&
      FillStyle       =   0  'Solid
      Height          =   840
      Left            =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   60
   End
   Begin axWidget.axWidgetc karo18 
      Height          =   1050
      Left            =   14295
      TabIndex        =   22
      Top             =   0
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   1852
      BackColor       =   11640208
      BackColorOpacity=   0
      BorderColor     =   6771791
      BorderPosition  =   0
      Caption1        =   "frmExample.frx":0196
      Caption2        =   "frmExample.frx":01C4
      Cursor          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FontAwesome"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      PictureAlignmentH=   1
      PictureAlignmentV=   1
   End
   Begin axWidget.axWidgetc karo17 
      Height          =   510
      Left            =   13335
      TabIndex        =   21
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   900
      BackColor       =   7902971
      BackColorOpacity=   0
      Border          =   -1  'True
      BorderColor     =   7902971
      BorderRadius    =   100
      BorderWidth     =   1
      Caption1        =   "frmExample.frx":01E6
      Caption2        =   "frmExample.frx":0214
      Cursor          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   7902971
   End
   Begin axWidget.axWidgetc karo6 
      Height          =   495
      Index           =   5
      Left            =   8520
      TabIndex        =   20
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      BackColorOpacity=   0
      Caption1        =   "frmExample.frx":0242
      Caption2        =   "frmExample.frx":0270
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "FontAwesome"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin axWidget.axWidgetc karo16 
      Height          =   510
      Left            =   3840
      TabIndex        =   18
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   900
      BackColorOpacity=   0
      Border          =   -1  'True
      BorderColor     =   16777215
      BorderRadius    =   100
      BorderWidth     =   1
      Caption1        =   "frmExample.frx":0292
      Caption2        =   "frmExample.frx":02C0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin axWidget.axWidgetc karo15 
      Height          =   240
      Left            =   1800
      TabIndex        =   17
      Top             =   1200
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   423
      BackColor       =   8251018
      Border          =   -1  'True
      BorderColor     =   6771791
      BorderPosition  =   0
      BorderRadius    =   100
      BorderWidth     =   2
      Caption1        =   "frmExample.frx":02EC
      Caption2        =   "frmExample.frx":030E
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
   Begin axWidget.axWidgetc karo14 
      Height          =   1050
      Left            =   3300
      TabIndex        =   16
      Top             =   1050
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   1852
      BackColor       =   16777215
      Caption1        =   "frmExample.frx":032E
      Caption2        =   "frmExample.frx":035C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   6311498
   End
   Begin axWidget.axWidgetc karo13 
      Height          =   1050
      Left            =   3300
      TabIndex        =   15
      Top             =   0
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   1852
      BackColor       =   12889760
      Caption1        =   "frmExample.frx":038A
      Caption2        =   "frmExample.frx":03B8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientAngle   =   90
      GradientColor1  =   6311498
      GradientColor2  =   12889760
   End
   Begin axWidget.axWidgetc karo12 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   8640
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   661
      BackColorOpacity=   0
      BorderPosition  =   0
      Caption1        =   "frmExample.frx":03D8
      Caption2        =   "frmExample.frx":0406
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12889760
   End
   Begin axWidget.axWidgetc karo10 
      Height          =   90
      Left            =   240
      TabIndex        =   12
      Top             =   9240
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   159
      BackColor       =   7902971
      BorderColor     =   16777215
      BorderPosition  =   0
      BorderRadius    =   100
      Caption1        =   "frmExample.frx":044E
      Caption2        =   "frmExample.frx":047C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin axWidget.axWidgetc karo11 
      Height          =   90
      Left            =   1080
      TabIndex        =   13
      Top             =   9240
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   159
      BackColor       =   6311498
      BorderColor     =   16777215
      BorderPosition  =   0
      BorderRadius    =   100
      Caption1        =   "frmExample.frx":049C
      Caption2        =   "frmExample.frx":04CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin axWidget.axWidgetc karo9 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3480
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   661
      BackColorOpacity=   0
      BorderPosition  =   0
      Caption1        =   "frmExample.frx":04EA
      Caption2        =   "frmExample.frx":0518
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12889760
   End
   Begin axWidget.axWidgetc karo8 
      Height          =   330
      Left            =   2400
      TabIndex        =   10
      Top             =   7440
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   582
      BackColor       =   7902971
      BorderColor     =   16777215
      BorderPosition  =   0
      BorderRadius    =   100
      Caption1        =   "frmExample.frx":0558
      Caption2        =   "frmExample.frx":0586
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
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
   Begin axWidget.axWidgetc karo7 
      Height          =   330
      Left            =   2760
      TabIndex        =   9
      Top             =   4920
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      BackColor       =   15968515
      BorderColor     =   16777215
      BorderPosition  =   0
      BorderRadius    =   100
      Caption1        =   "frmExample.frx":05AC
      Caption2        =   "frmExample.frx":05DA
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
   Begin axWidget.axWidgetc karo5 
      Height          =   840
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   3840
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1482
      BackColor       =   6771791
      Caption1        =   "frmExample.frx":05FC
      Caption2        =   "frmExample.frx":062A
      CaptionPadding  =   8
      Cursor          =   1
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
      ForeColorOpacity=   60
   End
   Begin axWidget.axWidgetc karo4 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   450
      BackColorOpacity=   0
      BorderPosition  =   0
      Caption1        =   "frmExample.frx":065C
      Caption2        =   "frmExample.frx":068A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin axWidget.axWidgetc karo3 
      Height          =   1200
      Left            =   885
      TabIndex        =   2
      Top             =   1125
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2117
      BackColorOpacity=   10
      BorderColor     =   16777215
      BorderPosition  =   0
      BorderRadius    =   100
      BorderWidth     =   4
      Caption1        =   "frmExample.frx":06BA
      Caption2        =   "frmExample.frx":06DC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignmentH=   1
      PictureAlignmentV=   1
   End
   Begin axWidget.axWidgetc karo2 
      Height          =   525
      Left            =   390
      TabIndex        =   1
      Top             =   360
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   926
      BackColorOpacity=   10
      BorderColorOpacity=   0
      Caption1        =   "frmExample.frx":06FC
      Caption2        =   "frmExample.frx":071E
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
   Begin axWidget.axWidgetc karo5 
      Height          =   840
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   4680
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1482
      BackColor       =   6771791
      Caption1        =   "frmExample.frx":073E
      Caption2        =   "frmExample.frx":076C
      CaptionPadding  =   8
      Cursor          =   1
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
      ForeColorOpacity=   60
   End
   Begin axWidget.axWidgetc karo5 
      Height          =   840
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   5520
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1482
      BackColor       =   6771791
      Caption1        =   "frmExample.frx":079A
      Caption2        =   "frmExample.frx":07C8
      CaptionPadding  =   8
      Cursor          =   1
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
      ForeColorOpacity=   60
   End
   Begin axWidget.axWidgetc karo5 
      Height          =   840
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   6360
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1482
      BackColor       =   6771791
      Caption1        =   "frmExample.frx":07F6
      Caption2        =   "frmExample.frx":0824
      CaptionPadding  =   8
      Cursor          =   1
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
      ForeColorOpacity=   60
   End
   Begin axWidget.axWidgetc karo5 
      Height          =   840
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   7200
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1482
      BackColor       =   6771791
      Caption1        =   "frmExample.frx":0856
      Caption2        =   "frmExample.frx":0884
      CaptionPadding  =   8
      Cursor          =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   7902971
      ForeColorOpacity=   60
   End
   Begin axWidget.axWidgetc karo1 
      Height          =   9990
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   17621
      BackColor       =   6771791
      Caption1        =   "frmExample.frx":08B4
      Caption2        =   "frmExample.frx":08D6
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
   Begin axWidget.axWidgetc karo25 
      Height          =   1935
      Left            =   7560
      TabIndex        =   24
      Top             =   2520
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      BorderRadius    =   8
      Caption1        =   "frmExample.frx":08F6
      Caption2        =   "frmExample.frx":0924
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
   Begin axWidget.axWidgetc karo28 
      Height          =   1935
      Left            =   11400
      TabIndex        =   25
      Top             =   2520
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      BorderRadius    =   8
      Caption1        =   "frmExample.frx":0950
      Caption2        =   "frmExample.frx":097E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
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
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    karo2.Picture = App.Path & "\img\eliteadmin_logo.jpg"
    karo3.Picture = App.Path & "\img\eliteadmin_head.jpg"
    
    karo20.Caption1 = "Esta es una prueba multilinea" & vbCrLf & _
                      "Esta es una prueba multilinea" & vbCrLf & _
                      "Esta es una prueba multilinea" & vbCrLf & _
                      "Esta es una prueba multilinea"
    karo20.Caption1VDistance = 250
    
End Sub

Private Sub karo17_MouseEnter()
    karo17.BackColorOpacity = 100
    karo17.ForeColor = vbWhite
End Sub

Private Sub karo17_MouseLeave()
    karo17.BackColorOpacity = 0
    karo17.ForeColor = &H7896FB
End Sub

Private Sub karo18_MouseEnter()
    karo18.BackColorOpacity = 100
End Sub

Private Sub karo18_MouseLeave()
    karo18.BackColorOpacity = 0
End Sub

Private Sub karo19_MouseEnter()
    karo19.BackColorOpacity = 100
End Sub

Private Sub karo19_MouseLeave()
    karo19.BackColorOpacity = 0
End Sub

Private Sub karo30_Click()
    Unload Me
End Sub

Private Sub karo30_MouseEnter()
    karo30.BackColor = &HA7CE33
End Sub

Private Sub karo30_MouseLeave()
    karo30.BackColor = &H92C200
End Sub

Private Sub karo5_MouseEnter(Index As Integer)
    karo5(Index).ForeColorOpacity = 100
    'karo6(Index).ForeColorOpacity = 100
    'karo5(Index).BackColor = &H604E4A
    
    Shape1.Visible = True
    Shape1.Top = karo5(Index).Top
End Sub

Private Sub karo5_MouseLeave(Index As Integer)
    karo5(Index).ForeColorOpacity = 60
    'karo6(Index).ForeColorOpacity = 60
    'karo5(Index).BackColor = &H67544F
    
    Shape1.Visible = False
End Sub

Private Sub txtC1v_Change()
karo20.Caption1VDistance = CInt(txtC1v)
karo20.Refresh
End Sub
