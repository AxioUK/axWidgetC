VERSION 5.00
Begin VB.UserControl axWidgetc 
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   PropertyPages   =   "axWidget2.ctx":0000
   ScaleHeight     =   2625
   ScaleWidth      =   2325
   Tag             =   "Not Over"
   ToolboxBitmap   =   "axWidget2.ctx":0010
   Windowless      =   -1  'True
   Begin VB.Timer tmrGlow 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   555
      Top             =   195
   End
   Begin VB.Timer tmrMOUSEOVER 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "axWidgetc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const AppVersion = "1.2.7"


Private Declare Function GdiAlphaBlend& Lib "gdi32" (ByVal hDC&, ByVal X&, ByVal Y&, ByVal dx&, ByVal dy&, ByVal hdcSrc&, ByVal srcx&, ByVal srcy&, ByVal SrcdX&, ByVal SrcdY&, ByVal lBlendFunction&)
Private BackBuf As cCairoSurface 'we use a BackBuffer in the same PixelSize as the Control (avoiding AutoRedraw=True on the Control itself)
Private SVG As cSVG

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'===== FOR CUSTOM MOUSE CURSOR ===== ==================================================
'Used to convert icons/bitmaps to stdPicture objects
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, ipic As IPicture) As Long
'Used to load the current hand cursor
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Const IDC_HAND As Long = 32649
Private myHandCursor As StdPicture
Private myHand_handle As Long
'===== FOR CUSTOM MOUSE CURSOR ===== ==================================================
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long

Private Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type


'Custom Enums:
Public Enum BorderCorner
    bcAll
    bcBottom
    bcBottomLeft
    bcBottomRight
    bcLeft
    bcNone
    bcRight
    bcTop
    bcTopLeft
    bcTopRight
End Enum
Public Enum BorderPosition
    bpInside
    bpOutside
End Enum
Public Enum CaptionAlignmentH
    cLeft
    cCenter
    cRight
End Enum
Public Enum CaptionAlignmentV
    cTop
    cMiddle
    cBottom
End Enum
Public Enum Cursor
    curArrow
    curHand
End Enum
Public Enum PictureAlignmentH
    cLeft
    cCenter
    cRight
End Enum
Public Enum PictureAlignmentV
    cTop
    cMiddle
    cBottom
End Enum


'Default Property Values:
Const m_def_Glowing = False
Const m_def_BackColor = &H8000000F
Const m_def_BackColorOpacity = 100
Const m_def_BackColorP = &HC0C0C0
Const m_def_BackColorPOpacity = 100
Const m_def_Border = False
Const m_def_BorderColor = vbBlue
Const m_def_BorderColorOpacity = 100
Const m_def_BorderColorP = vbWhite
Const m_def_BorderColorPOpacity = 100
Const m_def_BorderCorner = 0
Const m_def_BorderPosition = 1
Const m_def_BorderRadius = 0
Const m_def_BorderSmoothEdge = False
Const m_def_BorderWidth = 0
Const m_def_Caption = "Caption"
Const m_def_CaptionAlignmentH = 0
Const m_def_CaptionAlignmentV = 0
Const m_def_CaptionPadding = 1
Const m_def_ChangeBorderColorOnMouseOver = False
Const m_def_ChangeColorOnClick = False
Const m_def_Cursor = 0
Const m_def_FontAwesome = ""
Const m_def_ForeColor = &H80000012
Const m_def_ForeColorOpacity = 100
Const m_def_ForeColorP = &H80000012
Const m_def_ForeColorPOpacity = 100
Const m_def_Gradient = False
Const m_def_GradientAngle = 0
Const m_def_GradientColor1 = &HD3A042
Const m_def_GradientColor1Opacity = 100
Const m_def_GradientColor2 = &HE96E9B
Const m_def_GradientColor2Opacity = 100
Const m_def_GradientColorP1 = &HE96E9B
Const m_def_GradientColorP1Opacity = 100
Const m_def_GradientColorP2 = &HD3A042
Const m_def_GradientColorP2Opacity = 100
Const m_def_ParentControl = ""
Const m_def_Picture = ""
Const m_def_PictureAlignmentH = 0
Const m_def_PictureAlignmentV = 0
Const m_def_PictureOpacity = 100
Const m_def_PicturePadding = 0
Const m_def_PictureSVGScale = 64
Const m_def_Value = False
Const m_def_WordWrap = True


'Property Variables:
Dim m_Glowing As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_BackColorOpacity As Integer
Dim m_BackColorP As OLE_COLOR
Dim m_BackColorPOpacity As Integer
Dim m_Border As Boolean
Dim m_BorderColor As OLE_COLOR
Dim m_BorderColorOpacity As Integer
Dim m_BorderColorP As OLE_COLOR
Dim m_BorderColorPOpacity As Integer
Dim m_BorderCorner As BorderCorner
Dim m_BorderPosition As BorderPosition
Dim m_BorderRadius As Integer
Dim m_BorderSmoothEdge As Boolean
Dim m_BorderWidth As Integer
Dim m_BorderGlow As Integer
Dim m_OldBorderWidth As Integer
Dim m_Caption1() As Byte
Dim m_Caption2() As Byte
Dim m_CaptionAlignmentH As CaptionAlignmentH
Dim m_CaptionAlignmentV As CaptionAlignmentV
Dim m_CaptionPadding As Integer
Dim m_ChangeBorderColorOnMouseOver As Boolean
Dim m_ChangeColorOnClick As Boolean
Dim m_Cursor As Cursor
Dim WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1
Dim m_FontAwesome As String
Dim m_ForeColor As OLE_COLOR
Dim m_ForeColorOpacity As Integer
Dim m_ForeColorP As OLE_COLOR
Dim m_ForeColorPOpacity As Integer
Dim m_Gradient As Boolean
Dim m_GradientAngle As Integer
Dim m_GradientColor1 As OLE_COLOR
Dim m_GradientColor1Opacity As Integer
Dim m_GradientColor2 As OLE_COLOR
Dim m_GradientColor2Opacity As Integer
Dim m_GradientColorP1 As OLE_COLOR
Dim m_GradientColorP1Opacity As Integer
Dim m_GradientColorP2 As OLE_COLOR
Dim m_GradientColorP2Opacity As Integer
Dim m_ParentControl As String
Dim m_Picture As String    'StdPicture
Dim m_PictureAlignmentH As PictureAlignmentH
Dim m_PictureAlignmentV As PictureAlignmentV
Dim m_PictureOpacity As Integer
Dim m_PicturePadding As Integer
Dim m_PictureSVGScale As Long
Dim m_Value As Boolean
Dim m_WordWrap As Boolean
'------------------------
Dim m_FontMinus As Integer
Dim m_C1VDistance As Single

Dim m_Clicked As Boolean
Dim m_MouseOver As Boolean
Dim isDrawed As Boolean

'Get Property Variables:
Dim lngRowCount As Long
Dim lngWordCount As Long

Dim objActiveControl As Object
Dim c_lhWnd As Long

'Custom Events:
Public Event Click()
Public Event DblClick()
Public Event Change(ByVal text As String)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Public Event MouseEnter()
Public Event MouseLeave()

Public Sub Refresh()
    Draw BackBuf.CreateContext, BackBuf.Width, BackBuf.Height
    UserControl.Refresh
End Sub

'Horizontal alignment after WordWrap
Private Function AlignHAfterWordWrap(CC As cCairoContext, strTextLine As String) As Long
    Select Case m_CaptionAlignmentH
            Case 0
                AlignHAfterWordWrap = m_CaptionPadding + CheckBorderPosition
            Case 1
                AlignHAfterWordWrap = (UserControl.ScaleWidth / 2) - (CC.GetTextExtents(strTextLine) / 2)
            Case 2
                AlignHAfterWordWrap = UserControl.ScaleWidth - CC.GetTextExtents(strTextLine) - m_CaptionPadding - CheckBorderPosition
        End Select
End Function

'One row horizontal alignment Caption2
Private Function AlignHOneRow(CC As cCairoContext) As Long
    Select Case m_CaptionAlignmentH
        Case 0
            AlignHOneRow = m_CaptionPadding + CheckBorderPosition
        Case 1
            AlignHOneRow = (UserControl.ScaleWidth / 2) - (CC.GetTextExtents(CStr(m_Caption2)) / 2)
        Case 2
            AlignHOneRow = UserControl.ScaleWidth - CC.GetTextExtents(CStr(m_Caption2)) - m_CaptionPadding - CheckBorderPosition
    End Select
End Function

'One row horizontal alignment Caption1
Private Function AlignHOneRowC1(CC As cCairoContext) As Long
    Select Case m_CaptionAlignmentH
        Case 0  'Left
            AlignHOneRowC1 = m_CaptionPadding + CheckBorderPosition
        Case 1  'Center
            AlignHOneRowC1 = (UserControl.ScaleWidth / 2) - (CC.GetTextExtents(CStr(m_Caption1)) / 2)
        Case 2  'Right
            AlignHOneRowC1 = UserControl.ScaleWidth - CC.GetTextExtents(CStr(m_Caption1)) - m_CaptionPadding - CheckBorderPosition
    End Select
End Function

'Back picture horizontal alignment
Private Function AlignHPicture(CC As cCairoContext, imgWidth As Long) As Long
    Select Case m_PictureAlignmentH
        Case 0
            AlignHPicture = m_PicturePadding + CheckBorderPosition
        Case 1
            AlignHPicture = (UserControl.ScaleWidth / 2) - (imgWidth / 2)
        Case 2
            AlignHPicture = UserControl.ScaleWidth - imgWidth - m_PicturePadding - CheckBorderPosition
    End Select
End Function

'Vertical alignment after WordWrap
Private Function AlignVAfterWordWrap(CC As cCairoContext, lngRowCounter As Long) As Long
    Select Case m_CaptionAlignmentV
        Case 0
            AlignVAfterWordWrap = 0 + m_CaptionPadding + CheckBorderPosition * (m_C1VDistance)
        Case 1
            AlignVAfterWordWrap = (UserControl.ScaleHeight / 2) + ((CC.GetFontHeight * lngRowCount) / 2) * (m_C1VDistance / 10)
        Case 2
            AlignVAfterWordWrap = UserControl.ScaleHeight + (CC.GetFontHeight * (lngRowCounter - lngRowCount)) - (CC.GetFontHeight * lngRowCounter) - m_CaptionPadding - CheckBorderPosition + (m_C1VDistance)
    End Select
End Function

'One row vertical alignment
Private Function AlignVOneRow(CC As cCairoContext) As Long
    Select Case m_CaptionAlignmentV
        Case 0
            AlignVOneRow = m_CaptionPadding + CheckBorderPosition
        Case 1
            AlignVOneRow = (UserControl.ScaleHeight / 2) - (CC.GetFontHeight / 2)
        Case 2
            AlignVOneRow = UserControl.ScaleHeight - CC.GetFontHeight - m_CaptionPadding - CheckBorderPosition
    End Select
End Function

Private Function AlignVOneRowC1(CC As cCairoContext) As Long
    Select Case m_CaptionAlignmentV
        Case 0
            AlignVOneRowC1 = m_CaptionPadding + CheckBorderPosition + (TextHeight(CStr(m_Caption1)) * (m_C1VDistance / 10))
        Case 1
            AlignVOneRowC1 = (UserControl.ScaleHeight / 2) - (CC.GetFontHeight / 2) - (m_C1VDistance / 2)   '(TextHeight(CStr(m_Caption1)) + 7)
        Case 2
            AlignVOneRowC1 = UserControl.ScaleHeight - CC.GetFontHeight - m_CaptionPadding - CheckBorderPosition - (TextHeight(CStr(m_Caption2)) + (m_C1VDistance / 2)) 'TextHeight(CStr(m_Caption1)))
    End Select
End Function

'Back picture vertical alignment
Private Function AlignVPicture(CC As cCairoContext, imgHeight As Long) As Long
    Select Case m_PictureAlignmentV
        Case 0
            AlignVPicture = m_PicturePadding + CheckBorderPosition
        Case 1
            AlignVPicture = (UserControl.ScaleHeight / 2) - (imgHeight / 2)
        Case 2
            AlignVPicture = UserControl.ScaleHeight - imgHeight - m_PicturePadding - CheckBorderPosition
    End Select
End Function

Private Sub Change(ByVal text As String)
    RaiseEvent Change(text)
End Sub

Private Sub ChangeMouseCursor()
    myHand_handle = LoadCursor(0, IDC_HAND)
            
    If myHand_handle <> 0 Then
        'Use function to convert memory handle to stdPicture so we can apply it to the MouseIcon
        Set myHandCursor = HandleToPicture(myHand_handle, False)
    End If
    
    If Not myHandCursor Is Nothing Then
        UserControl.MouseIcon = myHandCursor
        UserControl.MousePointer = vbCustom
    End If
End Sub

'Check border position
Private Function CheckBorderPosition() As Long
    If m_BorderPosition = bpInside Then
        CheckBorderPosition = 0
    Else
        CheckBorderPosition = m_BorderWidth
    End If
End Function

Private Sub Draw(CC As cCairoContext, ByVal dx As Long, ByVal dy As Long)
    Dim Pat As cCairoPattern

    CC.Operator = CAIRO_OPERATOR_CLEAR
    CC.Paint
    CC.Operator = CAIRO_OPERATOR_OVER

    CC.Save
        'This is need because our enum of m_BorderCorner doesn't correspond with enum of CornerMaskEnm
        Dim TempCornerEnum As CornerMaskEnm
        
        'I want to preserve alphabetical order of our enum
        Select Case m_BorderCorner
            Case 0
                TempCornerEnum = cmAll
            Case 1
                TempCornerEnum = cmBottom
            Case 2
                TempCornerEnum = cmBottomLeft
            Case 3
                TempCornerEnum = cmBottomRight
            Case 4
                TempCornerEnum = cmLeft
            Case 5
                TempCornerEnum = cmNone
            Case 6
                TempCornerEnum = cmRight
            Case 7
                TempCornerEnum = cmTop
            Case 8
                TempCornerEnum = cmTopLeft
            Case 9
                TempCornerEnum = cmTopRight
        End Select
        
        'Draw the main background
        If m_ChangeColorOnClick And m_Clicked Then
          CC.SetSourceColor m_BackColorP, m_BackColorPOpacity / 100
        Else
          CC.SetSourceColor m_BackColor, m_BackColorOpacity / 100
        End If
        'Proper displaying and resizing of the main background
        CC.RoundedRect m_BorderWidth - 1, m_BorderWidth - 1, UserControl.ScaleWidth - (m_BorderWidth * 2) + 2, UserControl.ScaleHeight - (m_BorderWidth * 2) + 2, m_BorderRadius - m_BorderWidth, True, TempCornerEnum, False
        CC.Fill True 'do NOT want to close the path yet...

        'Draw gradient background if value of Gradient is set to True
        If m_Gradient = True Then
            'Now here comes the magic. I want to make gradient angle/direction easy to use like Photoshop does.
            Select Case m_GradientAngle
                Case Is <= 45
                    Set Pat = Cairo.CreateLinearPattern((UserControl.ScaleWidth / 2) + (UserControl.ScaleWidth / 90) * m_GradientAngle, 0, (UserControl.ScaleWidth / 2) - ((UserControl.ScaleWidth / 90) * m_GradientAngle), UserControl.ScaleHeight)
                Case Is <= 135
                    Set Pat = Cairo.CreateLinearPattern(UserControl.ScaleWidth, (UserControl.ScaleHeight / 90) * (m_GradientAngle - 45), 0, UserControl.ScaleHeight - (UserControl.ScaleHeight / 90) * (m_GradientAngle - 45))
                Case Is <= 225
                    Set Pat = Cairo.CreateLinearPattern(UserControl.ScaleWidth - (UserControl.ScaleWidth / 90) * (m_GradientAngle - 135), UserControl.ScaleHeight, (UserControl.ScaleWidth / 90) * (m_GradientAngle - 135), 0)
                Case Is <= 315
                    Set Pat = Cairo.CreateLinearPattern(0, UserControl.ScaleHeight - ((UserControl.ScaleHeight) / 90) * (m_GradientAngle - 225), UserControl.ScaleWidth, (UserControl.ScaleHeight / 90) * (m_GradientAngle - 225))
                Case Is <= 359
                    Set Pat = Cairo.CreateLinearPattern((UserControl.ScaleWidth / 90) * (m_GradientAngle - 315), 0, UserControl.ScaleWidth - (UserControl.ScaleWidth / 90) * (m_GradientAngle - 315), UserControl.ScaleHeight)
            End Select
            If m_ChangeColorOnClick And m_Clicked Then
              Pat.AddColorStop 0, m_GradientColorP1, m_GradientColorP1Opacity / 100 'Color 1 with opacity
              Pat.AddColorStop 1, m_GradientColorP2, m_GradientColorP2Opacity / 100 'Color 2 with opacity
            Else
              Pat.AddColorStop 0, m_GradientColor1, m_GradientColor1Opacity / 100 'Color 1 with opacity
              Pat.AddColorStop 1, m_GradientColor2, m_GradientColor2Opacity / 100 'Color 2 with opacity
            End If
            CC.Fill True, Pat 'do NOT want to close the path yet...
        End If

        'Draw border if value of Border is set to True
        'Border is need to be drawn like an another rounded rectangle but inside it have to be empty. This we can achieve if we use clippin mask for inside.
        If m_Border Then
            'Fill rule first
            CC.FillRule = CAIRO_FILL_RULE_EVEN_ODD
            
            'Smooth border and fix glitch which cause clipping
            If m_BorderSmoothEdge = True Then
              CC.SetLineWidth 0.5
              If m_ChangeColorOnClick And m_Clicked Then
                CC.SetSourceColor m_BackColorP, m_BackColorPOpacity
              Else
                CC.SetSourceColor m_BackColor, m_BackColorOpacity
              End If
                CC.Stroke True
            End If
            
            'We want to set a path what we want to clip
            If m_BorderSmoothEdge = True Then
                CC.RoundedRect 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_BorderRadius, True, TempCornerEnum, False
            Else
                CC.RoundedRect -1, -1, UserControl.ScaleWidth + 2, UserControl.ScaleHeight + 2, m_BorderRadius, True, TempCornerEnum, False
            End If
            
            'Now the just defined RoundedRect-Path above is acting like as the clipping-mask
            CC.Clip 'start of clipping
            If m_ChangeBorderColorOnMouseOver And m_MouseOver Then
              CC.SetSourceColor m_BorderColorP
              CC.Paint m_BorderColorPOpacity / 100
            Else
              CC.SetSourceColor m_BorderColor
              CC.Paint m_BorderColorOpacity / 100
            End If
            CC.ResetClip 'end of clipping
        End If

        'Load background image if exists
        If m_Picture <> "" Then
            'Check if file exists or error will occur
            If Dir(m_Picture) <> "" Then
                'Filter to image extensions: .png, .jpg, .bmp, .gif, .svg
                If Right(m_Picture, 3) = "png" Or Right(m_Picture, 3) = "jpg" Or Right(m_Picture, 3) = "bmp" Or Right(m_Picture, 3) = "gif" Or Right(m_Picture, 3) = "svg" Then
                    If Right(m_Picture, 3) = "svg" Then 'SVG support
                        Set SVG = New_c.SVG
                        
                        SVG.ParseContent m_Picture
                        SVG.RenderFromDOM CC, AlignHPicture(CC, m_PictureSVGScale), AlignVPicture(CC, m_PictureSVGScale), m_PictureSVGScale, m_PictureSVGScale, 0, False
                        
                    Else
                        'Set path to clipping
                        If m_BorderSmoothEdge = True Then
                            CC.RoundedRect m_BorderWidth, m_BorderWidth, UserControl.ScaleWidth - (m_BorderWidth * 2), UserControl.ScaleHeight - (m_BorderWidth * 2), m_BorderRadius - m_BorderWidth, True, TempCornerEnum, False
                        Else
                            CC.RoundedRect m_BorderWidth - 1, m_BorderWidth - 1, UserControl.ScaleWidth - (m_BorderWidth * 2) + 2, UserControl.ScaleHeight - (m_BorderWidth * 2) + 2, m_BorderRadius - m_BorderWidth, True, TempCornerEnum, False
                        End If
                        'Start to clipping
                        CC.Clip
                        'Add image to our image list
                        Cairo.ImageList.AddImage "SourceImage", m_Picture
                        'Draw picture from ImageList inside the bounds of our border
                        CC.SetSourceSurface Cairo.ImageList("SourceImage"), AlignHPicture(CC, Cairo.ImageList.Item("SourceImage").Width), AlignVPicture(CC, Cairo.ImageList.Item("SourceImage").Height)
                        'Paint clipped content with opacity
                        CC.Paint m_PictureOpacity / 100
                        'Close clipping
                        CC.ResetClip
                    End If
                End If
            End If
        Else
            'If we don't have any image just fill the content
            CC.Fill 'now we can finally close the path
        End If
    CC.Restore

    
    'We have to reset public variable RowCount here
    lngRowCount = 0
    
LabelOne:
    'If Caption is empty then just draw the background and Exit Sub
    'TIP OF THE DAY: You can use the empty control for various graphic uses
    If Trim$(CStr(m_Caption1)) = "" Then GoTo LabelTwo
    'Font settings here
    CC.SelectFont m_Font, m_Font.Size - m_FontMinus, m_ForeColor, m_Font.Bold, m_Font.Italic, m_Font.Underline, m_Font.Strikethrough
    'Check the word count
    Dim fields() As String
    Dim m_temp_Caption As String
    'We need to create the temporary caption to check the word count in proper way with removed extra spaces
    'In this case the words "HELLO   WORLD" counts properly just as two words
    m_temp_Caption = m_Caption1
    m_temp_Caption = RemoveExtraSpaces(m_temp_Caption)
    
    fields() = Split(m_temp_Caption, " ")
    lngWordCount = UBound(fields()) + 1 'WordCount for public use
    'We have to do split again, but now there will be all spaces like in the Caption and display it properly later in result of the User Control
    'At first split we want to separate all rows with vbCrLf (new line) and put it in to the array
    Dim riadky() As String
    Dim slova() As String
    
    riadky() = Split(m_Caption1, vbCrLf)
    'WordWrap starts here
    If m_WordWrap Then
        'First condition is to determine if Caption is even longer or not that User Control. If not we can skip right to the end and make simple alignment
        If (m_BorderWidth * 2) + (m_CaptionPadding * 2) + CC.GetTextExtents(CStr(m_Caption1)) > UserControl.ScaleWidth Then
            Dim r As Long
            Dim i As Long
            Dim strResultString As String 'a string we will work with
            Dim lngRowCounter As Long 'for counting an actual rows
            'At first we need to know with how many rows we will do the vertical alignment
            For r = 0 To UBound(riadky())
                slova() = Split(riadky(r), " ")
                strResultString = ""
                For i = 0 To UBound(slova())
                    If (m_BorderWidth * 2) + (m_CaptionPadding * 2) + CC.GetTextExtents(strResultString & slova(i)) < UserControl.ScaleWidth Then
                        strResultString = strResultString & slova(i) & " "
                    Else
                        lngRowCount = lngRowCount + 1 'RowCount for public and internal use
                        strResultString = slova(i) & " "
                    End If
                Next i
                lngRowCount = lngRowCount + 1 'RowCount for public and internal use
            Next r
            'Now we can run through the loop once more again but now we will display the text itself
            'First loop for rows only
            For r = 0 To UBound(riadky())
                'Split rows to words
                slova() = Split(riadky(r), " ")
                'Clear the result for next use
                strResultString = ""
                'Second loop for words in current row
                For i = 0 To UBound(slova())
                    'If [ BorderWidth LEFT + CaptionPadding LEFT + SOME WORDS + NEW ADDED WORD + CaptionPadding RIGHT + BorderWidth RIGHT ] is smaller the width of the User Control then continue...
                    If (m_BorderWidth * 2) + (m_CaptionPadding * 2) + CC.GetTextExtents(strResultString & slova(i)) < UserControl.ScaleWidth Then
                        'Store all words which fits in to the width of User Control with CaptionPadding included
                        strResultString = strResultString & slova(i) & " "
                    Else
                        'When words are longer than width of the User control, then show them all which fits and go to the next row
                        CC.TextOut AlignHAfterWordWrap(CC, strResultString), AlignVAfterWordWrap(CC, lngRowCounter) + (lngRowCounter * CC.GetFontHeight), strResultString, False, m_ForeColorOpacity / 100, True
                        lngRowCounter = lngRowCounter + 1
                        strResultString = slova(i) & " "
                    End If
                Next i
                'Show all other words and go to the next row
                CC.TextOut AlignHAfterWordWrap(CC, strResultString), AlignVAfterWordWrap(CC, lngRowCounter) + (lngRowCounter * CC.GetFontHeight), strResultString, False, m_ForeColorOpacity / 100, True
                lngRowCounter = lngRowCounter + 1
            Next r
        Else
            'If WordWrap is set to True, but caption is still smaller than User Control then
            'do the simple one row horizontal and vertical alignment and text out the Caption
            CC.TextOut AlignHOneRowC1(CC), AlignVOneRowC1(CC), CStr(m_Caption1), , m_ForeColorOpacity / 100, True
            lngRowCount = lngRowCount + 1 'RowCount for public and internal use
        End If
    Else
        'If WordWrap is set to False then
        'do the simple one row horizontal and vertical alignment and text out the Caption
        CC.TextOut AlignHOneRowC1(CC), AlignVOneRowC1(CC), CStr(m_Caption1), , m_ForeColorOpacity / 100, True
        lngRowCount = lngRowCount + 1 'RowCount for public and internal use
    End If
 '------------------------------------------------------>>
LabelTwo:
    'If Caption is empty then just draw the background and Exit Sub
    'TIP OF THE DAY: You can use the empty control for various graphic uses
    If Trim$(CStr(m_Caption2)) = "" Then GoTo FillCairo
    'Font settings here
    If m_ChangeColorOnClick And m_Clicked Then
    CC.SelectFont m_Font, m_Font.Size, m_ForeColorP, m_Font.Bold, m_Font.Italic, m_Font.Underline, m_Font.Strikethrough
    CC.TextOut AlignHOneRow(CC), AlignVOneRow(CC), CStr(m_Caption2), , m_ForeColorPOpacity / 100, True

    Else
    CC.SelectFont m_Font, m_Font.Size, m_ForeColor, m_Font.Bold, m_Font.Italic, m_Font.Underline, m_Font.Strikethrough
    CC.TextOut AlignHOneRow(CC), AlignVOneRow(CC), CStr(m_Caption2), , m_ForeColorOpacity / 100, True
    End If
    lngRowCount = lngRowCount + 1 'RowCount for public and internal use
 '------------------------------------------------------>>
FillCairo:
    'At the end fill all the stored text
    CC.Fill
End Sub

'Convert an icon/bitmap handle to a Picture object
Private Function HandleToPicture(ByVal hHandle As Long, isBitmap As Boolean) As IPicture
    Dim pic As PICTDESC
    Dim guid(0 To 3) As Long
    
    'Initialize the PictDesc structure
    pic.cbSize = Len(pic)
    If isBitmap Then pic.pictType = vbPicTypeBitmap Else pic.pictType = vbPicTypeIcon
    pic.hIcon = hHandle
    
    'This is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    'We use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    
    'Create the picture,
    'Return an object reference right into the function result
    OleCreatePictureIndirect pic, guid(0), True, HandleToPicture
End Function

'Remove extra spaces for proper word count
Private Function RemoveExtraSpaces(strVal As String) As String
    Dim regEx
    Set regEx = New RegExp
    regEx.Global = True
    regEx.IgnoreCase = True
    
    regEx.Pattern = "\s{2,}"
    
    RemoveExtraSpaces = Trim(regEx.Replace(strVal, " "))
End Function

Private Sub tmrGlow_Timer()
If m_Glowing And m_BorderGlow <= 10 Then
  'm_BorderPosition = bpOutside
  m_BorderGlow = m_BorderGlow + 1
  m_BorderWidth = m_BorderGlow
  m_BorderColorOpacity = 100 - (m_BorderGlow * 10)
  Refresh
Else
  m_BorderGlow = 1
  m_BorderWidth = m_BorderGlow
  m_BorderColorOpacity = 60
End If
End Sub

Private Sub tmrMOUSEOVER_Timer()
If isDrawed And Not IsMouseInExtender Then Exit Sub

    If Not IsMouseInExtender Then
        tmrMOUSEOVER.Enabled = False
        Set objActiveControl = Nothing
        m_MouseOver = False
        RaiseEvent MouseLeave
    Else
        m_MouseOver = True
        RaiseEvent MouseEnter
    End If
    Refresh
End Sub

Public Function IsMouseInExtender() As Boolean
    Dim PT As POINTAPI
    Dim CPT As POINTAPI
    Dim TR As RECT
    Dim bArea As Boolean
    
    Call GetCursorPos(PT)
    Call ClientToScreen(c_lhWnd, CPT)
    
    CPT.X = PT.X - CPT.X
    CPT.Y = PT.Y - CPT.Y

    With TR
        .Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode) ' / nScale
        .Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode) ' / nScale
        .Right = .Left + UserControl.ScaleWidth
        .Bottom = .Top + UserControl.ScaleHeight
    End With
    
    bArea = PtInRect(TR, CPT.X, CPT.Y)
    
    If bArea And WindowFromPoint(PT.X, PT.Y) = c_lhWnd Then
        IsMouseInExtender = True
        isDrawed = True
    Else
        IsMouseInExtender = False
        isDrawed = False
    End If

End Function



'===== CUSTOM EVENTS ===== ======================================= CUSTOM EVENTS ===================================== CUSTOM EVENTS ======================================

Private Sub UserControl_Click()
    RaiseEvent Click

    'Take responsibility for a parent click event
    If m_ParentControl <> "" Then
        UserControl.Parent.Controls(m_ParentControl).Value = True
    End If
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_Initialize()
    Set m_Font = New StdFont
    Set UserControl.Font = m_Font
    isDrawed = False
    ScaleMode = vbPixels
    ClipBehavior = 0
    BackStyle = 0
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_BackColorOpacity = m_def_BackColorOpacity
    m_BackColorP = m_def_BackColorP
    m_BackColorPOpacity = m_def_BackColorPOpacity
    m_Border = m_def_Border
    m_BorderColor = m_def_BorderColor
    m_BorderColorOpacity = m_def_BorderColorOpacity
    m_BorderColorP = m_def_BorderColorP
    m_BorderColorPOpacity = m_def_BorderColorPOpacity
    m_BorderCorner = m_def_BorderCorner
    m_BorderPosition = m_def_BorderPosition
    m_BorderRadius = m_def_BorderRadius
    m_BorderSmoothEdge = m_def_BorderSmoothEdge
    m_BorderWidth = m_def_BorderWidth
    m_Caption1 = m_def_Caption
    m_Caption2 = m_def_Caption
    m_CaptionAlignmentH = m_def_CaptionAlignmentH
    m_CaptionAlignmentV = m_def_CaptionAlignmentV
    m_CaptionPadding = m_def_CaptionPadding
    m_ChangeBorderColorOnMouseOver = m_def_ChangeBorderColorOnMouseOver
    m_ChangeColorOnClick = m_def_ChangeColorOnClick
    m_Cursor = m_def_Cursor
    Set m_Font = UserControl.Ambient.Font
    m_ForeColor = m_def_ForeColor
    m_ForeColorOpacity = m_def_ForeColorOpacity
    m_ForeColorP = m_def_ForeColorP
    m_ForeColorPOpacity = m_def_ForeColorPOpacity
    m_Gradient = m_def_Gradient
    m_GradientAngle = m_def_GradientAngle
    m_GradientColor1 = m_def_GradientColor1
    m_GradientColor1Opacity = m_def_GradientColor1Opacity
    m_GradientColor2 = m_def_GradientColor2
    m_GradientColor2Opacity = m_def_GradientColor2Opacity
    m_GradientColorP1 = m_def_GradientColorP1
    m_GradientColorP1Opacity = m_def_GradientColorP1Opacity
    m_GradientColorP2 = m_def_GradientColorP2
    m_GradientColorP2Opacity = m_def_GradientColorP2Opacity
    m_ParentControl = m_def_ParentControl
    m_Picture = m_def_Picture
    m_PictureAlignmentH = m_def_PictureAlignmentH
    m_PictureAlignmentV = m_def_PictureAlignmentV
    m_PictureOpacity = m_def_PictureOpacity
    m_PicturePadding = m_def_PicturePadding
    m_PictureSVGScale = m_def_PictureSVGScale
    m_Value = m_def_Value
    m_WordWrap = m_def_WordWrap
    m_FontMinus = 5
    m_C1VDistance = 20
    
    c_lhWnd = UserControl.ContainerHwnd

  m_Glowing = m_def_Glowing
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    m_Clicked = True
    Refresh
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    'UserControl.Parent.ScaleMode = 1
    tmrMOUSEOVER.Enabled = True
    Set objActiveControl = UserControl.Extender
    
'    'Take responsibility for a parent mousemove event
    If m_ParentControl <> "" Then
        Call UserControl.Parent.Controls(m_ParentControl).MoveAndClick(X, Y, Abs(UserControl.Parent.Controls(m_ParentControl).Left - UserControl.Extender.Left), Abs(UserControl.Parent.Controls(m_ParentControl).Top - UserControl.Extender.Top))
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    m_Clicked = False
    Refresh
End Sub

Private Sub UserControl_Paint() 'we use only the GDI-call in the Paint-Routine (for BackBuffer-Refreshs, use the Refresh-Method)
    GdiAlphaBlend hDC, 0, 0, BackBuf.Width, BackBuf.Height, BackBuf.GetDC, _
                       0, 0, BackBuf.Width, BackBuf.Height, 2 ^ 24 + &HFF0000 * 1
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    c_lhWnd = UserControl.ContainerHwnd
    
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_BackColorOpacity = PropBag.ReadProperty("BackColorOpacity", m_def_BackColorOpacity)
    m_BackColorP = PropBag.ReadProperty("BackColorPress", m_def_BackColorP)
    m_BackColorPOpacity = PropBag.ReadProperty("BackColorPressOpacity", m_def_BackColorPOpacity)
    m_Border = PropBag.ReadProperty("Border", m_def_Border)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderColorOpacity = PropBag.ReadProperty("BorderColorOpacity", m_def_BorderColorOpacity)
    m_BorderColorP = PropBag.ReadProperty("BorderColorOnMouseOver", m_def_BorderColorP)
    m_BorderColorPOpacity = PropBag.ReadProperty("BorderColorOnMouseOverOpacity", m_def_BorderColorPOpacity)
    m_BorderCorner = PropBag.ReadProperty("BorderCorner", m_def_BorderCorner)
    m_BorderPosition = PropBag.ReadProperty("BorderPosition", m_def_BorderPosition)
    m_BorderRadius = PropBag.ReadProperty("BorderRadius", m_def_BorderRadius)
    m_BorderSmoothEdge = PropBag.ReadProperty("BorderSmoothEdge", m_def_BorderSmoothEdge)
    m_BorderWidth = PropBag.ReadProperty("BorderWidth", m_def_BorderWidth)
    m_CaptionAlignmentH = PropBag.ReadProperty("CaptionAlignmentH", m_def_CaptionAlignmentH)
    m_CaptionAlignmentV = PropBag.ReadProperty("CaptionAlignmentV", m_def_CaptionAlignmentV)
    m_Caption1 = PropBag.ReadProperty("CaptionSub", m_def_Caption)
    m_Caption2 = PropBag.ReadProperty("CaptionMain", m_def_Caption)
    m_CaptionPadding = PropBag.ReadProperty("CaptionPadding", m_def_CaptionPadding)
    m_ChangeColorOnClick = PropBag.ReadProperty("ChangeColorOnClick", m_def_ChangeColorOnClick)
    m_ChangeBorderColorOnMouseOver = PropBag.ReadProperty("ChangeBorderColorOnMouseOver", m_def_ChangeBorderColorOnMouseOver)
    m_Cursor = PropBag.ReadProperty("Cursor", m_def_Cursor)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set m_Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_ForeColorOpacity = PropBag.ReadProperty("ForeColorOpacity", m_def_ForeColorOpacity)
    m_ForeColorP = PropBag.ReadProperty("ForeColorOnPress", m_def_ForeColorP)
    m_ForeColorPOpacity = PropBag.ReadProperty("ForeColorOnPressOpacity", m_def_ForeColorPOpacity)
    m_Gradient = PropBag.ReadProperty("Gradient", m_def_Gradient)
    m_GradientAngle = PropBag.ReadProperty("GradientAngle", m_def_GradientAngle)
    m_GradientColor1 = PropBag.ReadProperty("GradientColor1", m_def_GradientColor1)
    m_GradientColor1Opacity = PropBag.ReadProperty("GradientColor1Opacity", m_def_GradientColor1Opacity)
    m_GradientColor2 = PropBag.ReadProperty("GradientColor2", m_def_GradientColor2)
    m_GradientColor2Opacity = PropBag.ReadProperty("GradientColor2Opacity", m_def_GradientColor2Opacity)
    m_GradientColorP1 = PropBag.ReadProperty("GradientColorP1", m_def_GradientColorP1)
    m_GradientColorP1Opacity = PropBag.ReadProperty("GradientColorP1Opacity", m_def_GradientColorP1Opacity)
    m_GradientColorP2 = PropBag.ReadProperty("GradientColorP2", m_def_GradientColorP2)
    m_GradientColorP2Opacity = PropBag.ReadProperty("GradientColorP2Opacity", m_def_GradientColorP2Opacity)
    
    m_ParentControl = PropBag.ReadProperty("ParentControl", m_def_ParentControl)
    m_Picture = PropBag.ReadProperty("Picture", m_def_Picture)
    m_PictureAlignmentH = PropBag.ReadProperty("PictureAlignmentH", m_def_PictureAlignmentH)
    m_PictureAlignmentV = PropBag.ReadProperty("PictureAlignmentV", m_def_PictureAlignmentV)
    m_PictureOpacity = PropBag.ReadProperty("PictureOpacity", m_def_PictureOpacity)
    m_PicturePadding = PropBag.ReadProperty("PicturePadding", m_def_PicturePadding)
    m_PictureSVGScale = PropBag.ReadProperty("PictureSVGScale", m_def_PictureSVGScale)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
    m_FontMinus = PropBag.ReadProperty("CaptionSubSizeMinus", 5)
    m_C1VDistance = PropBag.ReadProperty("CaptionSubVDistance", 20)
  m_Glowing = PropBag.ReadProperty("Glowing", m_def_Glowing)
End Sub

Private Sub UserControl_Resize()
    Set BackBuf = Cairo.CreateWin32Surface(ScaleWidth, ScaleHeight)
    
    Refresh
End Sub

Private Sub UserControl_Show()
    Refresh

    'Check if parent control have hand cursor on and apply it to a child control
    If m_ParentControl <> "" Then
        If UserControl.Parent.Controls(m_ParentControl).Cursor = 1 Then
            ChangeMouseCursor
            
            Exit Sub
        End If
    End If
    
    'Also show custom mouse cursor, in this case it will be the hand cursor
    If m_Cursor = curHand Then ChangeMouseCursor
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BackColorOpacity", m_BackColorOpacity, m_def_BackColorOpacity)
        Call PropBag.WriteProperty("BackColorPress", m_BackColorP, m_def_BackColorP)
        Call PropBag.WriteProperty("BackColorPressOpacity", m_BackColorPOpacity, m_def_BackColorPOpacity)
   
    Call PropBag.WriteProperty("Border", m_Border, m_def_Border)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderColorOpacity", m_BorderColorOpacity, m_def_BorderColorOpacity)
        Call PropBag.WriteProperty("BorderColorOnMouseOver", m_BorderColorP, m_def_BorderColorP)
        Call PropBag.WriteProperty("BorderColorOpacityOnMouseOver", m_BorderColorPOpacity, m_def_BorderColorPOpacity)
    
    Call PropBag.WriteProperty("BorderCorner", m_BorderCorner, m_def_BorderCorner)
    Call PropBag.WriteProperty("BorderPosition", m_BorderPosition, m_def_BorderPosition)
    Call PropBag.WriteProperty("BorderRadius", m_BorderRadius, m_def_BorderRadius)
    Call PropBag.WriteProperty("BorderSmoothEdge", m_BorderSmoothEdge, m_def_BorderSmoothEdge)
    Call PropBag.WriteProperty("BorderWidth", m_BorderWidth, m_def_BorderWidth)
    Call PropBag.WriteProperty("CaptionAlignmentH", m_CaptionAlignmentH, m_def_CaptionAlignmentH)
    Call PropBag.WriteProperty("CaptionAlignmentV", m_CaptionAlignmentV, m_def_CaptionAlignmentV)
    Call PropBag.WriteProperty("CaptionSub", m_Caption1, m_def_Caption)
    Call PropBag.WriteProperty("CaptionMain", m_Caption2, m_def_Caption)
    Call PropBag.WriteProperty("CaptionPadding", m_CaptionPadding, m_def_CaptionPadding)
        Call PropBag.WriteProperty("ForeColorOnPress", m_ForeColorP, m_def_ForeColorP)
        Call PropBag.WriteProperty("ForeColorOnPressOpacity", m_ForeColorPOpacity, m_def_ForeColorPOpacity)
        Call PropBag.WriteProperty("ChangeColorOnClick", m_ChangeColorOnClick, m_def_ChangeColorOnClick)
        Call PropBag.WriteProperty("ChangeBorderColorOnMouseOver", m_ChangeBorderColorOnMouseOver, m_def_ChangeBorderColorOnMouseOver)
    
    Call PropBag.WriteProperty("Cursor", m_Cursor, m_def_Cursor)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("ForeColorOpacity", m_ForeColorOpacity, m_def_ForeColorOpacity)
    Call PropBag.WriteProperty("Gradient", m_Gradient, m_def_Gradient)
    Call PropBag.WriteProperty("GradientAngle", m_GradientAngle, m_def_GradientAngle)
    Call PropBag.WriteProperty("GradientColor1", m_GradientColor1, m_def_GradientColor1)
    Call PropBag.WriteProperty("GradientColor1Opacity", m_GradientColor1Opacity, m_def_GradientColor1Opacity)
    Call PropBag.WriteProperty("GradientColor2", m_GradientColor2, m_def_GradientColor2)
    Call PropBag.WriteProperty("GradientColor2Opacity", m_GradientColor2Opacity, m_def_GradientColor2Opacity)
        Call PropBag.WriteProperty("GradientColorP1", m_GradientColorP1, m_def_GradientColorP1)
        Call PropBag.WriteProperty("GradientColorP1Opacity", m_GradientColorP1Opacity, m_def_GradientColorP1Opacity)
        Call PropBag.WriteProperty("GradientColorP2", m_GradientColorP2, m_def_GradientColorP2)
        Call PropBag.WriteProperty("GradientColorP2Opacity", m_GradientColorP2Opacity, m_def_GradientColorP2Opacity)
    
    Call PropBag.WriteProperty("ParentControl", m_ParentControl, m_def_ParentControl)
    Call PropBag.WriteProperty("Picture", m_Picture, m_def_Picture)
    Call PropBag.WriteProperty("PictureAlignmentH", m_PictureAlignmentH, m_def_PictureAlignmentH)
    Call PropBag.WriteProperty("PictureAlignmentV", m_PictureAlignmentV, m_def_PictureAlignmentV)
    Call PropBag.WriteProperty("PictureOpacity", m_PictureOpacity, m_def_PictureOpacity)
    Call PropBag.WriteProperty("PicturePadding", m_PicturePadding, m_def_PicturePadding)
    Call PropBag.WriteProperty("PictureSVGScale", m_PictureSVGScale, m_def_PictureSVGScale)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
    Call PropBag.WriteProperty("CaptionSubSizeMinus", m_FontMinus, 5)
    Call PropBag.WriteProperty("CaptionSubVDistance", m_C1VDistance, 20)
  Call PropBag.WriteProperty("Glowing", m_Glowing, m_def_Glowing)
End Sub

'===== PUBLIC PROPERTIES ===== =============================== PROPERTIES ====================================== PROPERTIES ================================================

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get BackColorOpacity() As Integer
    BackColorOpacity = m_BackColorOpacity
End Property

Public Property Let BackColorOpacity(ByVal New_BackColorOpacity As Integer)
    Select Case New_BackColorOpacity
        Case Is > 100
            New_BackColorOpacity = 100
        
        Case Is < 0
            New_BackColorOpacity = 0
    End Select
    
    m_BackColorOpacity = New_BackColorOpacity
    PropertyChanged "BackColorOpacity"
    
    Refresh
End Property

'<<------------------------------------------------------->>
Public Property Get BackColorPress() As OLE_COLOR
    BackColorPress = m_BackColorP
End Property

Public Property Let BackColorPress(ByVal New_BackColorP As OLE_COLOR)
    m_BackColorP = New_BackColorP
    PropertyChanged "BackColorPress"
    Refresh
End Property

Public Property Get BackColorPressOpacity() As Integer
    BackColorPressOpacity = m_BackColorOpacity
End Property

Public Property Let BackColorPressOpacity(ByVal New_BackColorPOpacity As Integer)
    Select Case New_BackColorPOpacity
        Case Is > 100
            New_BackColorPOpacity = 100
        
        Case Is < 0
            New_BackColorPOpacity = 0
    End Select
    
    m_BackColorPOpacity = New_BackColorPOpacity
    PropertyChanged "BackColorPressOpacity"
    
    Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    Refresh
End Property
'<<--------------------------------------------->>
Public Property Get BorderColorOnMouseOver() As OLE_COLOR
    BorderColorOnMouseOver = m_BorderColorP
End Property

Public Property Let BorderColorOnMouseOver(ByVal New_BorderColorP As OLE_COLOR)
    m_BorderColorP = New_BorderColorP
    PropertyChanged "BorderColorOnMouseOver"
    Refresh
End Property

Public Property Get BorderColorOpacity() As Integer
    BorderColorOpacity = m_BorderColorOpacity
End Property

Public Property Let BorderColorOpacity(ByVal New_BorderColorOpacity As Integer)
    Select Case New_BorderColorOpacity
        Case Is > 100
            New_BorderColorOpacity = 100
        
        Case Is < 0
            New_BorderColorOpacity = 0
    End Select
    
    m_BorderColorOpacity = New_BorderColorOpacity
    PropertyChanged "BorderColorOpacity"
    
    
    Refresh
End Property

Public Property Get BorderColorOpacityOnMouseOver() As Integer
    BorderColorOpacityOnMouseOver = m_BorderColorPOpacity
End Property

Public Property Let BorderColorOpacityOnMouseOver(ByVal New_BorderColorPOpacity As Integer)
    Select Case New_BorderColorPOpacity
        Case Is > 100
            New_BorderColorPOpacity = 100
        
        Case Is < 0
            New_BorderColorPOpacity = 0
    End Select
    
    m_BorderColorPOpacity = New_BorderColorPOpacity
    PropertyChanged "BorderColorOpacityOnMouseOver"
    
    Refresh
End Property
'<<--------------------------------------------->>

Public Property Get BorderCorner() As BorderCorner
    BorderCorner = m_BorderCorner
End Property

Public Property Let BorderCorner(ByVal New_BorderCorner As BorderCorner)
    m_BorderCorner = New_BorderCorner
    PropertyChanged "BorderCorner"

    Refresh
End Property

Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    PropertyChanged "Border"

    Refresh
End Property

Public Property Get BorderPosition() As BorderPosition
    BorderPosition = m_BorderPosition
End Property

Public Property Let BorderPosition(ByVal New_BorderPosition As BorderPosition)
    m_BorderPosition = New_BorderPosition
    PropertyChanged "BorderPosition"

    Refresh
End Property

Public Property Get BorderRadius() As Integer
    BorderRadius = m_BorderRadius
End Property

Public Property Let BorderRadius(ByVal New_BorderRadius As Integer)
    Select Case New_BorderRadius
        Case Is < 0
            New_BorderRadius = 0
    End Select
    
    m_BorderRadius = New_BorderRadius
    PropertyChanged "BorderRadius"
    
    
    Refresh
End Property

Public Property Get BorderSmoothEdge() As Boolean
    BorderSmoothEdge = m_BorderSmoothEdge
End Property

Public Property Let BorderSmoothEdge(ByVal New_BorderSmoothEdge As Boolean)
    m_BorderSmoothEdge = New_BorderSmoothEdge
    PropertyChanged "BorderSmoothEdge"

    Refresh
End Property

Public Property Get BorderWidth() As Integer
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    Select Case New_BorderWidth
        Case Is > 100
            New_BorderWidth = 100
            
        Case Is < 0
            New_BorderWidth = 0
    End Select
    
    m_BorderWidth = New_BorderWidth
    PropertyChanged "BorderWidth"
    m_OldBorderWidth = m_BorderWidth
    Refresh
End Property

Public Property Get CaptionSub() As String
    CaptionSub = m_Caption1
End Property

Public Property Let CaptionSub(ByVal New_Caption As String)
    m_Caption1 = New_Caption
    PropertyChanged "CaptionSub"

    Refresh
    Call Change(m_Caption1)
End Property

Public Property Get CaptionSubSizeMinus() As Integer
  CaptionSubSizeMinus = m_FontMinus
End Property

Public Property Let CaptionSubSizeMinus(ByVal newSize As Integer)
  m_FontMinus = newSize
  PropertyChanged "CaptionSubSizeMinus"
  Refresh
End Property

Public Property Get CaptionSubVDistance() As Integer
  CaptionSubVDistance = m_C1VDistance
End Property

Public Property Let CaptionSubVDistance(ByVal newDist As Integer)
  m_C1VDistance = newDist
  PropertyChanged "CaptionSubVDistance"
  Refresh
End Property

Public Property Get CaptionMain() As String
    CaptionMain = m_Caption2
End Property

Public Property Let CaptionMain(ByVal New_Caption As String)
    m_Caption2 = New_Caption
    PropertyChanged "CaptionMain"

    Refresh
    Call Change(m_Caption2)
End Property

Public Property Get CaptionAlignmentH() As CaptionAlignmentH
    CaptionAlignmentH = m_CaptionAlignmentH
End Property

Public Property Let CaptionAlignmentH(ByVal New_CaptionAlignmentH As CaptionAlignmentH)
    m_CaptionAlignmentH = New_CaptionAlignmentH
    PropertyChanged "CaptionAlignmentH"

    Refresh
End Property

Public Property Get CaptionAlignmentV() As CaptionAlignmentV
    CaptionAlignmentV = m_CaptionAlignmentV
End Property

Public Property Let CaptionAlignmentV(ByVal New_CaptionAlignmentV As CaptionAlignmentV)
    m_CaptionAlignmentV = New_CaptionAlignmentV
    PropertyChanged "CaptionAlignmentV"

    Refresh
End Property

Public Property Get CaptionPadding() As Integer
    CaptionPadding = m_CaptionPadding
End Property

Public Property Let CaptionPadding(ByVal New_CaptionPadding As Integer)
    Select Case New_CaptionPadding
        Case Is < 1
            New_CaptionPadding = 1
    End Select
    
    m_CaptionPadding = New_CaptionPadding
    PropertyChanged "CaptionPadding"
    
    Refresh
End Property
'ChangeBorderOnMouseOver
Public Property Get ChangeBorderColorOnMouseOver() As Boolean
    ChangeBorderColorOnMouseOver = m_ChangeBorderColorOnMouseOver
End Property

Public Property Let ChangeBorderColorOnMouseOver(ByVal New_Change As Boolean)
    m_ChangeBorderColorOnMouseOver = New_Change
    PropertyChanged "ChangeBorderColorOnMouseOver"
    Refresh
End Property

Public Property Get ChangeColorOnClick() As Boolean
    ChangeColorOnClick = m_ChangeColorOnClick
End Property

Public Property Let ChangeColorOnClick(ByVal New_Change As Boolean)
    m_ChangeColorOnClick = New_Change
    PropertyChanged "ChangeColorOnClick"
    Refresh
End Property

Public Property Get Cursor() As Cursor
    Cursor = m_Cursor
End Property

Public Property Let Cursor(ByVal New_Cursor As Cursor)
    m_Cursor = New_Cursor
    PropertyChanged "Cursor"

    Refresh
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    UserControl.Enabled = NewValue
End Property

Public Property Get FontAwesome() As String
    Dim B As Byte, i As Long, Hex() As Byte
    
    If UBound(m_Caption2) >= 1 Then
        ReDim Hex(UBound(m_Caption2) * 4 + 3)
        For i = 0 To UBound(m_Caption2)
            B = (m_Caption2(i) And &HF0) \ &H10
            Select Case B
            Case 0 To 9
                Hex(i * 4) = B Or 48
            Case 10 To 15
                Hex(i * 4) = B + 55
            End Select
            B = m_Caption2(i) And &HF
            Select Case B
            Case 0 To 9
                Hex(i * 4 + 2) = B Or 48
            Case 10 To 15
                Hex(i * 4 + 2) = B + 55
            End Select
        Next i
        FontAwesome = Hex
    End If
End Property

Public Property Let FontAwesome(New_FontAwesome As String)
    Dim B As Byte, i As Long, Hex() As Byte, N As Byte
    
    Hex = UCase$(New_FontAwesome)
    If UBound(Hex) >= 3 Then
        ReDim m_Caption2(0 To (UBound(Hex) - 3) \ 4)
        For i = 0 To UBound(m_Caption2)
            B = Hex(i * 4)
            Select Case B
            Case 48 To 57
                N = (B - 48) * &H10
            Case 65 To 70
                N = (B - 55) * &H10
            Case Else
                N = 0
            End Select
            B = Hex(i * 4 + 2)
            Select Case B
            Case 48 To 57
                N = N Or (B - 48)
            Case 65 To 70
                N = N Or (B - 55)
            End Select
            m_Caption2(i) = N
        Next i
    Else
        m_Caption2 = vbNullString
    End If

    Refresh
End Property

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(New_Font As StdFont)
    With m_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With
    PropertyChanged "Font"
    
    Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Refresh
End Property

''>>----------------------------------------
Public Property Get ForeColorOnPress() As OLE_COLOR
    ForeColorOnPress = m_ForeColorP
End Property

Public Property Let ForeColorOnPress(ByVal New_ForeColorP As OLE_COLOR)
    m_ForeColorP = New_ForeColorP
    PropertyChanged "ForeColorOnPress"
    Refresh
End Property

Public Property Get ForeColorOnPressOpacity() As Integer
    ForeColorOnPressOpacity = m_ForeColorOpacity
End Property

Public Property Let ForeColorOnPressOpacity(ByVal New_ForeColorPOpacity As Integer)
    Select Case New_ForeColorPOpacity
        Case Is > 100
            New_ForeColorPOpacity = 100
        
        Case Is < 0
            New_ForeColorPOpacity = 0
    End Select
    
    m_ForeColorPOpacity = New_ForeColorPOpacity
    PropertyChanged "ForeColorOnPressOpacity"
    
    Refresh
End Property
''<<----------------------------------------

Public Property Get ForeColorOpacity() As Integer
    ForeColorOpacity = m_ForeColorOpacity
End Property

Public Property Let ForeColorOpacity(ByVal New_ForeColorOpacity As Integer)
    Select Case New_ForeColorOpacity
        Case Is > 100
            New_ForeColorOpacity = 100
        
        Case Is < 0
            New_ForeColorOpacity = 0
    End Select
    
    m_ForeColorOpacity = New_ForeColorOpacity
    PropertyChanged "ForeColorOpacity"
    
    
    Refresh
End Property

Public Property Get GradientAngle() As Integer
    GradientAngle = m_GradientAngle
End Property

Public Property Let GradientAngle(ByVal New_GradientAngle As Integer)
    Select Case New_GradientAngle
        Case Is > 359
            New_GradientAngle = 359
        
        Case Is < 0
            New_GradientAngle = 0
    End Select
    
    m_GradientAngle = New_GradientAngle
    PropertyChanged "GradientAngle"
    
    
    Refresh
End Property

Public Property Get GradientColor1() As OLE_COLOR
    GradientColor1 = m_GradientColor1
End Property

Public Property Let GradientColor1(ByVal New_GradientColor1 As OLE_COLOR)
    m_GradientColor1 = New_GradientColor1
    PropertyChanged "GradientColor1"

    Refresh
End Property

Public Property Get GradientColor1Opacity() As Integer
    GradientColor1Opacity = m_GradientColor1Opacity
End Property

Public Property Let GradientColor1Opacity(ByVal New_GradientColor1Opacity As Integer)
    Select Case New_GradientColor1Opacity
        Case Is > 100
            New_GradientColor1Opacity = 100
        
        Case Is < 0
            New_GradientColor1Opacity = 0
    End Select
    
    m_GradientColor1Opacity = New_GradientColor1Opacity
    PropertyChanged "GradientColor1Opacity"
    
    
    Refresh
End Property

Public Property Get GradientColor2() As OLE_COLOR
    GradientColor2 = m_GradientColor2
End Property

Public Property Let GradientColor2(ByVal New_GradientColor2 As OLE_COLOR)
    m_GradientColor2 = New_GradientColor2
    PropertyChanged "GradientColor2"

    Refresh
End Property

Public Property Get GradientColor2Opacity() As Integer
    GradientColor2Opacity = m_GradientColor2Opacity
End Property

Public Property Let GradientColor2Opacity(ByVal New_GradientColor2Opacity As Integer)
    Select Case New_GradientColor2Opacity
        Case Is > 100
            New_GradientColor2Opacity = 100
        
        Case Is < 0
            New_GradientColor2Opacity = 0
    End Select
    
    m_GradientColor2Opacity = New_GradientColor2Opacity
    PropertyChanged "GradientColor2Opacity"
    
    Refresh
End Property

'<<------------------------------------------------->>
Public Property Get GradientColorP1() As OLE_COLOR
    GradientColorP1 = m_GradientColorP1
End Property

Public Property Let GradientColorP1(ByVal New_GradientColorP1 As OLE_COLOR)
    m_GradientColorP1 = New_GradientColorP1
    PropertyChanged "GradientColorP1"
    Refresh
End Property

Public Property Get GradientColorP1Opacity() As Integer
    GradientColorP1Opacity = m_GradientColorP1Opacity
End Property

Public Property Let GradientColorP1Opacity(ByVal New_GradientColorP1Opacity As Integer)
    Select Case New_GradientColorP1Opacity
        Case Is > 100
            New_GradientColorP1Opacity = 100
        
        Case Is < 0
            New_GradientColorP1Opacity = 0
    End Select
    
    m_GradientColorP1Opacity = New_GradientColorP1Opacity
    PropertyChanged "GradientColorP1Opacity"
    
    Refresh
End Property

Public Property Get GradientColorP2() As OLE_COLOR
    GradientColorP2 = m_GradientColorP2
End Property

Public Property Let GradientColorP2(ByVal New_GradientColorP2 As OLE_COLOR)
    m_GradientColorP2 = New_GradientColorP2
    PropertyChanged "GradientColorP2"
    Refresh
End Property

Public Property Get GradientColorP2Opacity() As Integer
    GradientColorP2Opacity = m_GradientColorP2Opacity
End Property

Public Property Let GradientColorP2Opacity(ByVal New_GradientColorP2Opacity As Integer)
    Select Case New_GradientColorP2Opacity
        Case Is > 100
            New_GradientColorP2Opacity = 100
        
        Case Is < 0
            New_GradientColorP2Opacity = 0
    End Select
    
    m_GradientColorP2Opacity = New_GradientColorP2Opacity
    PropertyChanged "GradientColorP2Opacity"
    
    Refresh
End Property

Public Property Get Gradient() As Boolean
    Gradient = m_Gradient
End Property

Public Property Let Gradient(ByVal New_Gradient As Boolean)
    m_Gradient = New_Gradient
    PropertyChanged "Gradient"

    Refresh
End Property

Public Property Get ParentControl() As String
    ParentControl = m_ParentControl
End Property

Public Property Let ParentControl(New_ParentControl As String)
    m_ParentControl = New_ParentControl
    PropertyChanged "ParentControl"
    Refresh
End Property

Public Property Get PictureAlignmentH() As PictureAlignmentH
    PictureAlignmentH = m_PictureAlignmentH
End Property

Public Property Let PictureAlignmentH(ByVal New_PictureAlignmentH As PictureAlignmentH)
    m_PictureAlignmentH = New_PictureAlignmentH
    PropertyChanged "PictureAlignmentH"

    Refresh
End Property

Public Property Get PictureAlignmentV() As PictureAlignmentV
    PictureAlignmentV = m_PictureAlignmentV
End Property

Public Property Let PictureAlignmentV(ByVal New_PictureAlignmentV As PictureAlignmentV)
    m_PictureAlignmentV = New_PictureAlignmentV
    PropertyChanged "PictureAlignmentV"

    Refresh
End Property

Public Property Get Picture() As String
    Picture = m_Picture
End Property

Public Property Let Picture(New_Picture As String)
    m_Picture = New_Picture
    PropertyChanged "Picture"
    
    Refresh
End Property

Public Property Get PictureOpacity() As Integer
    PictureOpacity = m_PictureOpacity
End Property

Public Property Let PictureOpacity(ByVal New_PictureOpacity As Integer)
    Select Case New_PictureOpacity
        Case Is > 100
            New_PictureOpacity = 100
        
        Case Is < 0
            New_PictureOpacity = 0
    End Select
    
    m_PictureOpacity = New_PictureOpacity
    PropertyChanged "PictureOpacity"
    
    
    Refresh
End Property

Public Property Get PicturePadding() As Integer
    PicturePadding = m_PicturePadding
End Property

Public Property Let PicturePadding(ByVal New_PicturePadding As Integer)
    Select Case New_PicturePadding
        Case Is < 0
            New_PicturePadding = 0
    End Select
    
    m_PicturePadding = New_PicturePadding
    PropertyChanged "PicturePadding"
    
    
    Refresh
End Property

Public Property Get PictureSVGScale() As Long
    PictureSVGScale = m_PictureSVGScale
End Property

Public Property Let PictureSVGScale(ByVal New_PictureSVGScale As Long)
    Select Case New_PictureSVGScale
        Case Is < 1
            New_PictureSVGScale = m_PictureSVGScale
    End Select
    
    m_PictureSVGScale = New_PictureSVGScale
    PropertyChanged "PictureSVGScale"
    
    Refresh
End Property

'===== PUBLIC GET PROPERTIES ===== ====================== PUBLIC GET PROPERTIES ========================== PUBLIC GET PROPERTIES ===========================================

Public Property Get RowCount() As Long
    RowCount = lngRowCount
End Property

Public Property Get WordCount() As Long
    WordCount = lngWordCount
End Property

Public Property Get WordWrap() As Boolean
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    m_WordWrap = New_WordWrap
    PropertyChanged "WordWrap"

    Refresh
End Property

'Value is not visible in Properties
Private Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    PropertyChanged "Value"

    If m_Value = True Then RaiseEvent Click
End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get Glowing() As Boolean
  Glowing = m_Glowing
End Property

Public Property Let Glowing(ByVal New_Glowing As Boolean)
  m_Glowing = New_Glowing
  PropertyChanged "Glowing"
  tmrGlow.Enabled = New_Glowing
  If Not New_Glowing Then m_BorderWidth = m_OldBorderWidth
End Property

