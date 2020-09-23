VERSION 5.00
Begin VB.UserControl wbxProgress 
   AutoRedraw      =   -1  'True
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   FillColor       =   &H8000000D&
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
End
Attribute VB_Name = "wbxProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Workbox Frame 10XPT
' XP Theme Ready.
'
' Written and developed by (persiancity@gmail.com)
' Copyright 2000-2004, All rights reserved.
'
'
Option Explicit

' XP Theme
Private hTheme As Long, isThemeEnabled As Boolean, isNoXP As Boolean
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseTheme Lib "uxtheme.dll" Alias "CloseThemeData" (ByVal hTheme As Long) As Long
Private Declare Function IsThemeActive Lib "uxtheme.dll" () As Boolean
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Boolean
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, pBoundingRect As Any, pContentRect As Any) As Long
Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Private Declare Function GetThemeMetric Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, ByVal piVal As Long) As Long
Private Const PROGRESSCHUNKSIZE = 2411 '// size of progress control chunks

Private Enum xpParts
BAR = 1
BARVERT = 2
Chunk = 3
CHUNKVERT = 4
End Enum

' /XP Theme

Enum wbeProgressStyle
wbeNone = 0
wbeFixed = 1
End Enum

Enum wbeProgressModel
wbeNormal = 0
wbeXPTheme = 1
End Enum

Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

Const m_def_ProgressModel = 0
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Value = 50
Const m_def_ShowPercent = True
Const m_def_IgnoreBadValue = True
Const m_def_ShowValue = True

Dim m_ProgressModel As Integer
Dim m_Min As Long
Dim m_Max As Long
Dim m_Value As Long
Dim m_ShowPercent As Boolean
Dim m_IgnoreBadValue As Boolean
Dim m_ShowValue As Boolean



Private Sub load_theme(isOpen As Boolean)
On Local Error GoTo errs1
If isOpen Then
Dim ptrString As Long
ptrString = StrPtr("Progress")
If hTheme = 0 Then hTheme = OpenThemeData(ByVal hwnd, ByVal ptrString)
Else
If hTheme <> 0 Then Call CloseTheme(hTheme): hTheme = 0
End If
errs1:
End Sub

Public Property Get ProgressStyle() As wbeProgressStyle
    ProgressStyle = UserControl.BorderStyle
End Property

Public Property Let ProgressStyle(ByVal New_ProgressStyle As wbeProgressStyle)
    UserControl.BorderStyle() = New_ProgressStyle
    draw_all
    PropertyChanged "ProgressStyle"
End Property

Public Property Get Min() As Long
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Long)
    m_Min = New_Min
    draw_all
    PropertyChanged "Min"
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Long)
    m_Max = New_Max
    draw_all
    PropertyChanged "Max"
End Property

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Long)
    m_Value = New_Value
    draw_all
    PropertyChanged "Value"
End Property

Public Property Get ShowPercent() As Boolean
    ShowPercent = m_ShowPercent
End Property

Public Property Let ShowPercent(ByVal New_ShowPercent As Boolean)
    m_ShowPercent = New_ShowPercent
    draw_all
    PropertyChanged "ShowPercent"
End Property

Public Property Get IgnoreBadValue() As Boolean
    IgnoreBadValue = m_IgnoreBadValue
End Property

Public Property Let IgnoreBadValue(ByVal New_IgnoreBadValue As Boolean)
    m_IgnoreBadValue = New_IgnoreBadValue
    draw_all
    PropertyChanged "IgnoreBadValue"
End Property

Public Property Get ShowValue() As Boolean
    ShowValue = m_ShowValue
End Property

Public Property Let ShowValue(ByVal New_ShowValue As Boolean)
    m_ShowValue = New_ShowValue
    draw_all
    PropertyChanged "ShowValue"
End Property

Private Sub UserControl_Initialize()
On Local Error Resume Next
isThemeEnabled = (IsThemeActive And IsAppThemed)
isNoXP = Not isThemeEnabled
End Sub

Private Sub UserControl_InitProperties()
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Value = m_def_Value
    m_ShowPercent = m_def_ShowPercent
    m_IgnoreBadValue = m_def_IgnoreBadValue
    m_ShowValue = m_def_ShowValue
    m_ProgressModel = m_def_ProgressModel
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BorderStyle = PropBag.ReadProperty("ProgressStyle", 0)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_ShowPercent = PropBag.ReadProperty("ShowPercent", m_def_ShowPercent)
    m_IgnoreBadValue = PropBag.ReadProperty("IgnoreBadValue", m_def_IgnoreBadValue)
    m_ShowValue = PropBag.ReadProperty("ShowValue", m_def_ShowValue)
    m_ProgressModel = PropBag.ReadProperty("ProgressModel", m_def_ProgressModel)
    If m_ProgressModel = 0 Then isNoXP = True Else If isThemeEnabled Then isNoXP = False Else isNoXP = True

draw_all
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.FillColor = PropBag.ReadProperty("FillColor", &H8000000D)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
End Sub
Private Sub draw_all()
On Local Error Resume Next
Dim hr As Long, rct As RECT, rcContent As RECT
Dim sw As Long, dws As Long, chs As Long
Dim stt As String, wd As Long


If Not isNoXP And m_ProgressModel = wbeXPTheme Then
load_theme True
If UserControl.BorderStyle <> 0 Then UserControl.BorderStyle = 0
rct.Right = ScaleWidth
rct.Bottom = ScaleHeight

Cls
sw = xpParts.BAR
hr = DrawThemeBackground(ByVal hTheme, _
     ByVal hdc, sw, _
     dws, rct, rct)
hr = GetThemeBackgroundContentRect(hTheme, _
    ByVal hdc, sw, _
     dws, rct, rcContent)
sw = xpParts.Chunk
GetThemeMetric ByVal hTheme, ByVal hdc, ByVal sw, ByVal dws, _
    ByVal PROGRESSCHUNKSIZE, ByVal VarPtr(chs)

rct.Left = rcContent.Left
rct.Bottom = rcContent.Bottom
rct.Top = rcContent.Top

wd = (rcContent.Right - rcContent.Left) / chs
rct.Right = ((wd * m_Value) \ (m_Max - m_Min)) * chs

hr = DrawThemeBackground(ByVal hTheme, _
     ByVal hdc, sw, _
     dws, rct, rcContent)

If m_ShowValue Then
stt = m_Value
If m_Value > m_Max Then stt = m_Max
If m_Value < m_Min Then stt = m_Min
If m_ShowPercent Then stt = stt & "%"

rcContent.Left = (rcContent.Right - rcContent.Left - TextWidth(stt)) / 2
rcContent.Top = (rcContent.Bottom - rcContent.Top - TextHeight(stt)) / 2
     
hr = DrawThemeText(ByVal hTheme, ByVal hdc, _
     sw, dws, _
     ByVal StrPtr(stt), Len(stt), _
     &H0& Or &H0& Or &H20&, _
     IIf(UserControl.Enabled, 0, &H1&), rcContent)
End If
Else
load_theme False

stt = m_Value
If m_Value > m_Max Then stt = m_Max
If m_Value < m_Min Then stt = m_Min
If m_ShowPercent Then stt = stt & "%"

wd = (ScaleWidth * m_Value) / (m_Max - m_Min)
Cls
If m_Value > m_Min Then
DrawMode = 4
Line (0, 0)-(wd, Height), FillColor, BF
DrawMode = 4
Line (0, 0)-(wd, Height), FillColor, BF
CurrentY = (ScaleHeight - TextHeight(stt)) / 2
CurrentX = (ScaleWidth - TextWidth(stt)) / 2
If m_ShowValue Then Print stt
DrawMode = 14
Line (0, 0)-(wd, Height), FillColor, BF
End If
Refresh
End If
End Sub

Private Sub UserControl_Resize()
draw_all
End Sub

Private Sub UserControl_Terminate()
load_theme False

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ProgressStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("ShowPercent", m_ShowPercent, m_def_ShowPercent)
    Call PropBag.WriteProperty("IgnoreBadValue", m_IgnoreBadValue, m_def_IgnoreBadValue)
    Call PropBag.WriteProperty("ShowValue", m_ShowValue, m_def_ShowValue)
    Call PropBag.WriteProperty("ProgressModel", m_ProgressModel, m_def_ProgressModel)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("FillColor", UserControl.FillColor, &H8000000D)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
End Sub

Public Property Get ProgressModel() As wbeProgressModel
    ProgressModel = m_ProgressModel
End Property

Public Property Let ProgressModel(ByVal New_ProgressModel As wbeProgressModel)
    m_ProgressModel = New_ProgressModel
    draw_all
    PropertyChanged "ProgressModel"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get FillColor() As OLE_COLOR
    FillColor = UserControl.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    UserControl.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

