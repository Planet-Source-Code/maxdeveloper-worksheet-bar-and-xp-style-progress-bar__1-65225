VERSION 5.00
Begin VB.UserControl wbxSheetbar 
   Alignable       =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   ControlContainer=   -1  'True
   FillColor       =   &H80000010&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000010&
   HasDC           =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   495
   ScaleWidth      =   6630
   Begin VB.Timer dropt 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5550
      Top             =   60
   End
   Begin VB.HScrollBar hsc 
      Height          =   495
      Left            =   6210
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   360
   End
   Begin VB.PictureBox sw 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H80000010&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000010&
      Height          =   450
      Left            =   0
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   450
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "wbxSheetbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' wbxSheetbar
' Written and developed by (persiancity@gmail.com)
' Copyright 2000-2004, All rights reserved.
'
Option Explicit
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private m_Sheets As New Collection
Private m_props As New Collection
Private selsht As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Const FLOODFILLBORDER = 0
Private Const FLOODFILLSURFACE = 1
Private m_SheetWidth As Single
Private oX As Single
Private isIND As Boolean

Event SheetClick(Index As Long)
Event MouseOver(Index As Long)
Event SheetFocus(Index As Long)
Event Resize()

Private Declare Function GetFocus Lib "user32" () As Long
'Default Property Values:
Const m_def_RightSpace = 0
Const m_def_LeftSpace = 0
'Property Variables:
Dim m_RightSpace As Single
Dim m_LeftSpace As Single


Public Sub AddSheet(strCaption As String, Optional strSettings As String)
On Local Error GoTo errs1
m_props.Add IIf(strSettings = "", strCaption, strSettings), IIf(strSettings = "", "", strSettings)
m_Sheets.Add strCaption
redraw_all
errs1:
End Sub
Private Sub check_mover()
If sw.Left > hsc.Left Or sw.Left < 0 Or sw.Width > hsc.Left Then
hsc.Visible = True
Else
hsc.Visible = False
End If
End Sub

Private Sub draw_fcous()
On Local Error GoTo errs1
Dim oX As Single, ic As Long
Dim rct As RECT
ic = selsht
oX = (ic - 1) * m_SheetWidth + m_LeftSpace
If oX < 0 Then oX = 0
sw.Font.Bold = True
rct.Left = ((oX + (m_SheetWidth - sw.TextWidth(getstr(m_Sheets(ic)))) / 2)) / Screen.TwipsPerPixelX - 1
rct.Right = (sw.TextWidth(getstr(m_Sheets(ic))) / Screen.TwipsPerPixelX) + rct.Left + 3
rct.Top = ((sw.ScaleHeight - sw.TextHeight(" ")) / 2) / Screen.TwipsPerPixelY - 1
rct.Bottom = rct.Top + (sw.TextHeight(" ")) / Screen.TwipsPerPixelY + 3
DrawFocusRect sw.hdc, rct
errs1:
End Sub
Private Function getstr(strT As String) As String
If sw.TextWidth(strT) <= m_SheetWidth - 200 - (ScaleHeight / 2 / 50) * 2 Then getstr = strT: Exit Function
Dim ob As Boolean
ob = sw.Font.Bold

sw.Font.Bold = True
Dim t$, ic As Integer
For ic = 1 To Len(strT)
If sw.TextWidth(t$ + Mid(strT, ic, 1) & "...") <= m_SheetWidth - 200 - (ScaleHeight / 2 / 50) * 2 Then t$ = t$ & Mid(strT, ic, 1)
Next ic
sw.Font.Bold = ob

getstr = t$ & "..."
End Function

Public Sub RemoveSheet(Index As Long)
On Local Error GoTo errs1
m_Sheets.Remove Index
m_props.Remove Index
redraw_all
errs1:
End Sub
Public Function SelectSheet(Index As Long) As Long
On Local Error GoTo errs1
Dim t$
dropt = False
SelectSheet = selsht
t$ = m_Sheets(Index)
selsht = Index
redraw_all
errs1:
SelectSheet = selsht
End Function

Public Function SheetCaption(Index As Long) As String
On Local Error GoTo errs1
SheetCaption = m_Sheets(Index)
errs1:
End Function
Public Function SheetSettings(Index As Long) As String
On Local Error GoTo errs1
SheetSettings = m_props(Index)
errs1:
End Function

Public Sub ChangeSheet(Index As Long, NewCaption As String, NewSetting As String)
On Local Error GoTo errs1
m_props.Add IIf(NewSetting = "", NewCaption, NewSetting), IIf(NewSetting = "", "", NewSetting), , Index
m_Sheets.Add NewCaption, , , Index
RemoveSheet Index
errs1:
End Sub

Private Sub redraw_all()

On Local Error Resume Next
Dim ic As Long
Dim oX As Single
Dim selc As Long, rc As Long
isIND = True
hsc.SmallChange = hsc.Left \ m_SheetWidth + 1
hsc.LargeChange = hsc.Left \ m_SheetWidth + 1
hsc.Max = m_Sheets.Count
hsc.Min = 0
If hsc.Value <> selsht Then hsc.Value = selsht
check_mover
isIND = False
sw.Cls
sw.BackColor = UserControl.BackColor
sw.PSet (1, 1), vbWindowBackground
selc = sw.Point(1, 1)
sw.PSet (1, 1), sw.BackColor

sw.Width = ((m_Sheets.Count) * m_SheetWidth) + 35 + m_LeftSpace

For ic = 1 To m_Sheets.Count
oX = (ic - 1) * m_SheetWidth
If oX < 0 Then oX = 0
oX = oX + 15 + m_LeftSpace

sw.ForeColor = vbWhite
sw.Line (oX + 50, 0)-(oX + m_SheetWidth - 50, 0)
sw.Line (oX + m_SheetWidth - 50, 0)-(oX + m_SheetWidth, sw.ScaleHeight + 15)
sw.Line (oX + m_SheetWidth, sw.ScaleHeight)-(oX, sw.ScaleHeight + 15)
sw.Line (oX, sw.ScaleHeight)-(oX + 50, 0)

If selsht = ic Then
sw.Font.Bold = True

sw.FillColor = selc
ExtFloodFill sw.hdc, (oX / Screen.TwipsPerPixelX) + 1, (sw.ScaleHeight / Screen.TwipsPerPixelY) - 1, sw.ForeColor, FLOODFILLBORDER

sw.Line (oX + 50, 0)-(oX + m_SheetWidth - 50, 0), vb3DDKShadow 'vbButtonShadow
sw.Line (oX + m_SheetWidth - 50, 0)-(oX + m_SheetWidth, sw.ScaleHeight + 15), vb3DDKShadow
sw.Line (oX + m_SheetWidth, sw.ScaleHeight)-(oX, sw.ScaleHeight + 15), vbButtonShadow
sw.Line (oX, sw.ScaleHeight)-(oX + 50, 0), vb3DDKShadow 'vbButtonShadow

sw.Line (oX + 65, 15)-(oX + m_SheetWidth - 65, 15), UserControl.BackColor   'vbButtonFace
sw.Line (oX + m_SheetWidth - 65, 15)-(oX + m_SheetWidth - 15, sw.ScaleHeight + 15), UserControl.BackColor  'vbButtonFace
'sw.Line (oX + m_SheetWidth, sw.ScaleHeight - 15)-(oX, sw.ScaleHeight - 15), vbButtonShadow
sw.Line (oX + 15, sw.ScaleHeight + 15)-(oX + 65, 15), UserControl.BackColor  'vbButtonFace

Else
sw.Font.Bold = False

sw.Line (oX + 65, 15)-(oX + m_SheetWidth - 65, 15), vb3DHighlight
sw.Line (oX + 15, sw.ScaleHeight)-(oX + 65, 15), vb3DHighlight
sw.Line (oX + m_SheetWidth - 65, 0)-(oX + m_SheetWidth - 15, sw.ScaleHeight + 15), vbButtonShadow

sw.Line (oX + 50, 0)-(oX + m_SheetWidth - 50, 0), vbButtonShadow
sw.Line (oX + m_SheetWidth - 50, 0)-(oX + m_SheetWidth, sw.ScaleHeight + 15), vb3DDKShadow
'sw.Line (oX + m_SheetWidth, sw.ScaleHeight)-(oX, sw.ScaleHeight), vbButtonShadow
sw.Line (oX, sw.ScaleHeight)-(oX + 50, 0), vbButtonShadow

End If

sw.ForeColor = vbWindowText
sw.CurrentX = oX + (m_SheetWidth - sw.TextWidth(getstr(m_Sheets(ic)))) / 2
sw.CurrentY = (sw.ScaleHeight - sw.TextHeight(" ")) / 2
sw.Print getstr(m_Sheets(ic))
Next ic


If GetFocus = hwnd Or GetFocus = sw.hwnd Then draw_fcous
Refresh
End Sub

Private Sub dropt_Timer()
dropt = False
On Local Error GoTo errs1
sw_Click
errs1:
End Sub

Private Sub hsc_Change()
Dim wl As Single, nl As Single
nl = (hsc.Value * m_SheetWidth)
If sw.Left + nl < m_SheetWidth Then
wl = -((hsc.Value * m_SheetWidth) - m_SheetWidth)
If wl > 0 Then wl = 0
sw.Left = wl
ElseIf sw.Left + nl > hsc.Left Then
wl = -nl + hsc.Left - m_SheetWidth
sw.Left = wl
End If
hsc.SmallChange = hsc.Left \ m_SheetWidth + 1
hsc.LargeChange = hsc.Left \ m_SheetWidth + 1
Refresh
End Sub


Private Sub hsc_Scroll()
hsc_Change
End Sub

Private Sub sw_Click()
On Local Error Resume Next
Dim ic As Long
ic = ((oX - m_LeftSpace) / m_SheetWidth) + 0.5
If SelectSheet(ic) = ic Then RaiseEvent SheetClick(ic)
End Sub

Private Sub sw_GotFocus()
RaiseEvent SheetFocus(selsht)
redraw_all
End Sub

Private Sub sw_LostFocus()
redraw_all
End Sub


Private Sub sw_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo errs1
Dim ic As Long
ic = ((X - m_LeftSpace) / m_SheetWidth) + 0.5
sw.ToolTipText = Extender.ToolTipText
RaiseEvent MouseOver(ic)
errs1:
End Sub

Private Sub sw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
oX = X
Else
oX = sw.Left - 100
End If
End Sub

Private Sub sw_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
If X = 0 Then
dropt = False
Else
oX = X
dropt = True
End If
End Sub

Private Sub UserControl_Click()
RaiseEvent SheetClick(-1)
End Sub

Private Sub UserControl_GotFocus()
redraw_all
End Sub

Private Sub UserControl_InitProperties()
m_SheetWidth = 1024
    m_RightSpace = m_def_RightSpace
    m_LeftSpace = m_def_LeftSpace
End Sub

Private Sub UserControl_LostFocus()
redraw_all
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseOver(-1)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_SheetWidth = PropBag.ReadProperty("SheetWidth", 1024)
    Set sw.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_RightSpace = PropBag.ReadProperty("RightSpace", m_def_RightSpace)
    m_LeftSpace = PropBag.ReadProperty("LeftSpace", m_def_LeftSpace)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
redraw_all
End Sub


Private Sub UserControl_Resize()
On Local Error Resume Next
hsc.Move ScaleWidth - hsc.Width - m_RightSpace, 0, hsc.Width, ScaleHeight
check_mover
hsc.SmallChange = hsc.Left \ m_SheetWidth + 1
hsc.LargeChange = hsc.Left \ m_SheetWidth + 1
sw.Height = ScaleHeight - 15
sw.Top = 15
redraw_all
RaiseEvent Resize
End Sub


Public Function SheetCount()
SheetCount = m_Sheets.Count
End Function

Public Sub SelectFromSettings(strSettings As String)
On Local Error GoTo errs1
Dim ic As Long
For ic = 1 To m_props.Count
If m_props(ic) = strSettings Then GoTo ok2
Next ic
GoTo errs1
ok2:
SelectSheet ic
errs1:
End Sub
Public Sub RemoveFromSettings(strSettings As String)
On Local Error GoTo errs1
Dim ic As Long
For ic = 1 To m_props.Count
If m_props(ic) = strSettings Then GoTo ok2
Next ic
GoTo errs1
ok2:
RemoveSheet ic
errs1:
End Sub


Public Property Get SheetWidth() As Single
SheetWidth = m_SheetWidth
End Property

Public Property Let SheetWidth(ByVal vNewValue As Single)
m_SheetWidth = vNewValue
redraw_all
PropertyChanged "SheetWidth"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SheetWidth", m_SheetWidth, 1024)
    Call PropBag.WriteProperty("Font", sw.Font, Ambient.Font)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("RightSpace", m_RightSpace, m_def_RightSpace)
    Call PropBag.WriteProperty("LeftSpace", m_LeftSpace, m_def_LeftSpace)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=sw,sw,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = sw.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set sw.Font = New_Font
    redraw_all
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get RightSpace() As Single
    RightSpace = m_RightSpace
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get LeftSpace() As Single
    LeftSpace = m_LeftSpace
End Property

Public Property Let RightSpace(ByVal New_RightSpace As Single)
    m_RightSpace = New_RightSpace
    hsc.Move ScaleWidth - hsc.Width - m_RightSpace, 0, hsc.Width, ScaleHeight
    check_mover
    hsc.SmallChange = hsc.Left \ m_SheetWidth + 1
    hsc.LargeChange = hsc.Left \ m_SheetWidth + 1
    sw.Height = ScaleHeight - 15
    sw.Top = 15
    redraw_all
    PropertyChanged "RightSpace"
End Property

Public Property Let LeftSpace(ByVal New_LeftSpace As Single)
    m_LeftSpace = New_LeftSpace
    check_mover
    redraw_all
    PropertyChanged "LeftSpace"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    redraw_all
    PropertyChanged "BackColor"
End Property

