VERSION 5.00
Begin VB.UserControl TPanel 
   Alignable       =   -1  'True
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   ControlContainer=   -1  'True
   ScaleHeight     =   1965
   ScaleWidth      =   3720
   ToolboxBitmap   =   "TPanel.ctx":0000
End
Attribute VB_Name = "TPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : TPanel
'    Project    : CSSplitter
'    Created By : Project Administrator
'    Description: This Panel Control was based on the work of
'                 Houston McClung and Resize Aligned Usercontrol Template in PSC and
'                 Thomas Kabir with Light3D control  http://www.vbfrood.de
'
'    Modified   : 17/4/2004 10:48:57
'--------------------------------------------------------------------------------
'</CSCC>
Option Explicit

Public Enum PanelBackStyleConstants
    PanelTransparent = 0
    PanelOpaque = 1
End Enum

' panel border styles
Public Enum PanelBorderStyleConstants
    PanelBorderNone = 0
    PanelBorderRaisedOuter = 1
    PanelBorderRaisedInner = 2
    PanelBorderRaised = 3
    PanelBorderSunkenOuter = 4
    PanelBorderSunkenInner = 5
    PanelBorderSunken = 6
    PanelBorderEtched = 7
    PanelBorderBump = 8
    PanelBorderMono = 9
    PanelBorderFlat = 10
    PanelBorderSoft = 11
End Enum

'Property Variables:
Private m_BorderStyle As PanelBorderStyleConstants
Private m_bDirty As Boolean

'Event Declarations:
Public Event Resize()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Public Property Get ScaleWidth() As Long
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Get ScaleHeight() As Long
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get BackStyle() As PanelBackStyleConstants
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As PanelBackStyleConstants)
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set UserControl.Picture = New_Picture
    UserControl.Palette = UserControl.Picture
    PropertyChanged "Picture"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Sub Cls()
    UserControl.Cls
End Sub

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

Public Property Get AutoRedraw() As Boolean
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    UserControl.Refresh
    PropertyChanged "BackColor"
End Property

Public Sub PSet_(X As Single, Y As Single, Color As Long)
    UserControl.PSet (X + 1, Y + 1), Color
End Sub

Public Property Get Image() As Picture
    Set Image = UserControl.Image
End Property

Public Property Get BorderStyle() As PanelBorderStyleConstants
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_Style As PanelBorderStyleConstants)
    m_BorderStyle = New_Style
    PropertyChanged "Style"
    UserControl_Paint
End Property

Private Sub UserControl_InitProperties()
    Set Font = Ambient.Font
    m_BorderStyle = PanelBorderNone
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    BorderStyle = PropBag.ReadProperty("BorderStyle", PanelBorderNone)
    BackStyle = PropBag.ReadProperty("BackStyle", PanelOpaque)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, PanelOpaque)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, PanelBorderNone)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim nParam  As Long
    Dim ext As VBControlExtender
    
    With UserControl
        Set ext = .Extender
        Select Case ext.Align
            Case Is = 0     'not aligned
                If (Y > UserControl.Height - 100) Then
                    nParam = HTBOTTOM
                End If
                If (Y > 0 And Y < 100) Then
                    nParam = HTTOP
                End If
                If (X > 0 And X < 100) Then
                    nParam = HTLEFT
                End If
                If (X > UserControl.Width - 100) Then
                    nParam = HTRIGHT
                End If
                
            Case Is = 1     'top
                If (Y > UserControl.Height - 100) Then
                    nParam = HTBOTTOM
                End If
                
            Case Is = 2     'bottom
                If (Y > 0 And Y < 100) Then
                    nParam = HTTOP
                End If
                
            Case Is = 3     'left
                If (X > UserControl.Width - 100) Then
                    nParam = HTRIGHT
                End If
                
            Case Is = 4     'right
                If (X > 0 And X < 100) Then
                    nParam = HTLEFT
                End If
        End Select
        Set ext = Nothing

        If nParam Then
            Call ReleaseCapture
            Call SendMessage(.hWnd, WM_NCLBUTTONDOWN, nParam, 0)
            UserControl_Resize
            If nParam = HTRIGHT Or nParam = HTLEFT Then UserControl.Width = UserControl.Width
            If nParam = HTTOP Or nParam = HTBOTTOM Then UserControl.Height = UserControl.Height
            UserControl_Resize
        End If
    End With
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    UserControl_Paint
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewPointer As MousePointerConstants
    Dim ext As VBControlExtender
    
    With UserControl
        Set ext = .Extender
        Select Case ext.Align
            Case Is = 0     'not aligned
                If (Y > 0 And Y < 100) Then
                    NewPointer = vbSizeNS
                Else
                    NewPointer = vbDefault
                End If
                If (Y > UserControl.Height - 100) Then
                    NewPointer = vbSizeNS
                Else
                    NewPointer = vbDefault
                End If
                If (X > UserControl.Width - 100) Then
                    NewPointer = vbSizeWE
                Else
                    NewPointer = vbDefault
                End If
                If (X > 0 And X < 100) Then
                    NewPointer = vbSizeWE
                Else
                    NewPointer = vbDefault
                End If
            
            Case Is = 1     'top
                If (Y > UserControl.Height - 100) Then
                    NewPointer = vbSizeNS
                Else
                    NewPointer = vbDefault
                End If
            
            Case Is = 2     'bottom
                If (Y > 0 And Y < 100) Then
                    NewPointer = vbSizeNS
                Else
                    NewPointer = vbDefault
                End If
            
            Case Is = 3     'left
                If (X > UserControl.Width - 100) Then
                    NewPointer = vbSizeWE
                Else
                    NewPointer = vbDefault
                End If
                
            Case Is = 4     'right
                If (X > 0 And X < 100) Then
                    NewPointer = vbSizeWE
                Else
                    NewPointer = vbDefault
                End If
        End Select
        Set ext = Nothing
    End With

    If NewPointer <> UserControl.MousePointer Then
        UserControl.MousePointer = NewPointer
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    If m_bDirty Then Exit Sub
    
    Dim rc As RECT
    Dim bdrStyle As Long
    Dim bdrSides As Long
    Dim Offset
    
    ' set this to true to avoid redrawing
    m_bDirty = True
    ' the split bar must be on top of other controls, so
    ' cal zorder method
    Extender.ZOrder
    ' all sides must be updated
    bdrSides = BF_RECT
    ' update border styles
    If m_BorderStyle = PanelBorderFlat Then bdrSides = bdrSides Or BF_FLAT
    If m_BorderStyle = PanelBorderMono Then bdrSides = bdrSides Or BF_MONO
    If m_BorderStyle = PanelBorderSoft Then bdrSides = bdrSides Or BF_SOFT
    Select Case m_BorderStyle
        Case PanelBorderRaisedOuter: bdrStyle = BDR_RAISEDOUTER
        Case PanelBorderRaisedInner: bdrStyle = BDR_RAISEDINNER
        Case PanelBorderRaised: bdrStyle = EDGE_RAISED
        Case PanelBorderSunkenOuter: bdrStyle = BDR_SUNKENOUTER
        Case PanelBorderSunkenInner: bdrStyle = BDR_SUNKENINNER
        Case PanelBorderSunken: bdrStyle = EDGE_SUNKEN
        Case PanelBorderEtched: bdrStyle = EDGE_ETCHED
        Case PanelBorderBump: bdrStyle = EDGE_BUMP
        Case PanelBorderFlat: bdrStyle = BDR_SUNKEN
        Case PanelBorderMono: bdrStyle = BDR_SUNKEN
        Case PanelBorderSoft: bdrStyle = BDR_RAISED
    End Select
    ' update rect
    rc.Left = 0
    rc.Top = 0
    rc.Bottom = CLng(Extender.Height / Screen.TwipsPerPixelY)
    rc.Right = CLng(Extender.Width / Screen.TwipsPerPixelX)
    ' clear controls contents
    UserControl.Cls
    ' Simply call the API and draw the edge.
    DrawEdge hDC, rc, bdrStyle, bdrSides
    ' clear flag
    m_bDirty = False
End Sub

Public Sub Refresh()
    UserControl_Paint
End Sub
'-- end code
