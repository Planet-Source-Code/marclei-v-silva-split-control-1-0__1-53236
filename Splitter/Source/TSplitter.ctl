VERSION 5.00
Begin VB.UserControl TSplitter 
   CanGetFocus     =   0   'False
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   150
   ScaleHeight     =   3180
   ScaleWidth      =   150
   ToolboxBitmap   =   "TSplitter.ctx":0000
End
Attribute VB_Name = "TSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ******************************************************************************
' Class         : TSplitter
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 13/12/00 19:00:31
'
' Credits       : To SP McMahon from cSplitDC class
'                 All the credits for the splitting functions
'                 routines and declarations using
'                 DC goes to SPMcMahon.
'                 The Draw edge operations in the Paint() method
'                 was extracted from a project downloaded from planet-
'                 source-code, but i can't remember the author name
'                 however many thanks to him :)
'
' Modifications : removed the splitting events from cSplitDC class
'                 and added the getKeyState() method to detect the end
'                 of splitting operation
'
' Description   : Splitter control that automatically splits
'                 objects added to its control list
' ******************************************************************************
Option Explicit
Option Compare Text

' define where the controls added to the list are
' positioned concerning the split bar position and
' orientation
Public Enum spControlPosition
    spLeft = 1
    spRight = 2
    spTop = 3
    spBottom = 4
    spClientLeft = 5
    spClientRight = 6
    spClientTop = 7
    spClientBottom = 8
End Enum

' Orientation of the split bar
Public Enum spOrientationConstants
    spVertical = 1
    spHorizontal = 2
End Enum

' Split bar border styles
Public Enum spBorderStyles
    bdrNone = 0
    bdrRaisedOuter = 1
    bdrRaisedInner = 2
    bdrRaised = 3
    bdrSunkenOuter = 4
    bdrSunkenInner = 5
    bdrSunken = 6
    bdrEtched = 7
    bdrBump = 8
    bdrMono = 9
    bdrFlat = 10
    bdrSoft = 11
End Enum

' This is from cSplitDC project and it is not used
' directly by the control but it is here for further use
Private Enum ESplitBorderTypes
   espbLeft = 1
   espbTop = 2
   espbRight = 3
   espbBottom = 4
End Enum

' some global declarations
Private bDraw As Boolean
Private rcCurrent As RECT
Private rcNew As RECT
Private rcWindow As RECT
Private rcSplit As RECT
Private m_bDirty As Boolean
Private m_hWnd As Long
Private m_lBorder(1 To 4) As Long
Private m_bSplitting As Boolean
Private m_Controls As TControls

' Default Property Values:
Const m_def_BorderStyle = 0
Const m_def_Orientation = 1
Const m_def_Thickness = 80
Const m_def_MinHeight = 1200
Const m_def_MinWidth = 1200
Const m_def_BorderSize = 0

' Property Variables:
Private m_BorderStyle As spBorderStyles
Private m_Appearance As Integer
Private m_Orientation As spOrientationConstants
Private m_Thickness As Integer
Private m_MinHeight As Long
Private m_MinWidth As Long
Private m_BorderSize As Long

' public events
Public Event StartMoving()
Public Event EndMoving()
Public Event Moving(ByVal X As Long, ByVal Y As Long)

Public Property Let BorderSize(ByVal New_BorderSize As Long)
Attribute BorderSize.VB_Description = "Returns or sets the split bar border size. "
    m_BorderSize = New_BorderSize
    If m_Orientation = spHorizontal Then
        m_lBorder(espbLeft) = New_BorderSize / Screen.TwipsPerPixelX
        m_lBorder(espbRight) = New_BorderSize / Screen.TwipsPerPixelX
        m_lBorder(espbTop) = 0
        m_lBorder(espbBottom) = 0
    Else
        m_lBorder(espbTop) = New_BorderSize / Screen.TwipsPerPixelY
        m_lBorder(espbBottom) = New_BorderSize / Screen.TwipsPerPixelY
        m_lBorder(espbLeft) = 0
        m_lBorder(espbRight) = 0
    End If
End Property

Public Property Get BorderSize() As Long
Attribute BorderSize.VB_Description = "Returns or sets the split bar border size. "
    BorderSize = m_BorderSize
End Property

Public Property Get Thickness() As Integer
Attribute Thickness.VB_Description = "Returns or sets the split bar thickness "
    Thickness = m_Thickness
End Property

Public Property Let Thickness(ByVal New_Thickness As Integer)
Attribute Thickness.VB_Description = "Returns or sets the split bar thickness "
    m_Thickness = New_Thickness
    If (m_Orientation = spHorizontal) Then
        UserControl.Height = m_Thickness
    Else
        UserControl.Width = m_Thickness
    End If
    PropertyChanged "Thickness"
End Property

Public Property Get BorderStyle() As spBorderStyles
Attribute BorderStyle.VB_Description = "Returns or sets the border style of the split bar"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As spBorderStyles)
Attribute BorderStyle.VB_Description = "Returns or sets the border style of the split bar"
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    UserControl_Paint
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Mouse pointer of the split bar "
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
Attribute MousePointer.VB_Description = "Mouse pointer of the split bar "
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get MinHeight() As Long
    MinHeight = m_MinHeight
End Property

Public Property Let MinHeight(ByVal New_MinHeight As Long)
    m_MinHeight = New_MinHeight
    PropertyChanged "MinHeight"
End Property

Public Property Get MinWidth() As Long
    MinWidth = m_MinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Long)
    m_MinWidth = New_MinWidth
    PropertyChanged "MinWidth"
End Property

Public Property Get Orientation() As spOrientationConstants
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal eOrientation As spOrientationConstants)
    m_Orientation = eOrientation
    ' check orientation value
    If (m_Orientation = spHorizontal) Then
        ' redefine cursor only if current value is not
        ' default or vbSizeWE
        If UserControl.MousePointer = vbDefault Or _
           UserControl.MousePointer = vbSizeWE Then
            UserControl.MousePointer = vbSizeNS
        End If
        ' set the thickness
        UserControl.Width = UserControl.Height
        UserControl.Height = m_Thickness
    Else
        ' redefine cursor only if current value is not
        ' default or vbSizeNS
        If UserControl.MousePointer = vbDefault Or _
           UserControl.MousePointer = vbSizeNS Then
            UserControl.MousePointer = vbSizeWE
        End If
        UserControl.Height = UserControl.Width
        UserControl.Width = m_Thickness
    End If
    PropertyChanged "Orientation"
End Property

Private Sub UserControl_AmbientChanged(PropertyName As String)
'    Debug.Print PropertyName
End Sub

Private Sub Usercontrol_Initialize()
    ' create controls object
    ' this collection will hold conntrol references
    ' and position concerning the splitter object
    Set m_Controls = New TControls
    ' initially set orientation to vertical (common use)
    ' do not use the variable "m_orientation" here, because we
    ' want to update orientation dependents variables
    Orientation = spVertical
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BorderStyle = m_def_BorderStyle
    m_Orientation = m_def_Orientation
    m_MinHeight = m_def_MinHeight
    m_MinWidth = m_def_MinWidth
    m_Thickness = m_def_Thickness
    m_BorderSize = m_def_BorderSize
End Sub

' load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_MinHeight = PropBag.ReadProperty("MinHeight", m_def_MinHeight)
    m_MinWidth = PropBag.ReadProperty("MinWidth", m_def_MinWidth)
    
    BorderSize = PropBag.ReadProperty("BorderSize", m_def_BorderSize)
    Thickness = PropBag.ReadProperty("Thickness", m_def_Thickness)
    
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    
    ' set initial orientation
    If m_Orientation = spVertical Then
        If Extender.Left < m_MinWidth Then
            Extender.Left = m_MinWidth
        End If
    Else
        If Extender.Top < m_MinHeight Then
            Extender.Top = m_MinHeight
        End If
    End If
End Sub

' write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("Thickness", m_Thickness, m_def_Thickness)
    Call PropBag.WriteProperty("MinHeight", m_MinHeight, m_def_MinHeight)
    Call PropBag.WriteProperty("MinWidth", m_MinWidth, m_def_MinWidth)
    Call PropBag.WriteProperty("BorderSize", m_BorderSize, m_def_BorderSize)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim tP As POINTAPI
    Dim tPPrev As POINTAPI
    Dim rcDrag As RECT
    
    ' raise event
    RaiseEvent StartMoving
    ' get owner handle
    m_hWnd = Extender.Container.hWnd
    ' send subsequent mouse messages to the owner window
    SetCapture m_hWnd
    ' get the window rectangle on the desktop of the owner window:
    GetWindowRect m_hWnd, rcWindow
    GetDragArea rcDrag
    ' *** Marclei V Silva
    ' * It was removed the MDI form checking here all windows
    ' * are treated as container the same way
    ' ***
    If m_Orientation = spVertical Then
        rcWindow.Left = rcWindow.Left + rcDrag.Left
        rcWindow.Right = rcWindow.Left + rcDrag.Right
    Else
        rcWindow.Top = rcWindow.Top + rcDrag.Top
        rcWindow.Bottom = rcWindow.Top + rcDrag.Bottom
    End If
    ' avoid some overlaping
    If rcWindow.Top > rcWindow.Bottom Then
        rcWindow.Top = rcWindow.Bottom
    End If
    If rcWindow.Left > rcWindow.Right Then
        rcWindow.Left = rcWindow.Right
    End If
    ' **** Marclei V Silva
    ' * get the rect from split control to update some propeerties further
    ' ***
    GetWindowRect UserControl.hWnd, rcSplit
    ' clip the cursor so it can't move outside the window:
    ClipCursorRect rcWindow
    ' Get the client rectangle of the window in screen coordinates:
    GetClientRect m_hWnd, rcWindow
    tP.X = rcWindow.Left
    tP.Y = rcWindow.Top
    ClientToScreen m_hWnd, tP
    rcWindow.Left = tP.X
    rcWindow.Top = tP.Y
    tP.X = rcWindow.Right
    tP.Y = rcWindow.Bottom
    ClientToScreen m_hWnd, tP
    rcWindow.Right = tP.X
    rcWindow.Bottom = tP.Y
    bDraw = True  ' start actual drawing from next move message
    rcCurrent.Left = 0
    rcCurrent.Top = 0
    rcCurrent.Right = 0
    rcCurrent.Bottom = 0
    X = (Extender.Left + X) \ Screen.TwipsPerPixelX
    Y = (Extender.Top + Y) \ Screen.TwipsPerPixelY
    ' store the initial cursor position
    tPPrev.X = tP.X
    tPPrev.Y = tP.Y
    SplitterMouseMove tP.X, tP.Y
    ' **** Marclei V Silva
    ' * added a syncronous splitting once we don't want to
    ' * handle splitting with parent window events
    ' ***
    Do While GetKeyState(VK_LBUTTON) < 0
        GetCursorPos tP
        If tP.X <> tPPrev.X Or tP.Y <> tPPrev.Y Then
            tPPrev.X = tP.X
            tPPrev.Y = tP.Y
            SplitterMouseMove tP.X, tP.Y
            RaiseEvent Moving(tP.X, tP.Y)
        End If
        DoEvents
    Loop
    ' release drawing
    SplitterMouseUp tP.X, tP.Y
    ' raise endmoving() event, so that the end user may
    ' resize their controls his way
    RaiseEvent EndMoving
    ' refresh all splitters
    RefreshAll
End Sub

Private Sub SplitterMouseMove(ByVal X As Single, ByVal Y As Single)
    Dim hDC As Long
    Dim tP As POINTAPI
    Dim hWndClient As Long
    
    If (bDraw) Then
        ' Draw two rectangles in the screen DC to cause splitting:
        ' First get the Desktop DC:
        hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        ' Set the draw mode to XOR:
        SetROP2 hDC, R2_NOTXORPEN
        ' Draw over and erase the old rectangle
        ' (if this is the first time, all the coords will be 0 and nothing will get drawn):
        Rectangle hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
        ' It is simpler to use the mouse cursor position than try to translate
        ' X,Y to screen coordinates!
        GetCursorPos tP
        ' Determine where to draw the splitter:
        If (m_Orientation = spHorizontal) Then
            rcNew.Left = rcSplit.Left + m_lBorder(espbLeft)
            rcNew.Right = rcSplit.Right - m_lBorder(espbRight)
            If (tP.Y >= rcWindow.Top + m_lBorder(espbTop)) And (tP.Y < rcWindow.Bottom - m_lBorder(espbBottom)) Then
                rcNew.Top = tP.Y - 2
                rcNew.Bottom = tP.Y + 2
            Else
                If (tP.Y < rcWindow.Top + m_lBorder(espbTop)) Then
                    rcNew.Top = rcWindow.Top + m_lBorder(espbTop) - 2
                    rcNew.Bottom = rcNew.Top + 5
                Else
                    rcNew.Top = rcWindow.Bottom - m_lBorder(espbBottom) - 2
                    rcNew.Bottom = rcNew.Top + 5
                End If
            End If
        Else
            rcNew.Top = rcSplit.Top + m_lBorder(espbTop)
            rcNew.Bottom = rcSplit.Bottom - m_lBorder(espbBottom)
            If (tP.X >= rcWindow.Left + m_lBorder(espbLeft)) And (tP.X <= rcWindow.Right - m_lBorder(espbRight)) Then
                rcNew.Left = tP.X - 2
                rcNew.Right = tP.X + 2
            Else
                If (tP.X < rcWindow.Left + m_lBorder(espbLeft)) Then
                    rcNew.Left = rcWindow.Left + m_lBorder(espbLeft) - 2
                    rcNew.Right = rcNew.Left + 5
                Else
                    rcNew.Left = rcWindow.Right - m_lBorder(espbRight) - 2
                    rcNew.Right = rcNew.Left + 5
                End If
            End If
        End If
        ' Draw the new rectangle
        Rectangle hDC, rcNew.Left, rcNew.Top, rcNew.Right, rcNew.Bottom
        ' Store this position so we can erase it next time:
        LSet rcCurrent = rcNew
        ' Free the reference to the Desktop DC we got (make sure you do this!)
        DeleteDC hDC
    End If
End Sub

Private Sub SplitterMouseUp(ByVal X As Single, ByVal Y As Single)
    Dim hDC As Long
    Dim tP As POINTAPI
    Dim hWndClient As Long
    Dim Offset
    
    ' Don't leave orphaned rectangle on desktop; erase last rectangle.
    If (bDraw) Then
        bDraw = False
        ' Release mouse capture:
        ReleaseCapture
        ' Release the cursor clipping region (must do this!):
        ClipCursorClear 0&
        ' Get the Desktop DC:
        hDC = CreateDCAsNull("DISPLAY", 0, 0, 0)
        ' Set to XOR drawing mode:
        SetROP2 hDC, R2_NOTXORPEN
        ' Erase the last rectangle:
        Rectangle hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
        ' Clear up the desktop DC:
        DeleteDC hDC
        ' Here we ensure the splitter is within bounds before releasing:
        GetCursorPos tP
        If (tP.X < rcWindow.Left + m_lBorder(espbLeft)) Then
            tP.X = rcWindow.Left + m_lBorder(espbLeft)
        End If
        If (tP.X > rcWindow.Right - m_lBorder(espbRight)) Then
            tP.X = rcWindow.Right - m_lBorder(espbRight)
        End If
        If (tP.Y < rcWindow.Top + m_lBorder(espbTop)) Then
            tP.Y = rcWindow.Top + m_lBorder(espbTop)
        End If
        If (tP.Y > rcWindow.Bottom - m_lBorder(espbBottom)) Then
            tP.Y = rcWindow.Bottom - m_lBorder(espbBottom)
        End If
        ScreenToClient m_hWnd, tP
        ' Move the splitter to the validated final position:
        On Error Resume Next
        LockWindowUpdate Extender.Container.hWnd
        If (m_Orientation = spHorizontal) Then
            Offset = Extender.Height / 2
            SizeControls (tP.Y * Screen.TwipsPerPixelY), Offset
            Extender.Top = (tP.Y * Screen.TwipsPerPixelY) - Offset
        Else
            Offset = Extender.Width / 2
            SizeControls tP.X * Screen.TwipsPerPixelX, Offset
            Extender.Left = (tP.X * Screen.TwipsPerPixelX) - Offset
        End If
        LockWindowUpdate 0
    End If
End Sub

Private Sub UserControl_Paint()
    If m_bDirty Then Exit Sub
    
    Dim rc As RECT
    Dim bdrStyle As Long
    Dim bdrSides As Long
    Dim Offset
    
    ' set this to true to avoid redrawing
    m_bDirty = True
    ' update thickness
    Thickness = m_Thickness
    ' the split bar must be on top of other controls, so
    ' cal zorder method
    Extender.ZOrder
    ' all sides must be updated
    bdrSides = BF_RECT
    ' update border styles
    If m_BorderStyle = bdrFlat Then bdrSides = bdrSides Or BF_FLAT
    If m_BorderStyle = bdrMono Then bdrSides = bdrSides Or BF_MONO
    If m_BorderStyle = bdrSoft Then bdrSides = bdrSides Or BF_SOFT
    Select Case m_BorderStyle
        Case bdrRaisedOuter: bdrStyle = BDR_RAISEDOUTER
        Case bdrRaisedInner: bdrStyle = BDR_RAISEDINNER
        Case bdrRaised: bdrStyle = EDGE_RAISED
        Case bdrSunkenOuter: bdrStyle = BDR_SUNKENOUTER
        Case bdrSunkenInner: bdrStyle = BDR_SUNKENINNER
        Case bdrSunken: bdrStyle = EDGE_SUNKEN
        Case bdrEtched: bdrStyle = EDGE_ETCHED
        Case bdrBump: bdrStyle = EDGE_BUMP
        Case bdrFlat: bdrStyle = BDR_SUNKEN
        Case bdrMono: bdrStyle = BDR_SUNKEN
        Case bdrSoft: bdrStyle = BDR_RAISED
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
    ' if we are in run-mode update controls
    If Ambient.UserMode = True Then
        If m_Orientation = spVertical Then
            Offset = (Extender.Width / 2)
            SizeControls Extender.Left + Offset, Offset
        Else
            Offset = (Extender.Height / 2)
            SizeControls Extender.Top + Offset, Offset
        End If
        ' raise this event in order to
        ' call user-defined resizing operations
        '        RaiseEvent EndMoving
    End If
    ' clear flag
    m_bDirty = False
End Sub

Public Sub Refresh()
    UserControl_Paint
End Sub

Private Sub UserControl_Resize()
    Dim Offset
    
    If m_bDirty Then Exit Sub
    On Error Resume Next
    If Extender.Visible Then
        If m_Orientation = spVertical Then
            Offset = (UserControl.Width / 2)
            SizeControls Extender.Left + Offset, Offset
        Else
            Offset = (UserControl.Height / 2)
            SizeControls Extender.Top + Offset, Offset
        End If
        ' refresh all splitters found in the container
        RefreshAll
    End If
End Sub

Private Sub UserControl_Terminate()
    Set m_Controls = Nothing
End Sub

' ******************************************************************************
' Routine       : AddControl
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 04/12/00 23:36:30
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : add a control reference position to the
'                 splitter bar object postion control
'                 This is the main engine that makes the splitbar
'
' ******************************************************************************
Public Sub AddControl(Ctl As Object, Position As spControlPosition)
    If m_Orientation = spHorizontal Then
        If Position = spLeft Or Position = spRight Then
            Debug.Assert "Invalid control position for " & UserControl.Name
        End If
    Else
        If Position = spTop Or Position = spBottom Then
            Debug.Assert "Invalid control position for " & UserControl.Name
        End If
    End If
    ' add control reference to the control collection
    ' note:  only the object handle is stored in the collection
    ' not the object itself, this avoid many hazards and crashes
    m_Controls.Add PtrFromObject(Ctl), Position
End Sub

' ******************************************************************************
' Routine       : SizeControls
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 13/12/00 19:40:56
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Based on each control position this routine
'                 will try to align each control to its position
' ******************************************************************************
Private Sub SizeControls(ByVal Position, ByVal Offset)
    Dim Ctl As Control
    Dim elem As TControl
    Dim Obj As Object
    Dim rc As RECT
    Dim tx As Integer
    Dim ty As Integer
    
    tx = Screen.TwipsPerPixelX
    ty = Screen.TwipsPerPixelY
    
    On Error Resume Next
    GetClientArea rc
    ' check splitter orientation
    For Each elem In m_Controls
        Set Obj = ObjectFromPtr(elem.Handle)
        If m_Orientation = spVertical Then
            If elem.Position = spLeft Or elem.Position = spClientLeft Then
                If elem.Position = spClientLeft Then
                    Obj.Left = rc.Left
                    Obj.Top = rc.Top
                    Obj.Height = Extender.Height 'rc.Bottom
                End If
                Obj.Width = (Position - (Obj.Left + Offset)) - tx
            ElseIf elem.Position = spRight Or elem.Position = spClientRight Then
                Obj.Left = (Position + Offset) + tx
                If elem.Position = spClientRight Then
                    Obj.Top = Extender.Top
                    Obj.Height = Extender.Height
                    Obj.Width = (rc.Right - Obj.Left) - (2 * tx)
                End If
            End If
        Else
            If elem.Position = spTop Or elem.Position = spClientTop Then
                If elem.Position = spClientTop Then
                    Obj.Top = rc.Top + ty
                    Obj.Left = Extender.Left + tx
                    Obj.Width = Extender.Width - (2 * tx)
                    Obj.Height = Extender.Top - (Obj.Top + ty)
                Else
                    Obj.Height = (Position - Obj.Top) - (2 * ty)
                End If
            ElseIf elem.Position = spBottom Or elem.Position = spClientBottom Then
                Obj.Top = Extender.Top + Extender.Height + ty
                If elem.Position = spClientBottom Then
                    Obj.Left = Extender.Left + tx
                    Obj.Width = Extender.Width - (2 * tx)
                    Obj.Height = (rc.Bottom - Obj.Top) - (2 * ty)
                Else
                    'Obj.Height = (rc.Bottom - Obj.Top) - (2 * ty)
                End If
            End If
        End If
'        If TypeName(Obj) = "TSplitter" Then
'            Obj.Refresh
'        End If
    Next
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project      :       CSSplitter
' Procedure    :       GetClientArea
' Description  :       Get the available client area for splitting range
'                      based on aligned controls in the container
'
' Created by   :       Project Administrator
' Machine      :       ZEUS
' Date-Time    :       17/4/2004-12:09:36
'
' Parameters   :       rc (RECT)
' Return Values:
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub GetClientArea(rc As RECT)
    Dim Ctl As Control
    Dim MaxT As Long
    Dim MaxB As Long
    Dim MaxL As Long
    Dim MaxR As Long
    
    ' loop every control in the conatainer
    For Each Ctl In Extender.Parent.Controls
        ' if this is not a splitter bar...
        If TypeName(Ctl) <> "TSplitter" Then
            ' if this control has align property...
            If IsAligned(Ctl) Then
                If Ctl.Container.Name = Extender.Container.Name Then
                    ' will wil get the client rect in twips
                    ' based on align property of the control
                    If Ctl.Align = vbAlignTop Then
                        MaxT = Max(MaxT, Ctl.Top + Ctl.Height)
                    ElseIf Ctl.Align = vbAlignBottom Then
                        MaxB = Max(MaxB, Ctl.Height)
                    ElseIf Ctl.Align = vbAlignLeft Then
                        MaxL = Max(MaxL, Ctl.Width)
                    ElseIf Ctl.Align = vbAlignRight Then
                        MaxR = Max(MaxR, Ctl.Left)
                    End If
                End If
            End If
        End If
    Next
    ' set rect based on control position information
    With rc
        .Top = MaxT
        .Left = MaxL
        .Right = Extender.Container.ScaleWidth - (MaxR + MaxL)
        .Bottom = Extender.Container.ScaleHeight - (MaxB + MaxT)
    End With
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project      :       CSSplitter
' Procedure    :       RefreshAll
' Description  :       Refresh all splitters found in the owner form or container
' Created by   :       Project Administrator
' Machine      :       ZEUS
' Date-Time    :       17/4/2004-12:11:03
'
' Parameters   :
' Return Values:
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub RefreshAll()
    Dim Ctl As Control
    
    ' loop every container control
    For Each Ctl In Extender.Parent.Controls
        ' refresh all splitters found
        If TypeName(Ctl) = "TSplitter" Then
            If Ctl.Name <> Extender.Name Then
                Ctl.Refresh
            End If
        End If
    Next
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Project      :       CSSplitter
' Procedure    :       GetDragArea
' Description  :       Get the available drag area based on attached control
'                      dimensions
' Created by   :       Project Administrator
' Machine      :       ZEUS
' Date-Time    :       17/4/2004-12:11:27
'
' Parameters   :       rc (RECT)
' Return Values:
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub GetDragArea(rc As RECT)
    Dim CT As TControl
    Dim Ctl As Control
    Dim hT As Long
    Dim hB As Long
    Dim hL As Long
    Dim hR As Long
    
    ' note: m_Controls holds form control references and
    ' position concerning the split bar
    hR = 999999
    hB = 999999
    For Each CT In m_Controls
        Set Ctl = ObjectFromPtr(CT.Handle)
        If m_Orientation = spVertical Then
            If CT.Position = spLeft Or CT.Position = spClientLeft Then
                hL = Max(hL, Ctl.Left + m_MinWidth)
            ElseIf CT.Position = spRight Or CT.Position = spClientRight Then
                hR = Min(hR, Ctl.Left + (Ctl.Width - m_MinWidth))
            End If
        Else
            If CT.Position = spTop Or CT.Position = spClientTop Then
                hT = Max(hT, Ctl.Top + m_MinHeight)
            ElseIf CT.Position = spBottom Or CT.Position = spClientBottom Then
                hB = Min(hB, Ctl.Top + (Ctl.Height - m_MinHeight))
            End If
        End If
    Next
    ' set rect in pixelsa
    ' based on the himum information
    ' found for each control attached
    With rc
        If m_Orientation = spVertical Then
            .Left = hL / Screen.TwipsPerPixelX
            .Right = (hR / Screen.TwipsPerPixelX) - .Left
        Else
            .Top = hT / Screen.TwipsPerPixelY
            .Bottom = (hB / Screen.TwipsPerPixelY) - .Top
        End If
    End With
End Sub
'-- end code
