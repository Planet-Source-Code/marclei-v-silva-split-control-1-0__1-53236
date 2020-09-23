Attribute VB_Name = "modApi32"
Option Explicit
' ******************************************************************************
' Module        : modApi32
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 13/12/00 19:00:31
' Credits       :
' Modifications :
' Description   : Module that declares and defines several Api
'                 variables, routines and functions
' ******************************************************************************

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const VK_LBUTTON = &H1
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Public Const R2_BLACK = 1       '   0
Public Const R2_COPYPEN = 13    '  P
Public Const R2_LAST = 16
Public Const R2_MASKNOTPEN = 3  '  DPna
Public Const R2_MASKPEN = 9     '  DPa
Public Const R2_MASKPENNOT = 5  '  PDna
Public Const R2_MERGENOTPEN = 12        '  DPno
Public Const R2_MERGEPEN = 15   '  DPo
Public Const R2_MERGEPENNOT = 14        '  PDno
Public Const R2_NOP = 11        '  D
Public Const R2_NOT = 6 '  Dn
Public Const R2_NOTCOPYPEN = 4  '  PN
Public Const R2_NOTMASKPEN = 8  '  DPan
Public Const R2_NOTMERGEPEN = 2 '  DPon
Public Const R2_NOTXORPEN = 10  '  DPxn
Public Const R2_WHITE = 16      '   1
Public Const R2_XORPEN = 7      '  DPx

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Public Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXBORDER = 5
Public Const SM_CYBORDER = 6
Public Const SM_CYCAPTION = 4
Public Const SM_CYMENU = 15

' These constants define the style of border to draw.
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_FLAT = &H4000
Public Const BF_MONO = &H8000
Public Const BF_SOFT = &H1000      ' For softer buttons

Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

' These constants define which sides to draw.
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTBOTTOM = 15

'<CSCM>
'--------------------------------------------------------------------------------
' Project      :       CSSplitter
' Procedure    :       ObjectFromPtr
' Description  :       Returns an object from a memory pointer
' Created by   :       Project Administrator
' Machine      :       ZEUS
' Date-Time    :       18/4/2004-10:30:01
'
' Parameters   :       lPtr (Long)
' Return Values:
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oThis As Object

    ' Turn the pointer into an illegal, uncounted interface
    CopyMemory oThis, lPtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set ObjectFromPtr = oThis
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory oThis, 0&, 4
    ' OK, hit the End button if you must--you'll probably still crash,
    ' but this will be your code rather than the uncounted reference!
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project      :       CSSplitter
' Procedure    :       PtrFromObject
' Description  :       Retrieve a memory pointer from an object
' Created by   :       Project Administrator
' Machine      :       ZEUS
' Date-Time    :       18/4/2004-10:30:26
'
' Parameters   :       oThis (Object)
' Return Values:
'--------------------------------------------------------------------------------
'</CSCM>
Public Function PtrFromObject(ByRef oThis As Object) As Long
    ' Return the pointer to this object:
    PtrFromObject = ObjPtr(oThis)
End Function
'-- end code
