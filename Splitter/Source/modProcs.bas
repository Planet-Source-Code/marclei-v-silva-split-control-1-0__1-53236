Attribute VB_Name = "modProcs"
Option Explicit

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       CSSplitter
' Procedure  :       Max
' Description:       Returns the higher value between two numbers given
' Created by :       Project Administrator
' Machine    :       PERSEU
' Date-Time  :       17/4/2004-16:48:47
'
' Parameters :       v1 (Variant)
'                    v2 (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Max(v1, v2)
    If v1 > v2 Then
        Max = v1
    Else
        Max = v2
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       CSSplitter
' Procedure  :       Max
' Description:       Returns the higher value between two numbers given
' Created by :       Project Administrator
' Machine    :       PERSEU
' Date-Time  :       17/4/2004-16:48:47
'
' Parameters :       v1 (Variant)
'                    v2 (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Min(v1, v2)
    If v1 > v2 Then
        Min = v2
    Else
        Min = v1
    End If
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Project    :       CSSplitter
' Procedure  :       IsAligned
' Description:       Returns true if the control has "align" property
' Created by :       Project Administrator
' Machine    :       PERSEU
' Date-Time  :       17/4/2004-16:49:16
'
' Parameters :       Ctl (Control)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function IsAligned(Ctl As Control) As Boolean
    Dim Dummy As Boolean
    
    On Error Resume Next
    Dummy = (Ctl.Align = vbAlignNone)
    IsAligned = (Err.Number = 0)
End Function


