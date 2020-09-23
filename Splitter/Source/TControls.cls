VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"TControl"
Attribute VB_Ext_KEY = "Member0" ,"TControl"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' ******************************************************************************
' Class         : TControls
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 13/12/00 19:00:31
' Credits       :
' Modifications :
' Description   : private class that holds
'                 splitter control references
' ******************************************************************************

Private mCol As Collection

Public Function Add(Handle As Long, Position As spControlPosition, Optional sKey As String) As TControl
    Dim objNewMember As TControl
    
    ' create a new object
    Set objNewMember = New TControl
    ' set the properties passed into the method
    objNewMember.Handle = Handle
    objNewMember.Position = Position
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    ' return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As TControl
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

