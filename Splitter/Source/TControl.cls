VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' ******************************************************************************
' Class         : TControl
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 13/12/0019:00:31
' Credits       :
' Modifications :
' Description   : private class to hold splitter bar control reference
' ******************************************************************************

Private m_Handle As Long
Private m_Position As spControlPosition

Public Property Let Position(ByVal vData As spControlPosition)
    m_Position = vData
End Property

Public Property Get Position() As spControlPosition
    Position = m_Position
End Property

Public Property Let Handle(ByVal vData As Long)
    m_Handle = vData
End Property

Public Property Get Handle() As Long
    Handle = m_Handle
End Property



