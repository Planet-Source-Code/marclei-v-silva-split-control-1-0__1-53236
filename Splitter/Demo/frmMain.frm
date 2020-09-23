VERSION 5.00
Object = "*\A..\Source\CSSplitter.vbp"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Splitter Demo Application"
   ClientHeight    =   6480
   ClientLeft      =   3780
   ClientTop       =   2865
   ClientWidth     =   8685
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin CSSplitter.TPanel TPanel2 
      Align           =   2  'Align Bottom
      Height          =   1395
      Left            =   0
      TabIndex        =   7
      Top             =   5085
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   2461
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TSplit && TPanel Control version 1.0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   270
         Width           =   6285
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TSplit && TPanel Control version 1.0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   465
         Index           =   0
         Left            =   270
         TabIndex        =   9
         Top             =   300
         Width           =   6285
      End
   End
   Begin CSSplitter.TPanel TPanel1 
      Align           =   3  'Align Left
      Height          =   5085
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   8969
      BackColor       =   -2147483636
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   7
      BackColor       =   -2147483636
      Begin VB.ComboBox cboBack 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   780
         Width           =   1515
      End
      Begin VB.ComboBox cboBorder 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1350
         Width           =   1515
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1980
         Width           =   1515
      End
      Begin VB.Label lblDemo 
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom panel control"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   8
         Top             =   150
         Width           =   2355
      End
      Begin VB.Label lblDemo 
         BackStyle       =   0  'Transparent
         Caption         =   "BackColor"
         ForeColor       =   &H80000014&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   510
         Width           =   1515
      End
      Begin VB.Label lblDemo 
         BackStyle       =   0  'Transparent
         Caption         =   "BorderStyle:"
         ForeColor       =   &H80000014&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1110
         Width           =   1515
      End
      Begin VB.Label lblDemo 
         BackStyle       =   0  'Transparent
         Caption         =   "BackStyle:"
         ForeColor       =   &H80000014&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1710
         Width           =   1515
      End
   End
   Begin VB.Menu mnuF0MAIN 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New..."
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelpTOP 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About..."
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboBack_Click()
    TPanel2.BackColor = cboBack.ItemData(cboBack.ListIndex)
End Sub

Private Sub cboBorder_Click()
    TPanel2.BorderStyle = cboBorder.ItemData(cboBorder.ListIndex)
End Sub

Private Sub cboStyle_Click()
    TPanel2.BorderStyle = cboStyle.ItemData(cboStyle.ListIndex)
End Sub

Private Sub MDIForm_Load()
    
    AddItem cboBack, "Red", vbRed
    AddItem cboBack, "Blue", vbBlue
    AddItem cboBack, "Green", vbGreen
    AddItem cboBack, "ButtonFace", vbButtonFace
    AddItem cboBack, "AppWorkSpace", vbApplicationWorkspace

    AddItem cboStyle, "PanelTransparent", 0
    AddItem cboStyle, "PanelOpaque", 1

    AddItem cboBorder, "PanelBorderNone", 0
    AddItem cboBorder, "PanelBorderRaisedOuter", 1
    AddItem cboBorder, "PanelBorderRaisedInner", 2
    AddItem cboBorder, "PanelBorderRaised", 3
    AddItem cboBorder, "PanelBorderSunkenOuter", 4
    AddItem cboBorder, "PanelBorderSunkenInner", 5
    AddItem cboBorder, "PanelBorderSunken", 6
    AddItem cboBorder, "PanelBorderEtched", 7
    AddItem cboBorder, "PanelBorderBump", 8
    AddItem cboBorder, "PanelBorderMono", 9
    AddItem cboBorder, "PanelBorderFlat", 10
    AddItem cboBorder, "PanelBorderSoft", 11

    frmSplit.Show
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim fS As New frmSplit
            fS.Show
        Case 2
            Unload Me
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
    Dim f As frmAbout
    
    Set f = New frmAbout
    f.Show vbModal, Me
    Unload f
    Set f = Nothing
End Sub

Private Sub TPanel1_Resize()
    Dim I As Integer
    
    On Error Resume Next
    With TPanel1
        For I = 0 To 2
            lblDemo(I).Width = .ScaleWidth - lblDemo(I).Left * 2
            cboBack.Width = .ScaleWidth - lblDemo(I).Left * 2
            cboBorder.Width = .ScaleWidth - lblDemo(I).Left * 2
            cboStyle.Width = .ScaleWidth - lblDemo(I).Left * 2
        Next I
    End With
End Sub

Public Function AddItem(rList As Variant, Text As Variant, Optional Value As Variant, Optional Selected As Variant, Optional Index As Variant) As Integer
    Dim sText As String
    Dim lngValue As Long
    Dim bSelected As Boolean
    
    If Not IsNull(Text) Then
        sText = Text
    End If
    If Not IsMissing(Value) Then
        If Not IsNull(Value) Then
            lngValue = Val(Value)
        End If
    End If
    If IsMissing(Selected) Then
        bSelected = False
    Else
        bSelected = CBool(Selected)
    End If
    If IsMissing(Index) Then
        rList.AddItem sText
    Else
        rList.AddItem sText, Index
    End If
    rList.ItemData(rList.NewIndex) = lngValue
    On Error Resume Next
    rList.Selected = bSelected
    If bSelected Then
        rList.ListIndex = rList.NewIndex
        rList.Selected(rList.NewIndex) = True
    End If
    AddItem = rList.NewIndex
End Function
'-- end code
