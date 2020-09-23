VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Splitter Demo"
   ClientHeight    =   2355
   ClientLeft      =   6645
   ClientTop       =   3555
   ClientWidth     =   5055
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1622.307
   ScaleMode       =   0  'User
   ScaleWidth      =   4746.906
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSep 
      Height          =   75
      Left            =   0
      TabIndex        =   3
      Top             =   1350
      Width           =   5835
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   0
      Top             =   1860
      Width           =   1380
   End
   Begin VB.Label Label2 
      Caption         =   "marclei@bannerbox.net"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   1710
      Width           =   3165
   End
   Begin VB.Label Label1 
      Caption         =   "Author: Marclei V Silva"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   1470
      Width           =   4725
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   390
      Picture         =   "frmAbout.frx":000C
      Top             =   390
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   150
      Picture         =   "frmAbout.frx":0092
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   2760
      TabIndex        =   2
      Top             =   810
      Width           =   2145
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   870
      TabIndex        =   1
      Top             =   60
      Width           =   4155
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   645
      Left            =   60
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    Me.Icon = frmMain.Icon
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

