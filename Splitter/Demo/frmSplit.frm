VERSION 5.00
Object = "*\A..\Source\CSSplitter.vbp"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSplit 
   Caption         =   "Splittable Child Window"
   ClientHeight    =   5490
   ClientLeft      =   4065
   ClientTop       =   2175
   ClientWidth     =   9060
   Icon            =   "frmSplit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   9060
   Begin CSSplitter.TSplitter VSplitter3 
      Height          =   1755
      Left            =   5220
      Top             =   3600
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   3096
      BorderStyle     =   8
      MousePointer    =   9
   End
   Begin CSSplitter.TSplitter VSplitter2 
      Height          =   3240
      Left            =   4470
      ToolTipText     =   "VSplitter2"
      Top             =   30
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   5715
      BorderStyle     =   7
      MousePointer    =   9
   End
   Begin CSSplitter.TSplitter HSplitter1 
      Height          =   45
      Left            =   30
      ToolTipText     =   "HSplitter"
      Top             =   3420
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   79
      Orientation     =   2
      Thickness       =   50
      MousePointer    =   7
   End
   Begin CSSplitter.TSplitter HSplitter2 
      Height          =   105
      Left            =   2250
      ToolTipText     =   "HSplitter"
      Top             =   3420
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   185
      BorderStyle     =   3
      Orientation     =   2
      Thickness       =   100
      MousePointer    =   7
   End
   Begin CSSplitter.TSplitter VSplitter1 
      Height          =   5415
      Left            =   2040
      Top             =   30
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   9551
      BorderStyle     =   11
      Thickness       =   150
      MousePointer    =   9
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   30
      TabIndex        =   3
      Top             =   3600
      Width           =   1905
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   4680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmSplit.frx":014A
      Top             =   30
      Width           =   3552
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   2310
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmSplit.frx":0150
      Top             =   30
      Width           =   2025
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1665
      Left            =   2310
      TabIndex        =   4
      Top             =   3600
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2937
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmSplit.frx":0156
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   1665
      Left            =   5520
      TabIndex        =   5
      Top             =   3600
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2937
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmSplit.frx":01E1
   End
End
Attribute VB_Name = "frmSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Modified As Boolean

Private Sub Form_Load()
    Dim sText As String
    Dim sFile As String
    
    sFile = App.Path & "\readme.rtf"
    RichTextBox1.LoadFile sFile
    RichTextBox1.Tag = sFile
    
    sFile = App.Path & "\revisions.rtf"
    RichTextBox2.LoadFile sFile
    RichTextBox2.Tag = sFile
    
    ' reset modified flag
    m_Modified = False
    
    sFile = App.Path & "\frmMain.frm"
    If (GetFileText(sFile, sText)) Then
        Text1.Text = sText
    Else
        Text1.Text = "Source code to file '" & sFile & "' could not be found."
    End If
    
    sFile = App.Path & "\frmSplit.frm"
    If (GetFileText(sFile, sText)) Then
        Text2.Text = sText
    Else
        Text2.Text = "Source code to file '" & sFile & "' could not be found."
    End If
    
    Dir1.Path = App.Path
    
    ' The only thing you do is to attach controls to the splitters
    ' The split engine will take control of the limits it can resize
    ' and do the job for you.
    ' you can also control sizing through the events StartMoving() and EndMoving()
    VSplitter1.AddControl Text1, spRight
    VSplitter1.AddControl HSplitter2, spRight
    VSplitter1.AddControl HSplitter1, spLeft
    VSplitter1.AddControl Dir1, spLeft
    VSplitter1.AddControl File1, spLeft
    VSplitter1.AddControl RichTextBox1, spRight
    
    VSplitter2.AddControl Text1, spLeft
    VSplitter2.AddControl Text2, spClientRight
    
    VSplitter3.AddControl RichTextBox1, spLeft
    VSplitter3.AddControl RichTextBox2, spClientRight
    
    HSplitter1.AddControl Dir1, spTop
    HSplitter1.AddControl File1, spClientBottom
    
    HSplitter2.AddControl Text1, spTop
    HSplitter2.AddControl Text2, spTop
    HSplitter2.AddControl VSplitter2, spTop
    HSplitter2.AddControl RichTextBox1, spBottom
    HSplitter2.AddControl RichTextBox2, spBottom
    HSplitter2.AddControl VSplitter3, spBottom
    
    Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    VSplitter1.Height = Me.ScaleHeight - VSplitter1.Top
    HSplitter2.Width = Me.ScaleWidth - HSplitter2.Left
    HSplitter2_EndMoving
End Sub

Private Function GetFileText(ByVal sFile As String, ByRef sText As String) As Boolean
    Dim iFIle As Integer
    Dim lLen As Long
    
    iFIle = FreeFile
    On Error Resume Next
    Open sFile For Binary Access Read As #iFIle
    If (Err.Number = 0) Then
        lLen = LOF(iFIle)
        sText = String$(lLen, 0)
        Get #iFIle, , sText
        If (Err.Number = 0) Then
            GetFileText = True
        End If
        Close #iFIle
    End If

End Function

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub File1_DblClick()
    Dim sFile As String
    Dim sText As String
    
    ' open file specified
    sFile = File1.Path & "\" & File1.FileName
    If (GetFileText(sFile, sText)) Then
        Text1.Text = sText
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' check wether text was modified
    If m_Modified Then
        Dim nRet As VbMsgBoxResult
        
        nRet = MsgBox("File was modified. Do you want to save changes?", vbQuestion + vbYesNoCancel)
        Select Case nRet
            Case vbYes: Save
            Case vbCancel: Cancel = True
            Case vbNo: ' nothing
        End Select
    End If
End Sub

Private Sub HSplitter2_EndMoving()
    ' use end moving event to set other
    ' split bar parameters
    Dim lHeight As Long
    
    lHeight = ScaleHeight - (HSplitter2.Top + HSplitter2.Height)
    VSplitter3.Height = lHeight
    RichTextBox1.Height = lHeight
    'VSplitter3.Refresh
End Sub

Private Sub RichTextBox1_Change()
    ' text was modified
    m_Modified = True
End Sub

Private Sub RichTextBox2_Change()
    ' text was modified
    m_Modified = True
End Sub

Private Sub Save()
    ' save files
    RichTextBox1.SaveFile RichTextBox1.Tag
    RichTextBox2.SaveFile RichTextBox2.Tag
    m_Modified = False
End Sub

Private Sub VSplitter1_EndMoving()
    HSplitter2.Width = ScaleWidth - (VSplitter1.Left + VSplitter1.Width)
End Sub
'-- end code
