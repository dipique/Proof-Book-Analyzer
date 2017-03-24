VERSION 5.00
Begin VB.Form frmPrint 
   Caption         =   "Print Documents"
   ClientHeight    =   2595
   ClientLeft      =   3150
   ClientTop       =   2070
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   11040
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   495
      Left            =   9720
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer timProcess 
      Interval        =   3500
      Left            =   1800
      Top             =   1560
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox lstQueue 
      Height          =   1620
      ItemData        =   "Print.frx":0000
      Left            =   120
      List            =   "Print.frx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   10815
   End
   Begin VB.Label lblCurrent 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   960
      TabIndex        =   5
      Top             =   45
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Job:"
      Height          =   195
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   855
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oWord As Object
Dim oTemp As Object

Const FormWidth As Integer = 11160
Const FormHeight As Integer = 3105

Private Sub cmdDelete_Click()

lstQueue.RemoveItem (lstQueue.ListIndex)

End Sub

Private Sub cmdPause_Click()

If cmdPause.Caption = "&Pause" Then

cmdPause.Caption = "&Resume"
timProcess.Enabled = False

Else

cmdPause.Caption = "&Pause"
timProcess.Enabled = True

End If

End Sub

Private Sub cmdStop_Click()

If lstQueue.List(0) <> "" Then

Dim vbMessage As String
vbMessage = "Are you sure you want to quit all print jobs?"
If MsgBox(vbMessage, vbYesNo, "Cancel Print Jobs") = vbNo Then Exit Sub

End If

  lstQueue.Clear
  Call RetForm(Me, False)
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyDelete Then cmdDelete_Click

End Sub

Private Sub Form_Load()

  frmPrint.Show
  timProcess.Interval = 2000
  
End Sub

Private Sub Form_Resize()

frmPrint.Height = FormHeight
frmPrint.Width = FormWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set oWord = Nothing
  Set oTemp = Nothing
  
  Call RetForm(Me, False)
  
End Sub

Private Sub lstQueue_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyDelete Then cmdDelete_Click

End Sub

Private Sub timProcess_Timer()

  Dim CopyNumber As Integer
  Dim CopyString As String
  
  On Error GoTo Continue
  
  timProcess.Interval = 4500
  
  Set oTemp = Nothing
  oWord.Quit
  Set oWord = Nothing
    
Continue:
  
  If lstQueue.List(0) <> "" Then
    
    lblCurrent = lstQueue.List(0)
    lstQueue.RemoveItem (0)
    
    Set oWord = Nothing
    
    If Mid(lblCurrent.Caption, 2, 2) = ":" Then
      
      CopyString = Mid(lblCurrent.Caption, 3, Len(lblCurrent.Caption))
    
      CopyNumber = Val(Mid(lblCurrent.Caption, 1, 1))
      
    Else
      
      CopyString = Mid(lblCurrent.Caption, 3, Len(lblCurrent.Caption))

      CopyNumber = Val(Mid(lblCurrent.Caption, 1, 2))
      
    End If
    
    Set oWord = CreateObject("Word.Application")
    Set oTemp = oWord.Documents.Open(FileName:=CopyString, ReadOnly:=True)
        
    With oWord
 '     .Visible = True
      .ActiveDocument.PrintOut Copies:=1
    End With
    
  End If
  
End Sub

