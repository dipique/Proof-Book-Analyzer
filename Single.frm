VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSingle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Single Page"
   ClientHeight    =   3825
   ClientLeft      =   3495
   ClientTop       =   1515
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4095
   Begin VB.Timer timRestrict 
      Interval        =   50
      Left            =   1440
      Top             =   1680
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "&Single"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "Current mode.  Press to toggle between single page and section printing."
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      ToolTipText     =   "Return to main form."
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin MSComctlLib.TreeView treSingle 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5318
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WA As Object
Dim Doc As Object

Private Sub cmdExit_Click()

  Form_Unload (1)
  
End Sub

Private Sub cmdOpen_Click()

  Dim strFN As String
  strFN = Locations(treSingle.SelectedItem.Index)
  
  On Error GoTo lEnd
  
  Dim oWord As Object
  Set oWord = CreateObject("Word.Application")
  Dim oTemp
  Set oTemp = oWord.Documents.Open(FileName:=strFN)
  oWord.Visible = True
  
lEnd:
  
End Sub

Private Sub cmdPrint_Click()

  Dim n As Integer
  Dim LocStr As String
  Dim CopInt As Integer
  Dim CMult As Integer
  
  frmSingle.Visible = False
  
  If cmdToggle.Caption = "&Section" Then
    CopInt = Int(Val(InputBox("Copies", "How many copies?", "1")))
    If CopInt < 1 Or CopInt > 99 Then
      MsgBox "Please enter number over zero."
      Exit Sub
    End If
    
    'Each Copy Begins Here
    frmPrint.Show
    
    For CMult = 1 To CopInt
      For n = 1 To treSingle.SelectedItem.Children          ' sends to print queue
        With treSingle.SelectedItem
          LocStr = Locations(.Index + .Children - n + 1)
          If Mid(LocStr, 1, 4) <> "Note" Then Call SubmitDoc(LocStr, CopInt)
        End With
      Next n
    Next CMult
  Else                                                      ' for &Section statement
    
    CopInt = Int(Val(InputBox("Copies", "How many copies?", "1")))
    
    If CopInt < 1 Or CopInt > 99 Then
      MsgBox "Please enter number over zero."
      Exit Sub
    End If                                                  ' End Finding Copy numbers
    
    LocStr = Locations(treSingle.SelectedItem.Index)
    
    If Mid$(LocStr, 1, 4) <> "Note" Then Call SubmitDoc(LocStr, CopInt)
    
  End If
  
  Call RetForm(frmPrint, True)
  
End Sub

Private Sub cmdToggle_Click()

frmSingle.Caption = "Thinking..."
cmdToggle.Enabled = False
  
  Form_Load
  
    If cmdToggle.Caption = "&Single" Then
    
    cmdToggle.Caption = "&Section"
    frmSingle.Caption = "Print Section"
    
  Else
    
    cmdToggle.Caption = "&Single"
    frmSingle.Caption = "Print Single Page"
    
  End If
  
  cmdToggle.Enabled = True
  
End Sub

Private Sub Form_Load()

  timRestrict.Interval = 0
  
  'Create All Nodes
  Dim n As Integer
  Dim PreviousNode As String
  
  treSingle.Nodes.Clear
  
  For n = 1 To UBound(PageNames)
    
    If PageChild(n) = True Then                             ' if line is subcategory
      
      treSingle.Nodes.Add PreviousNode, tvwChild, PageNames(n), PageNames(n)
      
    Else                                                    ' if line is root node
      
      treSingle.Nodes.Add , , PageNames(n), PageNames(n)
      PreviousNode = PageNames(n)
      
    End If
    
  Next n

  Dim I As Integer
  For n = 1 To treSingle.Nodes.Count
    'Make all With Children Bold
    For I = 1 To treSingle.Nodes(n).Children
      If I = 1 Then treSingle.Nodes(n).Bold = True
    Next I
  Next n

  cmdPrint.Enabled = False
  Set treSingle.SelectedItem = treSingle.Nodes(1)
  timRestrict.Interval = 50
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Call RetForm(Me, False)
  
End Sub

Private Sub timRestrict_Timer()

  If Mid(treSingle.SelectedItem.Text, 4, 4) = ":" And _
    treSingle.SelectedItem.Bold = False Then cmdPrint.Enabled = False
  
  If cmdToggle.Caption = "&Section" Then
    cmdOpen.Enabled = False
  Else
    cmdOpen.Enabled = cmdPrint.Enabled
  End If
  
End Sub

Private Sub treSingle_Click()

  cmdPrint.Enabled = True
  
  If cmdToggle.Caption = "&Section" Then
    If treSingle.SelectedItem.Children Then cmdPrint.Enabled = False
    If treSingle.SelectedItem.Parent Is Nothing Then
      cmdPrint.Enabled = True
    Else
      cmdPrint.Enabled = False
    End If
  Else
    cmdPrint.Enabled = Not treSingle.SelectedItem.Children
  End If
  
  If treSingle.SelectedItem.Children And cmdToggle.Caption = "&Section" Then _
    cmdPrint.Enabled = True
  
  If Mid$(treSingle.SelectedItem.Text, 4, 1) = ":" And _
    treSingle.SelectedItem.Bold = Not True Then cmdPrint.Enabled = False
  
End Sub

