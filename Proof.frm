VERSION 5.00
Begin VB.Form frmProof 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Here's the Proof!"
   ClientHeight    =   1935
   ClientLeft      =   7860
   ClientTop       =   3255
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3255
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "All Pages"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrintSection 
      Caption         =   "Section"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrintSingle 
      Caption         =   "Single Page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit Page Names and Locations"
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Print:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPages 
      Caption         =   "&Pages"
      Begin VB.Menu mnuPrintSingle 
         Caption         =   "Print &Single Page"
      End
      Begin VB.Menu mnuPrintSection 
         Caption         =   "Print S&ection"
      End
      Begin VB.Menu mnuPrintAll 
         Caption         =   "Print &All Pages"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "E&dit Pages and Locations"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmProof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEdit_Click()

Dim Pswd As String

If InputBox("Enter Editing Password", "Password") = "Admin" Then Call RetForm(frmEdit, True)

End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdPrintAll_Click()

  Dim CopInt As Integer
  CopInt = InputBox("Copies", "How many copies?", "1")
  If CopInt < 1 Or CopInt > 99 Then
    MsgBox "Please enter number over zero."
    Exit Sub
  End If
  
  Dim CopyRepeat As Integer
  For CopyRepeat = 1 To CopInt
  Dim n As Integer
    For n = LBound(Locations) To UBound(Locations)
      If Locations(UBound(Locations) - n) <> vbNullString And Mid$(Locations(UBound(Locations) - n), 1, 4) <> "Note" Then _
        SubmitDoc Locations(UBound(Locations) - n), CopInt
    Next n
  Next CopyRepeat
  RetForm frmPrint, True
  
End Sub

Private Sub cmdPrintSection_Click()

  frmSingle.cmdToggle.Caption = "&Section"
  frmSingle.Caption = "Print Section"
  Call RetForm(frmSingle, True)
  
End Sub

Private Sub cmdPrintSingle_Click()

  frmSingle.cmdToggle.Caption = "&Single"
  frmSingle.Caption = "Print Single Page"
  Call RetForm(frmSingle, True)
  
End Sub

Private Sub Form_Load()
  Set MainFrm = frmProof
  RefreshFromFile
End Sub

Private Sub mnuAbout_Click()
  Call RetForm(frmAbout, True)
End Sub

Private Sub mnuEdit_Click()
  cmdEdit_Click
End Sub

Private Sub mnuExit_Click()
  cmdExit_Click
End Sub

Private Sub mnuPrintAll_Click()
  cmdPrintAll_Click
End Sub

Private Sub mnuPrintSection_Click()
  cmdPrintSection_Click
End Sub

Private Sub mnuPrintSingle_Click()
  cmdPrintSingle_Click
End Sub

