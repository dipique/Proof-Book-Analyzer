VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Pages and Locations"
   ClientHeight    =   5745
   ClientLeft      =   3870
   ClientTop       =   2055
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   1800
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   4680
      Width           =   735
   End
   Begin VB.Timer timName 
      Interval        =   100
      Left            =   1440
      Top             =   2640
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtLocation 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   4680
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4200
      Width           =   3135
   End
   Begin MSComctlLib.TreeView treMain 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7011
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lblLocations 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Location:"
      Height          =   195
      Left            =   45
      TabIndex        =   2
      Top             =   4740
      Width           =   660
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   4275
      Width           =   465
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
  If txtLocation.Enabled = True Then
    cdMain.FileName = vbNullString
    cdMain.ShowOpen
    If cdMain.FileName <> vbNullString Then txtLocation.Text = cdMain.FileName
  End If
  
End Sub

Private Sub cmdCancel_Click()
  Form_Unload (1)
End Sub

Private Sub cmdClear_Click()
  txtLocation.Text = vbNullString
End Sub

Private Sub cmdOK_Click()

  ChDir App.Path
  Dim FileString As String
  FileString = "Initiation Files\pages.loc"
  Dim n As Integer
  Open FileString For Output As #1
  For n = LBound(Locations) To UBound(Locations)
    Print #1, Locations(n)
  Next n
  Close #1
  Form_Unload (1)
  
End Sub

Private Sub Form_Load()

  Me.Show
  treMain.Nodes.Clear
  'Create All Nodes
  Dim n As Integer
  Dim PreviousNode As String
  For n = 1 To UBound(PageNames)
    If PageChild(n) = True Then                             ' if line is subcategory
      treMain.Nodes.Add PreviousNode, tvwChild, PageNames(n), PageNames(n)
    Else                                                    ' if line is root node
      treMain.Nodes.Add , , PageNames(n), PageNames(n)
      PreviousNode = PageNames(n)
    End If
  Next n

  Dim I As Integer
  For n = 1 To treMain.Nodes.Count 'Make all With Children Bold
    For I = 1 To treMain.Nodes(n).Children
      If I = 1 Then treMain.Nodes(n).Bold = True
    Next I
  Next n

  Set treMain.SelectedItem = treMain.Nodes(1)
  treMain_Click
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Call RetForm(Me, False)
  
End Sub

Private Sub timName_Timer()
  treMain_Click
  If txtLocation.Enabled = True Then
    If txtLocation.SelLength = 0 Then
      txtLocation.SelStart = Len(txtLocation.Text)
    End If
  End If
  If txtLocation.Enabled = True Then
    txtLocation.BackColor = vbWhite
  Else
    txtLocation.BackColor = &H8000000F
  End If
End Sub

Private Sub treMain_Click()
  txtName.Text = treMain.SelectedItem.Text
  txtLocation.Text = Locations(treMain.SelectedItem.Index)
  If PageChild(treMain.SelectedItem.Index) = False Then
    txtLocation.Enabled = False
  Else
    txtLocation.Enabled = True
  End If
  
End Sub

Private Sub treMain_DblClick()
  cmdBrowse_Click
End Sub

Private Sub txtLocation_Change()
  Call UpdateLocation(txtLocation.Text, treMain.SelectedItem.Index)
End Sub

