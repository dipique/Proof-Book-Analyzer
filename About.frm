VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1665
   ClientLeft      =   7500
   ClientTop       =   5115
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4125
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Updated June 10, 2004"
      Height          =   195
      Left            =   1170
      TabIndex        =   3
      Top             =   840
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "By Dan Kaschel"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1140
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Here's the Proof! V1.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   135
      TabIndex        =   1
      Top             =   120
      Width           =   3825
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

Form_Unload (1)

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Call RetForm(Me, False)
  
End Sub

