VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   6840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "  password  :"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "  username :"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "login.frx":0000
      Top             =   0
      Width           =   13500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()

End Sub

Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub


Private Sub Command1_Click()
Call konekdb
If Text1.Text = " " Then
MsgBox "Username masih kosong !", vbCritical, "Perhatian"
Text1.SetFocus
ElseIf Text2.Text = " " Then
MsgBox "Username masih kosong !", vbCritical, "Perhatian"
Text2.SetFocus
Else
query = "select * from login where username='" & Text1.Text & "' and password='" & Text2.Text & "'"
RsAdmin.Open (query), koneksi
    If RsAdmin.EOF Then
    MsgBox "Username atau Password Salah!", vbExclamation, "Gagal!"
    Text1.Text = " "
    Text2.Text = " "
    Text1.SetFocus
    Text2.SetFocus
    Else
    Unload Me
    Form2.Show
    End If
End If
    


End Sub

