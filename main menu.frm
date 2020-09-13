VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16320
   LinkTopic       =   "Form2"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   615
      Left            =   12120
      OLEDropMode     =   1  'Manual
      Picture         =   "main menu.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   8160
      Width           =   615
   End
   Begin VB.PictureBox Picture3 
      Height          =   615
      Left            =   1080
      Picture         =   "main menu.frx":35D3
      ScaleHeight     =   555
      ScaleWidth      =   2235
      TabIndex        =   2
      Top             =   5640
      Width           =   2295
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   1080
      Picture         =   "main menu.frx":80BA
      ScaleHeight     =   500
      ScaleMode       =   0  'User
      ScaleWidth      =   2200
      TabIndex        =   1
      Top             =   4680
      Width           =   2295
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reminder"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   1080
      Picture         =   "main menu.frx":C5E4
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   723.301
      TabIndex        =   0
      Top             =   3720
      Width           =   2295
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Analist"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "main menu.frx":10A78
      Top             =   0
      Width           =   13500
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub Label1_Click()
analisis.Show
Form2.Hide
Form1.Hide
about.Hide
kb.Hide
End Sub

Private Sub Label2_Click()
kb.Show
Form2.Hide
Form1.Hide
about.Hide
analisis.Hide
End Sub

Private Sub Label3_Click()
about.Show
Form2.Hide
Form1.Hide
kb.Hide
analisis.Hide
End Sub

Private Sub Picture4_Click()
about.Hide
Form2.Hide
Form1.Show
kb.Hide
analisis.Hide
End Sub
