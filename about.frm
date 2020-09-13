VERSION 5.00
Begin VB.Form about 
   Caption         =   "Form3"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   12840
      Picture         =   "about.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   8160
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   120
      Picture         =   "about.frx":35D3
      Top             =   0
      Width           =   13500
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()
about.Hide
Form2.Show
Form1.Hide
kb.Hide
analisis.Hide
End Sub
