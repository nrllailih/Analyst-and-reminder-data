VERSION 5.00
Begin VB.Form kb 
   Caption         =   "Form3"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13575
   LinkTopic       =   "Form3"
   ScaleHeight     =   9105
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2640
      TabIndex        =   23
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H8000000A&
      Caption         =   "BI Checking"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   6000
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   12480
      Picture         =   "kb.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   20
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2640
      TabIndex        =   19
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kirim Email"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   17
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H8000000A&
      Caption         =   "Surat Peringatan"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H8000000A&
      Caption         =   "Loan Inquiry"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H8000000A&
      Caption         =   "SID"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H8000000A&
      Caption         =   "Rekening Koran"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H8000000A&
      Caption         =   "Perjanjian Kredit"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H8000000A&
      Caption         =   "KTP"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000A&
      Caption         =   "Berita Acara Klaim"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   5640
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Text            =   "Nama Bank"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Tanggal Pengajuan"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   22
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Debitur"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label10 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   8520
      TabIndex        =   16
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Label Label9 
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Tambahan Berkas"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   14
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   13
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8520
      TabIndex        =   12
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Penerima"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Berkas"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "kb.frx":35D3
      Top             =   0
      Width           =   13500
   End
End
Attribute VB_Name = "kb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pesan1, pesan2, pesan3, pesan4, pesan5, pesan6, pesan7, pesan8, ps8, ps1, ps2, ps3, ps4, ps5, ps6, ps7, deb, tglp, inf, bts As String
Dim oSmtp As New EASendMailObjLib.Mail

Private Sub Command1_Click()
tgl5 = DateAdd("m", 6, Text2.Text)
tglp = "Menindak lanjuti mengenai kelengkapan berkas pengajuan klaim KUR ONLINE BNI tanggal "
deb = ". Kami sampaikan bahwa debitur atas nama "
inf = " masih terdapat kekurangan data. Data yang kami maksud antara lain: "
bts = ". Sesuai Perjanjian Kerja Sama antara Bank BNI dengan PT.Askrindo Pasal 18 Ayat 2 tentang DALUARSA HAK KLAIM menyebutkan bahwa penerima jaminan tidak melengkapi dokumen yang menjadi persyaratan klaim dalam waktu 6 (enam) bulan dari surat permintaan terakhir untuk melengkapi dokumen tersebut dari penjamin. Berkenaan dengan hal tersebut, kami mohon kekurangan data dapat segera dipenuhi sebelum tanggal "


If Check1.Value = 1 Then
pesan1 = "BAK ada" & vbCrLf
Else
ps1 = " Berita Acara Klaim, "
End If

If Check2.Value = 1 Then
pesan2 = "KTP ada" & vbCrLf
Else
ps2 = "KTP, "
End If

If Check3.Value = 1 Then
pesan3 = "Perjanjian Kredit" & vbCrLf
Else
ps3 = "Perjanjjian Kredit, "
End If

If Check4.Value = 1 Then
pesan4 = "Rekening koran ada" & vbCrLf
Else
ps4 = "Rekening Koran, "
End If

If Check5.Value = 1 Then
pesan5 = "SID ada" & vbCrLf
Else
ps5 = "SID, "
End If

If Check6.Value = 1 Then
pesan6 = "Loan Inquiry ada" & vbCrLf
Else
ps6 = "Loan Inquiry, "
End If


If Check7.Value = 1 Then
pesan7 = "Surat Peringatan ada" & vbCrLf
Else
ps7 = "Surat Peringatan, "
End If

If Check8.Value = 1 Then
pesan8 = "BI Checking ada" & vbCrLf
Else
ps8 = "Surat Peringatan, "
End If

MsgBox pesan1 & pesan2 & pesan3 & pesan4 & pesan5 & pesan6 & pesan7
Label10 = tglp & Text2.Text & deb & Text1.Text & inf & ps1 & ps2 & ps3 & ps4 & ps5 & ps6 & ps7 & bts & tgl5
End Sub

Private Sub Command2_Click()
Dim oSmtp As New EASendMailObjLib.Mail
    oSmtp.LicenseCode = "TryIt"
    
    ' Set your Gmail email address
    oSmtp.FromAddr = ""   'Enter your Email ID here
    
    ' Add recipient email address
    oSmtp.AddRecipientEx Label6, 0   'Enter Reciver Email ID here
    
    ' Set email subject
    oSmtp.Subject = "Tambahan Berkas KUR BNI"
    
    ' Set email body
    oSmtp.BodyText = Label10
       
     
    ' Gmail SMTP server address
    oSmtp.ServerAddr = "smtp.gmail.com"
    
    ' set direct SSL 465 port,
    oSmtp.ServerPort = 465
    
    ' detect SSL/TLS automatically
    oSmtp.SSL_init

    ' Gmail user authentication should use your
    ' Gmail email address as the user name.

    oSmtp.UserName = "" 'Enter your Email ID here again
    oSmtp.Password = ""    'Enter Your Mail Password
    
    MsgBox "start to send email ..."

    If oSmtp.SendMail() = 0 Then
        MsgBox "email was sent successfully!"
    Else
        MsgBox "failed to send email with the following error:" & oSmtp.GetLastErrDescription()
    End If
End Sub

Private Sub Form_Load()
Combo1.AddItem ("SKC Lula")
Combo1.AddItem ("SKC Tri")
Combo1.AddItem ("SKC coba")
End Sub
Private Sub Combo1_Click()
If Combo1.Text = "SKC Lula" Then
Label6 = ""
ElseIf Combo1.Text = "SKC Tri" Then
Label6 = ""
Else
Label6 = ""
End If
End Sub


Private Sub Picture1_Click()
about.Hide
Form2.Show
Form1.Hide
kb.Hide
analisis.Hide
End Sub
