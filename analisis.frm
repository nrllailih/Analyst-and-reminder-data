VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form analisis 
   Caption         =   "Form3"
   ClientHeight    =   9180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14130
   LinkTopic       =   "Form3"
   ScaleHeight     =   9180
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   12600
      Picture         =   "analisis.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   70
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   1920
      TabIndex        =   69
      Text            =   "Text22"
      Top             =   6360
      Width           =   3135
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   2760
      TabIndex        =   67
      Text            =   "Text21"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   4320
      TabIndex        =   65
      Text            =   "Text20"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   1800
      TabIndex        =   64
      Text            =   "Text16"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2160
      TabIndex        =   59
      Text            =   "Text9"
      Top             =   6000
      Width           =   2535
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   5520
      TabIndex        =   43
      Text            =   "Text19"
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   3960
      TabIndex        =   42
      Text            =   "Text18"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1680
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   8880
      TabIndex        =   39
      Top             =   4560
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   11520
      Top             =   4440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\anremix_pkl\bismillah ini program\dbpkl.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\anremix_pkl\bismillah ini program\dbpkl.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "hasil"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   8640
      TabIndex        =   38
      Text            =   "Text17"
      Top             =   840
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   11040
      Picture         =   "analisis.frx":35D3
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   37
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Batal"
      Height          =   375
      Left            =   11400
      TabIndex        =   35
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proses"
      Height          =   375
      Left            =   11400
      TabIndex        =   34
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   8280
      TabIndex        =   33
      Text            =   "Text15"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   7800
      TabIndex        =   32
      Text            =   "Text14"
      Top             =   1920
      Width           =   3135
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   2400
      TabIndex        =   28
      Text            =   "Text13"
      Top             =   8520
      Width           =   2655
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1920
      TabIndex        =   27
      Text            =   "Text12"
      Top             =   8160
      Width           =   3135
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1680
      TabIndex        =   26
      Text            =   "Text11"
      Top             =   7800
      Width           =   3375
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Text            =   "Text10"
      Top             =   7440
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   5640
      Width           =   3735
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Text            =   "Text7"
      Top             =   4920
      Width           =   3495
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   4560
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal :"
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
      Index           =   26
      Left            =   960
      TabIndex        =   68
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal Pengajuan :"
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
      Index           =   25
      Left            =   960
      TabIndex        =   66
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Kol kmk :"
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
      Index           =   24
      Left            =   3240
      TabIndex        =   63
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Kol ki :"
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
      Index           =   23
      Left            =   960
      TabIndex        =   62
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   255
      Left            =   8040
      TabIndex        =   61
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "No rekening :"
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
      Index           =   22
      Left            =   6840
      TabIndex        =   60
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   255
      Left            =   10920
      TabIndex        =   58
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   11040
      TabIndex        =   57
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   11280
      TabIndex        =   56
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   8280
      TabIndex        =   55
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   11280
      TabIndex        =   54
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   7800
      TabIndex        =   53
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   7560
      TabIndex        =   52
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal :"
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
      Index           =   21
      Left            =   10080
      TabIndex        =   51
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Status :"
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
      Index           =   20
      Left            =   10080
      TabIndex        =   50
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Kolekbilitas :"
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
      Index           =   19
      Left            =   10080
      TabIndex        =   49
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Plafond Kredit :"
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
      Index           =   18
      Left            =   6840
      TabIndex        =   48
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Outstanding :"
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
      Index           =   17
      Left            =   10080
      TabIndex        =   47
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "No KTP :"
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
      Index           =   16
      Left            =   6840
      TabIndex        =   46
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Nama :"
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
      Index           =   15
      Left            =   6840
      TabIndex        =   45
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "s/d"
      Height          =   255
      Left            =   5160
      TabIndex        =   44
      Top             =   7440
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "JWK :"
      Height          =   255
      Left            =   3360
      TabIndex        =   41
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "HASIL"
      BeginProperty Font 
         Name            =   "Bell Gothic Std Light"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   36
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Plafond Kredit :"
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
      Index           =   14
      Left            =   6840
      TabIndex        =   31
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal :"
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
      Index           =   13
      Left            =   6840
      TabIndex        =   30
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Sertifikat Penjamin."
      BeginProperty Font 
         Name            =   "Bell Gothic Std Light"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   6840
      TabIndex        =   29
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "No KTP :"
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
      Index           =   12
      Left            =   960
      TabIndex        =   24
      Top             =   8160
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Nama :"
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
      Index           =   11
      Left            =   960
      TabIndex        =   23
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Plafond Kredit :"
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
      Index           =   10
      Left            =   960
      TabIndex        =   22
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Tanggal :"
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
      Index           =   9
      Left            =   960
      TabIndex        =   21
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Perjanjian Kredit."
      BeginProperty Font 
         Name            =   "Bell Gothic Std Light"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   20
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Kolekbilitas :"
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
      Index           =   8
      Left            =   960
      TabIndex        =   19
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "No KTP :"
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
      Index           =   7
      Left            =   960
      TabIndex        =   17
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "SID."
      BeginProperty Font 
         Name            =   "Bell Gothic Std Light"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   16
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Outstanding :"
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
      Index           =   6
      Left            =   960
      TabIndex        =   13
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Nama :"
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
      Index           =   5
      Left            =   960
      TabIndex        =   12
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "No KTP :"
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
      Index           =   4
      Left            =   960
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Identitas"
      BeginProperty Font 
         Name            =   "Bell Gothic Std Light"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Saldo Akhir :"
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
      Index           =   3
      Left            =   960
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "No rekening :"
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
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Nama :"
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
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Rekening Koran."
      BeginProperty Font 
         Name            =   "Bell Gothic Std Light"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Nama :"
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
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Berita Acara Klaim."
      BeginProperty Font 
         Name            =   "Bell Gothic Std Light"
         Size            =   12
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   0
      Picture         =   "analisis.frx":6B8A
      Top             =   0
      Width           =   13500
   End
End
Attribute VB_Name = "analisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

nama1 = Text1.Text
nama2 = Text3.Text
nama3 = Text5.Text
nama4 = Text11.Text
tgl5 = DateAdd("m", 6, Text19.Text)



If nama1 = nama2 And (nama2 = nama3) And (nama3 = nama4) Then
var1 = Text1.Text
Label5 = var1
Else
var1 = "Tidak valid"
Label5 = var1
End If
If Text2.Text = Text8.Text And (Text8.Text = Text12.Text) Then
var2 = Text2.Text
Label6 = var2
Else
var2 = "Tidak valid"
Label6 = var2
End If
uang1 = Text7.Text
uang2 = Text4.Text
If Text7.Text = Text4.Text Then
var3 = Text7.Text
Label7 = var3
Else
var3 = "Tidak valid"
Label7 = var3
End If

tgl2 = Text21.Text
If Text10.Text = Text14.Text And (Text14.Text = Text18.Text) Then
var4 = Text10.Text
Label10 = var4
Else
var4 = "Tidak valid"
Label10 = var4
Label11 = "TOLAK"
End If
If Text22.Text = Text10.Text And (Val(Text16.Text) = "2") And (Val(Text20.Text) = "2") Then
Label11 = "Tolak"
ElseIf Text22.Text = Text10.Text And (Val(Text20.Text) = "2") Then
Label11 = "Tolak"
End If


uang3 = Text15.Text
uang4 = Text13.Text
If Text15.Text = Text13.Text Then
Var5 = Text15.Text
Label8 = Var5
Else
Var5 = "Tidak valid"
Label8 = Var5
End If
Label9 = Text9.Text
Label12 = Text6.Text
var6 = Label9
tgl3 = Text18.Text

tgl4 = Text19.Text

If Val(Text9.Text) = "5" And CDate(tgl2) < CDate(tgl5) Then
Label11 = "BAYAR"
ElseIf Val(Text9.Text) = "1" And CDate(tgl4) < CDate(tgl2) And CDate(tgl2) < CDate(tgl5) Then
Label11 = "BAYAR"
ElseIf Val(Text9.Text) = "2" And CDate(tgl4) < CDate(tgl2) And CDate(tgl2) < CDate(tgl5) Then
Label11 = "BAYAR"
ElseIf Val(Text9.Text) = "3" And CDate(tgl4) < CDate(tgl2) And CDate(tgl2) < CDate(tgl5) Then
Label11 = "BAYAR"
ElseIf Val(Text9.Text) = "4" And CDate(tgl4) < CDate(tgl2) And CDate(tgl2) < CDate(tgl5) Then
Label11 = "BAYAR"
ElseIf Val(Text9.Text) = "4" And CDate(tgl3) < CDate(tgl2) And CDate(tgl2) < CDate(tgl4) Then
Label11 = "BAYAR"
ElseIf Val(Text9.Text) = "5" And CDate(tgl3) < CDate(tgl2) And CDate(tgl2) < CDate(tgl4) Then
Label11 = "BAYAR"
Else
Label11 = "TOLAK"
End If

End Sub

Sub bersih()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text15.Text = ""
Text18.Text = ""
Text19.Text = ""
End Sub


Private Sub Command2_Click()
bersih
Text1.SetFocus
End Sub

Private Sub Command3_Click()
Call konekdb
        strSQL = "INSERT INTO hasil (norek, nama, noktp, outstanding, plaf, kolek, tanggal, status)" _
                & "VALUES (' " & Text6.Text & " ', ' " & Text1.Text & " ', ' " & Text2.Text & " ', ' " & Text4.Text & " ','" & Text13.Text & " ',' " & Text9.Text & " ',' " & Text10.Text & " ',' " & Label11 & "')"
        koneksi.Execute strSQL, , adCmdText
        MsgBox "konek"
        Text6.SetFocus

End Sub

'Private Sub Picture1_Click()
'Adodc1.Recordset.Find "norek='" + Text17.Text + "'", , adSearchForward, 1
'If Text17.Text = " " Then
'MsgBox "Anda belum memasukan nomer rekening debitur", vbQuestion, "CARI DEBITUR"
'strSQL.Refresh
'Text17.SetFocus
'If Not strSQL.Recordset.EOF Then
'rsstrSQL.Refresh
'Else
'MsgBox "Debitur tidak ada!", vbQuestion, "CARI DEBITUR"
'Text17.SetFocus
'strSQL.Refresh
'End If
'End If




'End Sub

Private Sub Picture2_Click()
about.Hide
Form2.Show
Form1.Hide
kb.Hide
analisis.Hide
End Sub
