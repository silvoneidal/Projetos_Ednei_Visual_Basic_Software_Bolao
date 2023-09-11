VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTreino 
   ClientHeight    =   10050
   ClientLeft      =   120
   ClientTop       =   720
   ClientWidth     =   25470
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   25470
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   278
      Top             =   8640
      Width           =   2415
   End
   Begin VB.Timer tmrRelogio 
      Interval        =   1000
      Left            =   1080
      Top             =   9120
   End
   Begin VB.Timer tmrDownload 
      Interval        =   3000
      Left            =   1560
      Top             =   9120
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   9120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox ptrCupom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   20400
      ScaleHeight     =   8265
      ScaleMode       =   0  'User
      ScaleWidth      =   4965
      TabIndex        =   276
      Top             =   240
      Width           =   5000
   End
   Begin VB.Frame Frame4 
      Caption         =   "JOGADOR 4"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   134
      Top             =   6600
      Width           =   20175
      Begin VB.CommandButton cmdReserva4 
         Caption         =   "Reserva"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   275
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdIniciar4 
         Caption         =   "Iniciar"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   271
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   23
         ItemData        =   "Form1.frx":0000
         Left            =   16680
         List            =   "Form1.frx":0002
         TabIndex        =   266
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   22
         ItemData        =   "Form1.frx":0004
         Left            =   16080
         List            =   "Form1.frx":0006
         TabIndex        =   265
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   21
         ItemData        =   "Form1.frx":0008
         Left            =   15480
         List            =   "Form1.frx":000A
         TabIndex        =   264
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   20
         ItemData        =   "Form1.frx":000C
         Left            =   14880
         List            =   "Form1.frx":000E
         TabIndex        =   263
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   19
         ItemData        =   "Form1.frx":0010
         Left            =   14280
         List            =   "Form1.frx":0012
         TabIndex        =   262
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   18
         ItemData        =   "Form1.frx":0014
         Left            =   13680
         List            =   "Form1.frx":0016
         TabIndex        =   261
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         ItemData        =   "Form1.frx":0018
         Left            =   12480
         List            =   "Form1.frx":001A
         TabIndex        =   260
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         ItemData        =   "Form1.frx":001C
         Left            =   11880
         List            =   "Form1.frx":001E
         TabIndex        =   259
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         ItemData        =   "Form1.frx":0020
         Left            =   11280
         List            =   "Form1.frx":0022
         TabIndex        =   258
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         ItemData        =   "Form1.frx":0024
         Left            =   10680
         List            =   "Form1.frx":0026
         TabIndex        =   257
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         ItemData        =   "Form1.frx":0028
         Left            =   10080
         List            =   "Form1.frx":002A
         TabIndex        =   256
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         ItemData        =   "Form1.frx":002C
         Left            =   9480
         List            =   "Form1.frx":002E
         TabIndex        =   255
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         ItemData        =   "Form1.frx":0030
         Left            =   7320
         List            =   "Form1.frx":0032
         TabIndex        =   254
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         ItemData        =   "Form1.frx":0034
         Left            =   6720
         List            =   "Form1.frx":0036
         TabIndex        =   253
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         ItemData        =   "Form1.frx":0038
         Left            =   6120
         List            =   "Form1.frx":003A
         TabIndex        =   252
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         ItemData        =   "Form1.frx":003C
         Left            =   5520
         List            =   "Form1.frx":003E
         TabIndex        =   251
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         ItemData        =   "Form1.frx":0040
         Left            =   4920
         List            =   "Form1.frx":0042
         TabIndex        =   250
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         ItemData        =   "Form1.frx":0044
         Left            =   4320
         List            =   "Form1.frx":0046
         TabIndex        =   249
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         ItemData        =   "Form1.frx":0048
         Left            =   3120
         List            =   "Form1.frx":004A
         TabIndex        =   248
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         ItemData        =   "Form1.frx":004C
         Left            =   2520
         List            =   "Form1.frx":004E
         TabIndex        =   247
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         ItemData        =   "Form1.frx":0050
         Left            =   1920
         List            =   "Form1.frx":0052
         TabIndex        =   246
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         ItemData        =   "Form1.frx":0054
         Left            =   1320
         List            =   "Form1.frx":0056
         TabIndex        =   245
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         ItemData        =   "Form1.frx":0058
         Left            =   720
         List            =   "Form1.frx":005A
         TabIndex        =   244
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         ItemData        =   "Form1.frx":005C
         Left            =   120
         List            =   "Form1.frx":005E
         TabIndex        =   243
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   57
         Left            =   18840
         TabIndex        =   173
         Text            =   "Total Final"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtTotal4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   7
         Left            =   18840
         Locked          =   -1  'True
         TabIndex        =   172
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   58
         Left            =   17880
         TabIndex        =   171
         Text            =   "T3+T4"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   59
         Left            =   17280
         TabIndex        =   170
         Text            =   "T4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   60
         Left            =   16680
         TabIndex        =   169
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   61
         Left            =   16080
         TabIndex        =   168
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   62
         Left            =   15480
         TabIndex        =   167
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   63
         Left            =   14880
         TabIndex        =   166
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text59 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14280
         TabIndex        =   165
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text60 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13680
         TabIndex        =   164
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   64
         Left            =   13080
         TabIndex        =   163
         Text            =   "T3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   65
         Left            =   12480
         TabIndex        =   162
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   66
         Left            =   11880
         TabIndex        =   161
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   67
         Left            =   11280
         TabIndex        =   160
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text61 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10680
         TabIndex        =   159
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text62 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   158
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   68
         Left            =   9480
         TabIndex        =   157
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   69
         Left            =   8520
         TabIndex        =   156
         Text            =   "T1+T2"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text63 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   155
         Text            =   "T2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text64 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   154
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   70
         Left            =   6720
         TabIndex        =   153
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   71
         Left            =   6120
         TabIndex        =   152
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text65 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   151
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text66 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   150
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   72
         Left            =   4320
         TabIndex        =   149
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   73
         Left            =   3720
         TabIndex        =   148
         Text            =   "T1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text67 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   147
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text68 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   146
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   74
         Left            =   1920
         TabIndex        =   145
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   75
         Left            =   1320
         TabIndex        =   144
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text69 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   143
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text70 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   142
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtTotal4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   13080
         Locked          =   -1  'True
         TabIndex        =   141
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   17280
         Locked          =   -1  'True
         TabIndex        =   140
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   17880
         Locked          =   -1  'True
         TabIndex        =   139
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTotal4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   138
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTotal4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   137
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   136
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox cboName4 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0060
         Left            =   3240
         List            =   "Form1.frx":0062
         Style           =   2  'Dropdown List
         TabIndex        =   135
         Top             =   360
         Width           =   6000
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   300
         Left            =   720
         TabIndex        =   174
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53280769
         CurrentDate     =   45154
      End
      Begin VB.Label Label7 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   176
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   175
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "JOGADOR 3"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   91
      Top             =   4440
      Width           =   20175
      Begin VB.CommandButton cmdReserva3 
         Caption         =   "Reserva"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   274
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdIniciar3 
         Caption         =   "Iniciar"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   270
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   23
         ItemData        =   "Form1.frx":0064
         Left            =   16680
         List            =   "Form1.frx":0066
         TabIndex        =   242
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   22
         ItemData        =   "Form1.frx":0068
         Left            =   16080
         List            =   "Form1.frx":006A
         TabIndex        =   241
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   21
         ItemData        =   "Form1.frx":006C
         Left            =   15480
         List            =   "Form1.frx":006E
         TabIndex        =   240
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   20
         ItemData        =   "Form1.frx":0070
         Left            =   14880
         List            =   "Form1.frx":0072
         TabIndex        =   239
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   19
         ItemData        =   "Form1.frx":0074
         Left            =   14280
         List            =   "Form1.frx":0076
         TabIndex        =   238
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   18
         ItemData        =   "Form1.frx":0078
         Left            =   13680
         List            =   "Form1.frx":007A
         TabIndex        =   237
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         ItemData        =   "Form1.frx":007C
         Left            =   12480
         List            =   "Form1.frx":007E
         TabIndex        =   236
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         ItemData        =   "Form1.frx":0080
         Left            =   11880
         List            =   "Form1.frx":0082
         TabIndex        =   235
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         ItemData        =   "Form1.frx":0084
         Left            =   11280
         List            =   "Form1.frx":0086
         TabIndex        =   234
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         ItemData        =   "Form1.frx":0088
         Left            =   10680
         List            =   "Form1.frx":008A
         TabIndex        =   233
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         ItemData        =   "Form1.frx":008C
         Left            =   10080
         List            =   "Form1.frx":008E
         TabIndex        =   232
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         ItemData        =   "Form1.frx":0090
         Left            =   9480
         List            =   "Form1.frx":0092
         TabIndex        =   231
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         ItemData        =   "Form1.frx":0094
         Left            =   7320
         List            =   "Form1.frx":0096
         TabIndex        =   230
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         ItemData        =   "Form1.frx":0098
         Left            =   6720
         List            =   "Form1.frx":009A
         TabIndex        =   229
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         ItemData        =   "Form1.frx":009C
         Left            =   6120
         List            =   "Form1.frx":009E
         TabIndex        =   228
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         ItemData        =   "Form1.frx":00A0
         Left            =   5520
         List            =   "Form1.frx":00A2
         TabIndex        =   227
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         ItemData        =   "Form1.frx":00A4
         Left            =   4920
         List            =   "Form1.frx":00A6
         TabIndex        =   226
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         ItemData        =   "Form1.frx":00A8
         Left            =   4320
         List            =   "Form1.frx":00AA
         TabIndex        =   225
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         ItemData        =   "Form1.frx":00AC
         Left            =   3120
         List            =   "Form1.frx":00AE
         TabIndex        =   224
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         ItemData        =   "Form1.frx":00B0
         Left            =   2520
         List            =   "Form1.frx":00B2
         TabIndex        =   223
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         ItemData        =   "Form1.frx":00B4
         Left            =   1920
         List            =   "Form1.frx":00B6
         TabIndex        =   222
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         ItemData        =   "Form1.frx":00B8
         Left            =   1320
         List            =   "Form1.frx":00BA
         TabIndex        =   221
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         ItemData        =   "Form1.frx":00BC
         Left            =   720
         List            =   "Form1.frx":00BE
         TabIndex        =   220
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         ItemData        =   "Form1.frx":00C0
         Left            =   120
         List            =   "Form1.frx":00C2
         TabIndex        =   219
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox cboName3 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":00C4
         Left            =   3240
         List            =   "Form1.frx":00C6
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   360
         Width           =   6000
      End
      Begin VB.TextBox txtTotal3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   129
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   127
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTotal3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   17880
         Locked          =   -1  'True
         TabIndex        =   126
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTotal3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   17280
         Locked          =   -1  'True
         TabIndex        =   125
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   13080
         Locked          =   -1  'True
         TabIndex        =   124
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text51 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   123
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text50 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   122
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   56
         Left            =   1320
         TabIndex        =   121
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   55
         Left            =   1920
         TabIndex        =   120
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text49 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   119
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text48 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   118
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   54
         Left            =   3720
         TabIndex        =   117
         Text            =   "T1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   53
         Left            =   4320
         TabIndex        =   116
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text47 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   115
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text46 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   114
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   52
         Left            =   6120
         TabIndex        =   113
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   51
         Left            =   6720
         TabIndex        =   112
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text45 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   111
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text44 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   110
         Text            =   "T2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   50
         Left            =   8520
         TabIndex        =   109
         Text            =   "T1+T2"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   49
         Left            =   9480
         TabIndex        =   108
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text43 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   107
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text42 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10680
         TabIndex        =   106
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   48
         Left            =   11280
         TabIndex        =   105
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   47
         Left            =   11880
         TabIndex        =   104
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   46
         Left            =   12480
         TabIndex        =   103
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   45
         Left            =   13080
         TabIndex        =   102
         Text            =   "T3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text41 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13680
         TabIndex        =   101
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text40 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14280
         TabIndex        =   100
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   44
         Left            =   14880
         TabIndex        =   99
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   43
         Left            =   15480
         TabIndex        =   98
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   42
         Left            =   16080
         TabIndex        =   97
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   41
         Left            =   16680
         TabIndex        =   96
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   40
         Left            =   17280
         TabIndex        =   95
         Text            =   "T4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   39
         Left            =   17880
         TabIndex        =   94
         Text            =   "T3+T4"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtTotal3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   7
         Left            =   18840
         Locked          =   -1  'True
         TabIndex        =   93
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   38
         Left            =   18840
         TabIndex        =   92
         Text            =   "Total Final"
         Top             =   840
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   300
         Left            =   720
         TabIndex        =   130
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53280769
         CurrentDate     =   45154
      End
      Begin VB.Label Label6 
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   133
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "JOGADOR 2"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   49
      Top             =   2280
      Width           =   20175
      Begin VB.CommandButton cmdReserva2 
         Caption         =   "Reserva"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   273
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdIniciar2 
         Caption         =   "Iniciar"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   269
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cboName2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":00C8
         Left            =   3240
         List            =   "Form1.frx":00CA
         Style           =   2  'Dropdown List
         TabIndex        =   268
         Top             =   360
         Width           =   6000
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   23
         ItemData        =   "Form1.frx":00CC
         Left            =   16680
         List            =   "Form1.frx":00CE
         TabIndex        =   218
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   22
         ItemData        =   "Form1.frx":00D0
         Left            =   16080
         List            =   "Form1.frx":00D2
         TabIndex        =   217
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   21
         ItemData        =   "Form1.frx":00D4
         Left            =   15480
         List            =   "Form1.frx":00D6
         TabIndex        =   216
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   20
         ItemData        =   "Form1.frx":00D8
         Left            =   14880
         List            =   "Form1.frx":00DA
         TabIndex        =   215
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   19
         ItemData        =   "Form1.frx":00DC
         Left            =   14280
         List            =   "Form1.frx":00DE
         TabIndex        =   214
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   18
         ItemData        =   "Form1.frx":00E0
         Left            =   13680
         List            =   "Form1.frx":00E2
         TabIndex        =   213
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         ItemData        =   "Form1.frx":00E4
         Left            =   12480
         List            =   "Form1.frx":00E6
         TabIndex        =   212
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         ItemData        =   "Form1.frx":00E8
         Left            =   11880
         List            =   "Form1.frx":00EA
         TabIndex        =   211
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         ItemData        =   "Form1.frx":00EC
         Left            =   11280
         List            =   "Form1.frx":00EE
         TabIndex        =   210
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         ItemData        =   "Form1.frx":00F0
         Left            =   10680
         List            =   "Form1.frx":00F2
         TabIndex        =   209
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         ItemData        =   "Form1.frx":00F4
         Left            =   10080
         List            =   "Form1.frx":00F6
         TabIndex        =   208
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         ItemData        =   "Form1.frx":00F8
         Left            =   9480
         List            =   "Form1.frx":00FA
         TabIndex        =   207
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         ItemData        =   "Form1.frx":00FC
         Left            =   7320
         List            =   "Form1.frx":00FE
         TabIndex        =   206
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         ItemData        =   "Form1.frx":0100
         Left            =   6720
         List            =   "Form1.frx":0102
         TabIndex        =   205
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         ItemData        =   "Form1.frx":0104
         Left            =   6120
         List            =   "Form1.frx":0106
         TabIndex        =   204
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         ItemData        =   "Form1.frx":0108
         Left            =   5520
         List            =   "Form1.frx":010A
         TabIndex        =   203
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         ItemData        =   "Form1.frx":010C
         Left            =   4920
         List            =   "Form1.frx":010E
         TabIndex        =   202
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         ItemData        =   "Form1.frx":0110
         Left            =   4320
         List            =   "Form1.frx":0112
         TabIndex        =   201
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         ItemData        =   "Form1.frx":0114
         Left            =   3120
         List            =   "Form1.frx":0116
         TabIndex        =   200
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         ItemData        =   "Form1.frx":0118
         Left            =   2520
         List            =   "Form1.frx":011A
         TabIndex        =   199
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         ItemData        =   "Form1.frx":011C
         Left            =   1920
         List            =   "Form1.frx":011E
         TabIndex        =   198
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         ItemData        =   "Form1.frx":0120
         Left            =   1320
         List            =   "Form1.frx":0122
         TabIndex        =   197
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         ItemData        =   "Form1.frx":0124
         Left            =   720
         List            =   "Form1.frx":0126
         TabIndex        =   196
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         ItemData        =   "Form1.frx":0128
         Left            =   120
         List            =   "Form1.frx":012A
         TabIndex        =   195
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTotal2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   17880
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTotal2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   17280
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   13080
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   81
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   80
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   37
         Left            =   1320
         TabIndex        =   79
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   36
         Left            =   1920
         TabIndex        =   78
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   77
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   76
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   35
         Left            =   3720
         TabIndex        =   75
         Text            =   "T1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   34
         Left            =   4320
         TabIndex        =   74
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   73
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   72
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   33
         Left            =   6120
         TabIndex        =   71
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   32
         Left            =   6720
         TabIndex        =   70
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   69
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   68
         Text            =   "T2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   31
         Left            =   8520
         TabIndex        =   67
         Text            =   "T1+T2"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   30
         Left            =   9480
         TabIndex        =   66
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text24 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   65
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10680
         TabIndex        =   64
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   29
         Left            =   11280
         TabIndex        =   63
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   28
         Left            =   11880
         TabIndex        =   62
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   27
         Left            =   12480
         TabIndex        =   61
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   26
         Left            =   13080
         TabIndex        =   60
         Text            =   "T3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13680
         TabIndex        =   59
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14280
         TabIndex        =   58
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   25
         Left            =   14880
         TabIndex        =   57
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   24
         Left            =   15480
         TabIndex        =   56
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   23
         Left            =   16080
         TabIndex        =   55
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   22
         Left            =   16680
         TabIndex        =   54
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   21
         Left            =   17280
         TabIndex        =   53
         Text            =   "T4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   20
         Left            =   17880
         TabIndex        =   52
         Text            =   "T3+T4"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtTotal2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   7
         Left            =   18840
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   19
         Left            =   18840
         TabIndex        =   50
         Text            =   "Total Final"
         Top             =   840
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   720
         TabIndex        =   88
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53280769
         CurrentDate     =   45154
      End
      Begin VB.Label Label4 
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   90
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   360
         Width           =   495
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   9720
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   $"Form1.frx":012C
      OLEDBString     =   $"Form1.frx":01F1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TabelaTreino"
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
   Begin VB.Timer tmrRegisters 
      Interval        =   2000
      Left            =   600
      Top             =   9120
   End
   Begin VB.Frame Frame1 
      Caption         =   "JOGADOR 1"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   20175
      Begin VB.CommandButton cmdReserva1 
         Caption         =   "Reserva"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12000
         Style           =   1  'Graphical
         TabIndex        =   272
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdIniciar1 
         Caption         =   "Iniciar"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   267
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   23
         ItemData        =   "Form1.frx":02B6
         Left            =   16680
         List            =   "Form1.frx":02B8
         TabIndex        =   194
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   22
         ItemData        =   "Form1.frx":02BA
         Left            =   16080
         List            =   "Form1.frx":02BC
         TabIndex        =   193
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   21
         ItemData        =   "Form1.frx":02BE
         Left            =   15480
         List            =   "Form1.frx":02C0
         TabIndex        =   192
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   20
         ItemData        =   "Form1.frx":02C2
         Left            =   14880
         List            =   "Form1.frx":02C4
         TabIndex        =   191
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   19
         ItemData        =   "Form1.frx":02C6
         Left            =   14280
         List            =   "Form1.frx":02C8
         TabIndex        =   190
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   18
         ItemData        =   "Form1.frx":02CA
         Left            =   13680
         List            =   "Form1.frx":02CC
         TabIndex        =   189
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         ItemData        =   "Form1.frx":02CE
         Left            =   12480
         List            =   "Form1.frx":02D0
         TabIndex        =   188
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         ItemData        =   "Form1.frx":02D2
         Left            =   11880
         List            =   "Form1.frx":02D4
         TabIndex        =   187
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         ItemData        =   "Form1.frx":02D6
         Left            =   11280
         List            =   "Form1.frx":02D8
         TabIndex        =   186
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         ItemData        =   "Form1.frx":02DA
         Left            =   10680
         List            =   "Form1.frx":02DC
         TabIndex        =   185
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         ItemData        =   "Form1.frx":02DE
         Left            =   10080
         List            =   "Form1.frx":02E0
         TabIndex        =   184
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         ItemData        =   "Form1.frx":02E2
         Left            =   9480
         List            =   "Form1.frx":02E4
         TabIndex        =   183
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         ItemData        =   "Form1.frx":02E6
         Left            =   7320
         List            =   "Form1.frx":02E8
         TabIndex        =   182
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         ItemData        =   "Form1.frx":02EA
         Left            =   6720
         List            =   "Form1.frx":02EC
         TabIndex        =   181
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         ItemData        =   "Form1.frx":02EE
         Left            =   6120
         List            =   "Form1.frx":02F0
         TabIndex        =   180
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         ItemData        =   "Form1.frx":02F2
         Left            =   5520
         List            =   "Form1.frx":02F4
         TabIndex        =   179
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         ItemData        =   "Form1.frx":02F6
         Left            =   4920
         List            =   "Form1.frx":02F8
         TabIndex        =   178
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         ItemData        =   "Form1.frx":02FA
         Left            =   4320
         List            =   "Form1.frx":02FC
         TabIndex        =   177
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   18840
         TabIndex        =   48
         Text            =   "Total Final"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   7
         Left            =   18840
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "0"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   17880
         TabIndex        =   46
         Text            =   "T3+T4"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   17280
         TabIndex        =   45
         Text            =   "T4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   16680
         TabIndex        =   44
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   16080
         TabIndex        =   43
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   15480
         TabIndex        =   42
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   14880
         TabIndex        =   41
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   14280
         TabIndex        =   40
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13680
         TabIndex        =   39
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   13080
         TabIndex        =   38
         Text            =   "T3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   12480
         TabIndex        =   37
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   11880
         TabIndex        =   36
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   11280
         TabIndex        =   35
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10680
         TabIndex        =   34
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10080
         TabIndex        =   33
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   9480
         TabIndex        =   32
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   8520
         TabIndex        =   31
         Text            =   "T1+T2"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7920
         TabIndex        =   30
         Text            =   "T2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   29
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   6720
         TabIndex        =   28
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   6120
         TabIndex        =   27
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   26
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   25
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4320
         TabIndex        =   24
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   3720
         TabIndex        =   23
         Text            =   "T1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         TabIndex        =   22
         Text            =   "5"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   21
         Text            =   "4"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   20
         Text            =   "3"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   19
         Text            =   "2"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   18
         Text            =   "1"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Text            =   "E"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   13080
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   17280
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   17880
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtTotal1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         ItemData        =   "Form1.frx":02FE
         Left            =   3120
         List            =   "Form1.frx":0300
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         ItemData        =   "Form1.frx":0302
         Left            =   2520
         List            =   "Form1.frx":0304
         TabIndex        =   9
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         ItemData        =   "Form1.frx":0306
         Left            =   1920
         List            =   "Form1.frx":0308
         TabIndex        =   8
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         ItemData        =   "Form1.frx":030A
         Left            =   1320
         List            =   "Form1.frx":030C
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         ItemData        =   "Form1.frx":030E
         Left            =   120
         List            =   "Form1.frx":0310
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         ItemData        =   "Form1.frx":0312
         Left            =   720
         List            =   "Form1.frx":0314
         TabIndex        =   5
         Top             =   1200
         Width           =   615
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   53280769
         CurrentDate     =   45154
      End
      Begin VB.ComboBox cboName1 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0316
         Left            =   3240
         List            =   "Form1.frx":0318
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   6000
      End
      Begin VB.Label Label2 
         Caption         =   "Data:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label lblRelogio 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   17280
      TabIndex        =   277
      Top             =   8640
      Width           =   2895
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuCadastro 
         Caption         =   "Cadastrar Jogador"
      End
      Begin VB.Menu mnuTreinos 
         Caption         =   "Registro de Treinos"
      End
      Begin VB.Menu mnuControle 
         Caption         =   "Controle de Registros"
      End
      Begin VB.Menu mnuBancoDeDados 
         Caption         =   "Banco de Dados"
         Begin VB.Menu mnuTestBD 
            Caption         =   "Teste de Conexão"
         End
         Begin VB.Menu mnuAddressBKP 
            Caption         =   "Endereço do Backup"
         End
         Begin VB.Menu mnuAddressBD 
            Caption         =   "Endereço do Registro"
         End
      End
   End
End
Attribute VB_Name = "frmTreino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Componentes:
' Microsoft Windows common Controls 6.0 (SP6)
' Microsoft Windos Common Controls-2.6.0 (SP4)
' Microsoft ADO Data Control 6.0 (OLEDB)
' Microsoft DataGrid Control 6.0 (OLEDB)

' Variáveis de uso global
Dim DTPicker As DTPicker ' valores temporários
Dim cboName As ComboBox  ' valores temporários
Dim Combo(0 To 23) As ComboBox ' valores temporários
Dim txtTotal(1 To 7) As TextBox ' valores temporários
Dim treino As Boolean ' flag para treino ativo

' Variáveis para uso do banco de dados
Public addressRegisters As String ' endereço do banco de dados
Public addressBackups As String ' endereço de backup do banco de dados
Public nameRegisters As String ' nome do banco de dados
Dim query As String ' sql para o banco de dados

'//////////////////////////////////////////////////////////////////////////////////////////////
' SQL PARA CONSULTA NO BANCO DE DADOS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub queryString(query As String)
    ' Comando para SQL
    Adodc1.RecordSource = query
    Adodc1.CommandType = adCmdText
    Adodc1.Refresh
    
End Sub


'//////////////////////////////////////////////////////////////////////////////////////////////
' INICIO DO FORMULÁRIO TREINO ( PRINCIPAL )
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Form_Load()
On Error GoTo Erro

    ' Configurações iniciais para o banco de dados
    nameRegisters = "\RegistrosBolao" 'Name do registro
    addressRegisters = ReadIniValue(App.Path & "\Config.ini", "REGISTROS", "AddressBD") ' Endereço do registro
    addressBackups = ReadIniValue(App.Path & "\Config.ini", "REGISTROS", "AddressBKP") ' Endereço de backup
    Adodc1.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};" & "Dbq= " & frmTreino.addressRegisters & nameRegisters & ";" & "Uid=;"  ' Pwd=1234"
    
    ' Busca nomes cadastrados no registro
    Call updateCboNames
    
    ' Inicializa lista dos comboBox
    Dim x, y As Integer
    For x = 0 To 23
        For y = 0 To 9
            Combo1(x).AddItem (y)
            Combo2(x).AddItem (y)
            Combo3(x).AddItem (y)
            Combo4(x).AddItem (y)
        Next y
    Next x
    
    ' Atualiza data atual
    DTPicker1.value = Date
    DTPicker2.value = Date
    DTPicker3.value = Date
    DTPicker4.value = Date
    
Exit Sub
Erro:
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' INICIA E FINALIZA JOGADOR 1
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdIniciar1_Click()

    ' Verifica se nome em branco
    If cboName1.Text = Empty Then
        cboName1.BackColor = vbYellow
        MsgBox "Nenhum nome selecionado.", vbInformation, "DALCOQUIO AUTOMAÇÃO"
        cboName1.BackColor = vbWhite
        Exit Sub
    End If
    
    ' Verifica se Iniciar ou Finalizar
    If cmdIniciar1.Caption = "Iniciar" Then
        ' Mensagem de confirmação
        Beep
        Dim value As Integer
        value = MsgBox("Iniciar o Treino com: " & cboName1.Text & " ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 2 Then Exit Sub ' Não
        If cmdIniciar1.BackColor = yellow Then
            ' Novo Registro (Reserva)
            treino = True
            Call NewRegister(1)
        Else
            ' Novo Registro (Titular)
            treino = True
            cmdIniciar1.BackColor = vbGreen
            cmdIniciar1.Caption = "Finalizar"
            cmdReserva1.Enabled = True
            DTPicker1.Enabled = False
            cboName1.Enabled = False
            Call NewRegister(1)
        End If
    Else
        ' Mensagem de confirmação
        Beep
        value = MsgBox("Finalizar o Treino com: " & cboName1.Text & " ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 2 Then Exit Sub ' Não
        ' Mensagem de confirmação
        Beep
        value = MsgBox("Deseja Imprimir Cupom dos Resultados?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 1 Then Call PrintCupom(1) 'Sim
        ' Finalizar
        treino = False
        Call UploadRegisters(1)
        Call EnableCombo(1)
        Call ClearValue(1)
        cmdIniciar1.Caption = "Iniciar"
        cmdIniciar1.BackColor = &H8000000F
        cmdReserva1.BackColor = &H8000000F
        cmdReserva1.Enabled = False
        DTPicker1.Enabled = True
        cboName1.Enabled = True
    End If
    
Exit Sub
Erro:
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' HABILITA RESERVA PARA JOGADOR 1
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdReserva1_Click()
On Error GoTo Erro

    ' Mensagem de confirmação
    Beep
    value = MsgBox("Iniciar o treino com Reserva ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
    If value = 2 Then Exit Sub ' Não
     ' Mensagem de confirmação
    Beep
    value = MsgBox("Deseja Imprimir Cupom dos Resultados?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
    If value = 1 Then Call PrintCupom(1) 'Sim
    ' Habilita Reserva
    treino = False
    Call UploadRegisters(1)
    cmdIniciar1.Caption = "Iniciar"
    cmdIniciar1.BackColor = &H8000000F
    cmdReserva1.BackColor = vbYellow
    DTPicker1.Enabled = True
    cboName1.Enabled = True
    Call DesableCombo(1)
    Call ClearValue(1)
    
Exit Sub
Erro:
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' INICIA E FINALIZA JOGADOR 2
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdIniciar2_Click()

    ' Verifica se nome em branco
    If cboName2.Text = Empty Then
        cboName2.BackColor = vbYellow
        MsgBox "Nenhum nome selecionado.", vbInformation, "DALCOQUIO AUTOMAÇÃO"
        cboName2.BackColor = vbWhite
        Exit Sub
    End If
    
    ' Verifica se Iniciar ou Finalizar
    If cmdIniciar2.Caption = "Iniciar" Then
        ' Mensagem de confirmação
        Beep
        Dim value As Integer
        value = MsgBox("Iniciar o Treino com: " & cboName2.Text & " ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 2 Then Exit Sub ' Não
        If cmdIniciar2.BackColor = yellow Then
            ' Novo Registro (Reserva)
            treino = True
            Call NewRegister(2)
        Else
            ' Novo Registro (Titular)
            treino = True
            cmdIniciar2.BackColor = vbGreen
            cmdIniciar2.Caption = "Finalizar"
            cmdReserva2.Enabled = True
            DTPicker2.Enabled = False
            cboName2.Enabled = False
            Call NewRegister(2)
        End If
    Else
        ' Mensagem de confirmação
        Beep
        value = MsgBox("Finalizar o Treino com: " & cboName2.Text & " ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 2 Then Exit Sub ' Não
        ' Mensagem de confirmação
        Beep
        value = MsgBox("Deseja Imprimir Cupom dos Resultados?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 1 Then Call PrintCupom(3) 'Sim
        ' Finalizar
        treino = False
        Call UploadRegisters(2)
        Call EnableCombo(2)
        Call ClearValue(2)
        cmdIniciar2.Caption = "Iniciar"
        cmdIniciar2.BackColor = &H8000000F
        cmdReserva2.BackColor = &H8000000F
        cmdReserva2.Enabled = False
        DTPicker2.Enabled = True
        cboName2.Enabled = True
    End If
    
Exit Sub
Erro:
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' HABILITA RESERVA PARA JOGADOR 2
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdReserva2_Click()
On Error GoTo Erro

    ' Mensagem de confirmação
    Beep
    value = MsgBox("Iniciar o treino com Reserva ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
    If value = 2 Then Exit Sub ' Não
     ' Mensagem de confirmação
    Beep
    value = MsgBox("Deseja Imprimir Cupom dos Resultados?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
    If value = 1 Then Call PrintCupom(2) 'Sim
    ' Habilita Reserva
    treino = False
    Call UploadRegisters(2)
    cmdIniciar2.Caption = "Iniciar"
    cmdIniciar2.BackColor = &H8000000F
    cmdReserva2.BackColor = vbYellow
    DTPicker2.Enabled = True
    cboName2.Enabled = True
    Call DesableCombo(2)
    Call ClearValue(2)
    
Exit Sub
Erro:
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' INICIA E FINALIZA JOGADOR 3
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdIniciar3_Click()

    ' Verifica se nome em branco
    If cboName3.Text = Empty Then
        cboName3.BackColor = vbYellow
        MsgBox "Nenhum nome selecionado.", vbInformation, "DALCOQUIO AUTOMAÇÃO"
        cboName3.BackColor = vbWhite
        Exit Sub
    End If
    
    ' Verifica se Iniciar ou Finalizar
    If cmdIniciar3.Caption = "Iniciar" Then
        ' Mensagem de confirmação
        Beep
        Dim value As Integer
        value = MsgBox("Iniciar o Treino com: " & cboName3.Text & " ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 2 Then Exit Sub ' Não
        If cmdIniciar3.BackColor = yellow Then
            ' Novo Registro (Reserva)
            treino = True
            Call NewRegister(3)
        Else
            ' Novo Registro (Titular)
            treino = True
            cmdIniciar3.BackColor = vbGreen
            cmdIniciar3.Caption = "Finalizar"
            cmdReserva3.Enabled = True
            DTPicker3.Enabled = False
            cboName3.Enabled = False
            Call NewRegister(3)
        End If
    Else
        ' Mensagem de confirmação
        Beep
        value = MsgBox("Finalizar o Treino com: " & cboName3.Text & " ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 2 Then Exit Sub ' Não
        ' Mensagem de confirmação
        Beep
        value = MsgBox("Deseja Imprimir Cupom dos Resultados?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 1 Then Call PrintCupom(3) 'Sim
        ' Finalizar
        treino = False
        Call UploadRegisters(3)
        Call EnableCombo(3)
        Call ClearValue(3)
        cmdIniciar3.Caption = "Iniciar"
        cmdIniciar3.BackColor = &H8000000F
        cmdReserva3.BackColor = &H8000000F
        cmdReserva3.Enabled = False
        DTPicker3.Enabled = True
        cboName1.Enabled = True
    End If
    
Exit Sub
Erro:
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' HABILITA RESERVA PARA JOGADOR 3
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdReserva3_Click()
On Error GoTo Erro

    ' Mensagem de confirmação
    Beep
    value = MsgBox("Iniciar o treino com Reserva ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
    If value = 2 Then Exit Sub ' Não
     ' Mensagem de confirmação
    Beep
    value = MsgBox("Deseja Imprimir Cupom dos Resultados?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
    If value = 1 Then Call PrintCupom(3) 'Sim
    ' Habilita Reserva
    treino = False
    Call UploadRegisters(3)
    cmdIniciar3.Caption = "Iniciar"
    cmdIniciar3.BackColor = &H8000000F
    cmdReserva3.BackColor = vbYellow
    DTPicker3.Enabled = True
    cboName3.Enabled = True
    Call DesableCombo(3)
    Call ClearValue(3)
    
Exit Sub
Erro:
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' INICIA E FINALIZA JOGADOR 4
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdIniciar4_Click()

    ' Verifica se nome em branco
    If cboName4.Text = Empty Then
        cboName4.BackColor = vbYellow
        MsgBox "Nenhum nome selecionado.", vbInformation, "DALCOQUIO AUTOMAÇÃO"
        cboName4.BackColor = vbWhite
        Exit Sub
    End If
    
    ' Verifica se Iniciar ou Finalizar
    If cmdIniciar4.Caption = "Iniciar" Then
        ' Mensagem de confirmação
        Beep
        Dim value As Integer
        value = MsgBox("Iniciar o Treino com: " & cboName4.Text & " ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 2 Then Exit Sub ' Não
        If cmdIniciar4.BackColor = yellow Then
            ' Novo Registro (Reserva)
            treino = True
            Call NewRegister(4)
        Else
            ' Novo Registro (Titular)
            treino = True
            cmdIniciar4.BackColor = vbGreen
            cmdIniciar4.Caption = "Finalizar"
            cmdReserva4.Enabled = True
            DTPicker4.Enabled = False
            cboName4.Enabled = False
            Call NewRegister(4)
        End If
    Else
        ' Mensagem de confirmação
        Beep
        value = MsgBox("Finalizar o Treino com: " & cboName4.Text & " ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 2 Then Exit Sub ' Não
        ' Mensagem de confirmação
        Beep
        value = MsgBox("Deseja Imprimir Cupom dos Resultados?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
        If value = 1 Then Call PrintCupom(4) 'Sim
        ' Finalizar
        treino = False
        Call UploadRegisters(4)
        Call EnableCombo(4)
        Call ClearValue(4)
        cmdIniciar4.Caption = "Iniciar"
        cmdIniciar4.BackColor = &H8000000F
        cmdReserva4.BackColor = &H8000000F
        cmdReserva4.Enabled = True
        DTPicker4.Enabled = True
        cboName4.Enabled = True
    End If
    
Exit Sub
Erro:
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' HABILITA RESERVA PARA JOGADOR 4
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdReserva4_Click()
On Error GoTo Erro

    ' Mensagem de confirmação
    Beep
    value = MsgBox("Iniciar o treino com Reserva ?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
    If value = 2 Then Exit Sub ' Não
     ' Mensagem de confirmação
    Beep
    value = MsgBox("Deseja Imprimir Cupom dos Resultados?", vbOKCancel, "DALCOQUIO AUTOMAÇÃO")
    If value = 1 Then Call PrintCupom(4) 'Sim
    ' Habilita Reserva
    treino = False
    Call UploadRegisters(4)
    cmdIniciar4.Caption = "Iniciar"
    cmdIniciar4.BackColor = &H8000000F
    cmdReserva4.BackColor = vbYellow
    DTPicker4.Enabled = True
    cboName4.Enabled = True
    Call DesableCombo(4)
    Call ClearValue(4)
    
Exit Sub
Erro:
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' CHAMADA DAS ATUALIZAÇÕES DE REGISTROS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Combo1_Click(index As Integer)
    Call UploadRegisters(1)
    
End Sub

Private Sub Combo2_Click(index As Integer)
    Call UploadRegisters(2)
    
End Sub

Private Sub Combo3_Click(index As Integer)
    Call UploadRegisters(3)
    
End Sub

Private Sub Combo4_Click(index As Integer)
    Call UploadRegisters(4)
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' SOMA PONTOS DOS TREINOS ATIVOS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub SomaPontos(index As Integer)
On Error GoTo Erro
    
    'JOGADOR 1
    '-------------------------------------------------------------------------------------------------------------------------------
    If index = 1 Then
    
        ' Soma T1
        txtTotal1(1).Text = "0"
        For i = 1 To 5
            If Combo1(i).Text <> Empty Then txtTotal1(1).Text = CInt(txtTotal1(1).Text) + CInt(Combo1(i).Text)
        Next i
        ' Soma T2
        txtTotal1(2).Text = "0"
        For i = 7 To 11
            If Combo1(i).Text <> Empty Then txtTotal1(2).Text = CInt(txtTotal1(2).Text) + CInt(Combo1(i).Text)
        Next i
        'Soma T1 + T2
        txtTotal1(3).Text = CInt(txtTotal1(1).Text) + CInt(txtTotal1(2).Text)
    
        ' Soma T3
        txtTotal1(4).Text = "0"
        For i = 13 To 17
            If Combo1(i).Text <> Empty Then txtTotal1(4).Text = CInt(txtTotal1(4).Text) + CInt(Combo1(i).Text)
        Next i
        ' Soma T4
        txtTotal1(5).Text = "0"
        For i = 19 To 23
            If Combo1(i).Text <> Empty Then txtTotal1(5).Text = CInt(txtTotal1(5).Text) + CInt(Combo1(i).Text)
        Next i
        'Soma T3 + T4
        txtTotal1(6).Text = CInt(txtTotal1(4).Text) + CInt(txtTotal1(5).Text)
    
        'Soma Total Final
        txtTotal1(7).Text = CInt(txtTotal1(3).Text) + CInt(txtTotal1(6).Text)
    
    End If
    

     'JOGADOR 2
     '-------------------------------------------------------------------------------------------------------------------------------
    If index = 2 Then
    
        ' Soma T1
        txtTotal2(1).Text = "0"
        For i = 1 To 5
            If Combo2(i).Text <> Empty Then txtTotal2(1).Text = CInt(txtTotal2(1).Text) + CInt(Combo2(i).Text)
        Next i
        ' Soma T2
        txtTotal2(2).Text = "0"
        For i = 7 To 11
            If Combo2(i).Text <> Empty Then txtTotal2(2).Text = CInt(txtTotal2(2).Text) + CInt(Combo2(i).Text)
        Next i
        'Soma T1 + T2
        txtTotal2(3).Text = CInt(txtTotal2(1).Text) + CInt(txtTotal2(2).Text)
    
        ' Soma T3
        txtTotal2(4).Text = "0"
        For i = 13 To 17
            If Combo2(i).Text <> Empty Then txtTotal2(4).Text = CInt(txtTotal2(4).Text) + CInt(Combo2(i).Text)
        Next i
        ' Soma T4
        txtTotal2(5).Text = "0"
        For i = 19 To 23
            If Combo2(i).Text <> Empty Then txtTotal2(5).Text = CInt(txtTotal2(5).Text) + CInt(Combo2(i).Text)
        Next i
        'Soma T1 + T2
        txtTotal2(6).Text = CInt(txtTotal2(4).Text) + CInt(txtTotal2(5).Text)
    
        'Soma Total Final
        txtTotal2(7).Text = CInt(txtTotal2(3).Text) + CInt(txtTotal2(6).Text)
    
    End If
    

     'JOGADOR 3
     '-------------------------------------------------------------------------------------------------------------------------------
    If index = 3 Then
    
        ' Soma T1
        txtTotal3(1).Text = "0"
        For i = 1 To 5
            If Combo3(i).Text <> Empty Then txtTotal3(1).Text = CInt(txtTotal3(1).Text) + CInt(Combo3(i).Text)
        Next i
        ' Soma T2
        txtTotal3(2).Text = "0"
        For i = 7 To 11
            If Combo3(i).Text <> Empty Then txtTotal3(2).Text = CInt(txtTotal3(2).Text) + CInt(Combo3(i).Text)
        Next i
        'Soma T1 + T2
        txtTotal3(3).Text = CInt(txtTotal3(1).Text) + CInt(txtTotal3(2).Text)
    
        ' Soma T3
        txtTotal3(4).Text = "0"
        For i = 13 To 17
            If Combo3(i).Text <> Empty Then txtTotal3(4).Text = CInt(txtTotal3(4).Text) + CInt(Combo3(i).Text)
        Next i
        ' Soma T4
        txtTotal3(5).Text = "0"
        For i = 19 To 23
            If Combo3(i).Text <> Empty Then txtTotal3(5).Text = CInt(txtTotal3(5).Text) + CInt(Combo3(i).Text)
        Next i
        'Soma T3 + T4
        txtTotal3(6).Text = CInt(txtTotal3(4).Text) + CInt(txtTotal3(5).Text)
    
        'Soma Total Final
        txtTotal3(7).Text = CInt(txtTotal3(3).Text) + CInt(txtTotal3(6).Text)

    End If
    
    
     'JOGADOR 4
     '-------------------------------------------------------------------------------------------------------------------------------
    If index = 4 Then
    
        ' Soma T1
        txtTotal4(1).Text = "0"
        For i = 1 To 5
            If Combo4(i).Text <> Empty Then txtTotal4(1).Text = CInt(txtTotal4(1).Text) + CInt(Combo4(i).Text)
        Next i
        ' Soma T2
        txtTotal4(2).Text = "0"
        For i = 7 To 11
            If Combo4(i).Text <> Empty Then txtTotal4(2).Text = CInt(txtTotal4(2).Text) + CInt(Combo4(i).Text)
        Next i
        'Soma T1 + T2
        txtTotal4(3).Text = CInt(txtTotal4(1).Text) + CInt(txtTotal4(2).Text)
    
        ' Soma T3
        txtTotal4(4).Text = "0"
        For i = 13 To 17
            If Combo4(i).Text <> Empty Then txtTotal4(4).Text = CInt(txtTotal4(4).Text) + CInt(Combo4(i).Text)
        Next i
        ' Soma T4
        txtTotal4(5).Text = "0"
        For i = 19 To 23
            If Combo4(i).Text <> Empty Then txtTotal4(5).Text = CInt(txtTotal4(5).Text) + CInt(Combo4(i).Text)
        Next i
        'Soma T3 + T4
        txtTotal4(6).Text = CInt(txtTotal4(4).Text) + CInt(txtTotal4(5).Text)
    
        'Soma Total Final
        txtTotal4(7).Text = CInt(txtTotal4(3).Text) + CInt(txtTotal4(6).Text)
    
    End If

Exit Sub
Erro:
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' ENVIA ATUALIZAÇÃO DE PONTOS PARA O REGISTRO (UPLOAD)
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub UploadRegisters(index As Integer)
On Error GoTo Erro

    Call SomaPontos(index) ' soma de pontos
    Call TempRegisters(index) ' registros temporários
    
    ' Configurações para Registros
    query = "SELECT * FROM TabelaTreino ORDER by Nome ASC "
    Call queryString(query)
    
    ' Busca critérios nos registros
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset("Nome") = cboName.Text And Adodc1.Recordset("Data") = DTPicker.value Then
            Exit Do ' Localizado
        End If
        Adodc1.Recordset.MoveNext
    Loop
           
    ' Atualiza registro de pontos
    If Adodc1.Recordset.EOF = False Then
        
        If treino = False Then Adodc1.Recordset("HORA-FIM") = Time
        If Combo(0).Text <> Empty Then Adodc1.Recordset("T1-E") = Combo(0).Text
        If Combo(1).Text <> Empty Then Adodc1.Recordset("T1-1") = Combo(1).Text
        If Combo(2).Text <> Empty Then Adodc1.Recordset("T1-2") = Combo(2).Text
        If Combo(3).Text <> Empty Then Adodc1.Recordset("T1-3") = Combo(3).Text
        If Combo(4).Text <> Empty Then Adodc1.Recordset("T1-4") = Combo(4).Text
        If Combo(5).Text <> Empty Then Adodc1.Recordset("T1-5") = Combo(5).Text
        Adodc1.Recordset("T1-Total") = txtTotal(1).Text
        If Combo(6).Text <> Empty Then Adodc1.Recordset("T2-E") = Combo(6).Text
        If Combo(7).Text <> Empty Then Adodc1.Recordset("T2-1") = Combo(7).Text
        If Combo(8).Text <> Empty Then Adodc1.Recordset("T2-2") = Combo(8).Text
        If Combo(9).Text <> Empty Then Adodc1.Recordset("T2-3") = Combo(9).Text
        If Combo(10).Text <> Empty Then Adodc1.Recordset("T2-4") = Combo(10).Text
        If Combo(11).Text <> Empty Then Adodc1.Recordset("T2-5") = Combo(11).Text
        Adodc1.Recordset("T2-Total") = txtTotal(2).Text
        Adodc1.Recordset("T1T2-SubTotal") = txtTotal(3).Text
        If Combo(12).Text <> Empty Then Adodc1.Recordset("T3-E") = Combo(12).Text
        If Combo(13).Text <> Empty Then Adodc1.Recordset("T3-1") = Combo(13).Text
        If Combo(14).Text <> Empty Then Adodc1.Recordset("T3-2") = Combo(14).Text
        If Combo(15).Text <> Empty Then Adodc1.Recordset("T3-3") = Combo(15).Text
        If Combo(16).Text <> Empty Then Adodc1.Recordset("T3-4") = Combo(16).Text
        If Combo(17).Text <> Empty Then Adodc1.Recordset("T3-5") = Combo(17).Text
        Adodc1.Recordset("T3-Total") = txtTotal(4).Text
        If Combo(18).Text <> Empty Then Adodc1.Recordset("T4-E") = Combo(18).Text
        If Combo(19).Text <> Empty Then Adodc1.Recordset("T4-1") = Combo(19).Text
        If Combo(20).Text <> Empty Then Adodc1.Recordset("T4-2") = Combo(20).Text
        If Combo(21).Text <> Empty Then Adodc1.Recordset("T4-3") = Combo(21).Text
        If Combo(22).Text <> Empty Then Adodc1.Recordset("T4-4") = Combo(22).Text
        If Combo(23).Text <> Empty Then Adodc1.Recordset("T4-5") = Combo(23).Text
        Adodc1.Recordset("T4-Total") = txtTotal(5).Text
        Adodc1.Recordset("T3T4-SubTotal") = txtTotal(6).Text
        Adodc1.Recordset("TotalFinal") = txtTotal(7).Text
        Adodc1.Recordset.Update
   
    End If
    
        If tmrDownload.Enabled = False Then
            Dim i As Integer
            For i = 1 To 4
                DownloadRegisters (i)
                Call SomaPontos(i) ' soma de pontos
            Next i
        
        End If
    
Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' BUSCA ATUALIZAÇÃO DE PONTOS ATUAIS NO REGISTRO (DOWNLOAD)
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub DownloadRegisters(index As Integer)
On Error GoTo Erro
    
    Call TempRegisters(index) ' registros temporários
    If TempRegisters(index) = 0 Then Exit Sub

    ' Configurações para Registros
    query = "SELECT * FROM TabelaTreino ORDER by Nome ASC "
    Call queryString(query)
    
    ' Busca critérios nos registros
    Do While Not Adodc1.Recordset.EOF
            If Adodc1.Recordset("NOME") = cboName.Text And Adodc1.Recordset("DATA") = DTPicker.value Then
                If IsNull(Adodc1.Recordset("HORA-FIM")) Then Exit Do  ' Localizado
            End If
        Adodc1.Recordset.MoveNext
    Loop
           
    ' Atualiza pontos no treino atual
    If Adodc1.Recordset.EOF = False Then
        If Not Adodc1.Recordset("T1-E") Then Combo(0).Text = Adodc1.Recordset("T1-E")
        If Not Adodc1.Recordset("T1-1") Then Combo(1).Text = Adodc1.Recordset("T1-1")
        If Not Adodc1.Recordset("T1-2") Then Combo(2).Text = Adodc1.Recordset("T1-2")
        If Not Adodc1.Recordset("T1-3") Then Combo(3).Text = Adodc1.Recordset("T1-3")
        If Not Adodc1.Recordset("T1-4") Then Combo(4).Text = Adodc1.Recordset("T1-4")
        If Not Adodc1.Recordset("T1-5") Then Combo(5).Text = Adodc1.Recordset("T1-5")
        If Not Adodc1.Recordset("T1-Total") Then txtTotal(1).Text = Adodc1.Recordset("T1-Total")
        If Not Adodc1.Recordset("T2-E") Then Combo(6).Text = Adodc1.Recordset("T2-E")
        If Not Adodc1.Recordset("T2-1") Then Combo(7).Text = Adodc1.Recordset("T2-1")
        If Not Adodc1.Recordset("T2-2") Then Combo(8).Text = Adodc1.Recordset("T2-2")
        If Not Adodc1.Recordset("T2-3") Then Combo(9).Text = Adodc1.Recordset("T2-3")
        If Not Adodc1.Recordset("T2-4") Then Combo(10).Text = Adodc1.Recordset("T2-4")
        If Not Adodc1.Recordset("T2-5") Then Combo(11).Text = Adodc1.Recordset("T2-5")
        If Not Adodc1.Recordset("T2-Total") Then txtTotal(2).Text = Adodc1.Recordset("T2-Total")
        If Not Adodc1.Recordset("T1T2-SubTotal") Then txtTotal(3).Text = Adodc1.Recordset("T1T2-SubTotal")
        If Not Adodc1.Recordset("T3-E") Then Combo(12).Text = Adodc1.Recordset("T3-E")
        If Not Adodc1.Recordset("T3-1") Then Combo(13).Text = Adodc1.Recordset("T3-1")
        If Not Adodc1.Recordset("T3-2") Then Combo(14).Text = Adodc1.Recordset("T3-2")
        If Not Adodc1.Recordset("T3-3") Then Combo(15).Text = Adodc1.Recordset("T3-3")
        If Not Adodc1.Recordset("T3-4") Then Combo(16).Text = Adodc1.Recordset("T3-4")
        If Not Adodc1.Recordset("T3-5") Then Combo(17).Text = Adodc1.Recordset("T3-5")
        If Not Adodc1.Recordset("T3-Total") Then txtTotal(4).Text = Adodc1.Recordset("T3-Total")
        If Not Adodc1.Recordset("T4-E") Then Combo(18).Text = Adodc1.Recordset("T4-E")
        If Not Adodc1.Recordset("T4-1") Then Combo(19).Text = Adodc1.Recordset("T4-1")
        If Not Adodc1.Recordset("T4-2") Then Combo(20).Text = Adodc1.Recordset("T4-2")
        If Not Adodc1.Recordset("T4-3") Then Combo(21).Text = Adodc1.Recordset("T4-3")
        If Not Adodc1.Recordset("T4-4") Then Combo(22).Text = Adodc1.Recordset("T4-4")
        If Not Adodc1.Recordset("T4-5") Then Combo(23).Text = Adodc1.Recordset("T4-5")
        If Not Adodc1.Recordset("T4-Total") Then txtTotal(5).Text = Adodc1.Recordset("T4-Total")
        If Not Adodc1.Recordset("T3T4-SubTotal") Then txtTotal(6).Text = Adodc1.Recordset("T3T4-SubTotal")
        If Not Adodc1.Recordset("TotalFinal") Then txtTotal(7).Text = Adodc1.Recordset("TotalFinal")
    End If

Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' BUSCA ATUALIZAÇÃO AUTOMÁTICA DE PONTOS (DOWNLOAD)
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub tmrDownload_Timer()
    If treino = True Then
        Dim i As Integer
        For i = 1 To 4
            'DownloadRegisters (i)
        Next i
    End If
    
End Sub

' Somente para testes de download
Private Sub cmdDownload_Click(index As Integer)
    Dim i As Integer
    For i = 1 To 4
        DownloadRegisters (i)
        Call SomaPontos(i) ' soma de pontos
    Next i
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' ENVIA INICIO DE NOVO TREINO PARA O REGISTRO (UPLOAD)
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub NewRegister(index As Integer)
On Error GoTo Erro

    ' Configurações para Registros
    query = "SELECT * FROM TabelaTreino ORDER by Nome ASC "
    Call queryString(query)

    ' Jogador 1
    If index = 1 Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset("Data") = DTPicker1.value
        Adodc1.Recordset("HORA") = Time
        Adodc1.Recordset("Nome") = cboName1.Text
        Adodc1.Recordset.Update
    End If
    
    ' Jogador 2
    If index = 2 Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset("Data") = DTPicker2.value
        Adodc1.Recordset("HORA") = Time
        Adodc1.Recordset("Nome") = cboName2.Text
        Adodc1.Recordset.Update
    End If
    
    ' Jogador 3
    If index = 3 Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset("Data") = DTPicker3.value
        Adodc1.Recordset("HORA") = Time
        Adodc1.Recordset("Nome") = cboName3.Text
        Adodc1.Recordset.Update
    End If
    
    ' Jogador 4
    If index = 4 Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset("Data") = DTPicker4.value
        Adodc1.Recordset("HORA") = Time
        Adodc1.Recordset("Nome") = cboName4.Text
        Adodc1.Recordset.Update
    End If
    
Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' REGISTROS TEMPORÁRIOS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Function TempRegisters(index As Integer) As Integer
    Dim i As Integer

    'Jogador 1
    If index = 1 Then
        If cboName1.Text = Empty Then
            TempRegisters = 0
            Exit Function
        Else
            TempRegisters = index
        End If
        Set DTPicker = DTPicker1
        DTPicker.value = DTPicker1.value
        Set cboName = cboName1
        cboName.Text = cboName1.Text
        For i = 1 To 7
            Set txtTotal(i) = txtTotal1(i)
            txtTotal(i).Text = txtTotal1(i).Text
        Next i
        For i = 0 To 23
            Set Combo(i) = Combo1(i)
            Combo(i).Text = Combo1(i).Text
        Next i
    End If

    'Jogador 2
    If index = 2 Then
        If cboName2.Text = Empty Then
            TempRegisters = 0
            Exit Function
        Else
            TempRegisters = index
        End If
        Set DTPicker = DTPicker2
        DTPicker.value = DTPicker2.value
        Set cboName = cboName2
        cboName.Text = cboName2.Text
        For i = 1 To 7
            Set txtTotal(i) = txtTotal2(i)
            txtTotal(i).Text = txtTotal2(i).Text
        Next i
        For i = 0 To 23
            Set Combo(i) = Combo2(i)
            Combo(i).Text = Combo2(i).Text
        Next i
    End If

    'Jogador 3
    If index = 3 Then
        If cboName3.Text = Empty Then
            TempRegisters = 0
            Exit Function
        Else
            TempRegisters = index
        End If
        Set DTPicker = DTPicker3
        DTPicker.value = DTPicker3.value
        Set cboName = cboName3
        cboName.Text = cboName3.Text
        For i = 1 To 7
            Set txtTotal(i) = txtTotal3(i)
            txtTotal(i).Text = txtTotal3(i).Text
        Next i
        For i = 0 To 23
            Set Combo(i) = Combo3(i)
            Combo(i).Text = Combo3(i).Text
        Next i
    End If

    'Jogador 4
    If index = 4 Then
        If cboName4.Text = Empty Then
            TempRegisters = 0
            Exit Function
        Else
            TempRegisters = index
        End If
        Set DTPicker = DTPicker4
        DTPicker.value = DTPicker4.value
        Set cboName = cboName4
        cboName.Text = cboName4.Text
        For i = 1 To 7
            Set txtTotal(i) = txtTotal4(i)
            txtTotal(i).Text = txtTotal4(i).Text
        Next i
        For i = 0 To 23
            Set Combo(i) = Combo4(i)
            Combo(i).Text = Combo4(i).Text
        Next i
    End If

End Function

'//////////////////////////////////////////////////////////////////////////////////////////////
' HABILITA VISIBILIDADE COMBOS DE PONTOS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub EnableCombo(index As Integer)
    Dim index2 As Integer
    
    If index = 1 Then
        For i = 0 To 23
            Combo1(i).Visible = True
        Next
        Combo1(1).SetFocus
    End If
    
    If index = 2 Then
        For i = 0 To 23
            Combo2(i).Visible = True
        Next
        Combo2(0).SetFocus
    End If
    
    If index = 3 Then
        For i = 0 To 23
            Combo3(i).Visible = True
        Next
        Combo3(0).SetFocus
    End If
    
    If index = 1 Then
        For i = 0 To 23
            Combo1(i).Visible = True
        Next
        Combo1(0).SetFocus
    End If

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' DESABILITA VISIBILIDADE COMBOS DE PONTOS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub DesableCombo(index As Integer)
    Dim index2 As Integer
    
    If index = 1 Then
        For i = 0 To 23
            If Combo1(i).Text <> Empty Then
                Combo1(i).Visible = False
                index2 = i
            End If
        Next
        Combo1(index2 + 1).SetFocus
    End If

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' ATUALIZA OS COMBOS DOS NOMES
'//////////////////////////////////////////////////////////////////////////////////////////////

Public Sub updateCboNames()
    On Error GoTo Erro
    
    ' Configurações para Registros
    query = "SELECT * FROM TabelaCadastro ORDER by Nome ASC "
    Call queryString(query)
    
    ' Clear Combos
    cboName1.Clear
    cboName2.Clear
    cboName3.Clear
    cboName4.Clear
    
    ' Atualiza lista
    Do While Not Adodc1.Recordset.EOF
        cboName1.AddItem Adodc1.Recordset("NOME")
        cboName2.AddItem Adodc1.Recordset("NOME")
        cboName3.AddItem Adodc1.Recordset("NOME")
        cboName4.AddItem Adodc1.Recordset("NOME")
        Adodc1.Recordset.MoveNext
    Loop
    
    ' Fecha conexão com o registro
    'Adodc1.Recordset.Close
    
Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' LIMPA VALORES DE PONTOS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub ClearValue(index As Integer)
On Error GoTo Erro

    Dim i  As Integer
    
    If index = 1 Then
        For i = 0 To 23
            Combo1(i).Text = Empty
        Next i
        For i = 1 To 7
            txtTotal1(i).Text = "0"
        Next i
    End If
    
    If index = 2 Then
        For i = 0 To 23
            Combo2(i).Text = Empty
        Next i
        For i = 1 To 7
            txtTotal2(i).Text = "0"
        Next i
    End If
    
    If index = 3 Then
        For i = 0 To 23
            Combo3(i).Text = Empty
        Next i
        For i = 1 To 7
            txtTotal3(i).Text = "0"
        Next i
    End If
    
    If index = 4 Then
        For i = 0 To 23
            Combo4(i).Text = Empty
        Next i
        For i = 1 To 7
            txtTotal4(i).Text = "0"
        Next i
    End If
    
    ptrCupom.Cls ' limpa cupom de treino
    
Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' IMPRESSÃO DO CUPOM DOS RESULTADOS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub PrintCupom(index As Integer)
On Error GoTo Erro

    Call TempRegisters(index)
    
    ptrCupom.AutoRedraw = True
    ptrCupom.Cls
    
    Fonte 10, False, False
    ptrCupom.Print String(45, " ") 'Pula uma Linha
    ptrCupom.Print String(45, "-") 'Faz uma Linha
    ptrCupom.Print "Data: " & DTPicker.value
    ptrCupom.Print String(45, "-") 'Faz uma Linha
    ptrCupom.Print "Nome: " & cboName.Text
    ptrCupom.Print String(45, " ") 'Pula uma Linha
    ptrCupom.Print "ToTal 1: " & txtTotal(1).Text
    ptrCupom.Print "ToTal 2: " & txtTotal(2).Text
    ptrCupom.Print "Sub Total: " & txtTotal(3).Text
    ptrCupom.Print "ToTal 3: " & txtTotal(4).Text
    ptrCupom.Print "ToTal 4: " & txtTotal(5).Text
    ptrCupom.Print "Sub Total: " & txtTotal(6).Text
    Fonte 12, True, False
    ptrCupom.Print "Total Final: " & txtTotal(7).Text
    Fonte 10, False, False
    ptrCupom.Print String(45, " ") 'Pula uma Linha
    ptrCupom.Print String(45, "-") 'Faz uma Linha
    ptrCupom.Print Tab(1); "DALCOQUIO AUTOMACAO"
    ptrCupom.Print String(45, "-") 'Faz uma Linha
    ptrCupom.Print String(45, " ") 'Pula uma Linha
    
    ' Imprimir
    CommonDialog1.ShowPrinter
    'Printer.PaintPicture ptrCupom.Image, PositionX, PosicionY, Width, Height
    Printer.PaintPicture ptrCupom.Image, 0, 0
    Printer.EndDoc
    ptrCupom.Cls
    
Exit Sub

Erro:
    Beep
    If Err.Number = 482 Then
        MsgBox "Processo cancelado !!!", vbExclamation, "DALCOQUIO AUTOMAÇÃO"
        ptrCupom.Cls
    Else
        MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
        'MsgBox Err.number, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    End If

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' FONTES PARA O CUPOM
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Fonte(Tamanho As Byte, Negrito As Boolean, Italico As Boolean) 'Altera a fonte
    ptrCupom.FontSize = Tamanho
    ptrCupom.FontBold = Negrito
    ptrCupom.FontItalic = Italico
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' TRATAMENTO DE TODOS OS MENUS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub mnuCadastro_Click()
    frmCadastro.Show

End Sub

Private Sub mnuTreinos_Click()
    frmRegistros.Show
    
End Sub

Private Sub mnuControle_Click()
    frmControle.Show vbModal

End Sub

Private Sub mnuAddressBD_Click()
    value = InputBox("Digite o Endereço do Banco de Dados.", "DALÇÓQUIO AUTOMAÇÃO", addressBackups)
    If value <> Empty Then
        addressRegisters = value
        WriteIniValue App.Path & "\Config.ini", "REGISTROS", "addressBD", addressRegisters
        MsgBox "Fechando sistema para atualização do endereço do bando de dados, você deverá abri-lo novamente...", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
        End ' Fecha o sistema
    End If

End Sub

Private Sub mnuAddressBKP_Click()
    value = InputBox("Digite o Endereço de Backup.", "DALÇÓQUIO AUTOMAÇÃO", addressRegisters)
    If value <> Empty Then
        addressBackups = value
        WriteIniValue App.Path & "\Config.ini", "REGISTROS", "addressBKP", addressBackups
        MsgBox "Fechando sistema para atualização do endereço de Backup, você deverá abri-lo novamente...", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
        End ' Fecha o sistema
    End If
End Sub

Private Sub mnuTestBD_Click()
On Error GoTo Erro
    
    Adodc1.ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};" & "Dbq= " & frmTreino.addressRegisters & nameRegisters & ";" & "Uid=;"  ' Pwd=1234"
    Adodc1.Refresh
    
    If Adodc1.Recordset.State <> adStateClosed Then
        MsgBox "Teste de conexão com o banco de dados efetuada com sucesso...", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
    End If
    
Exit Sub
Erro:
    MsgBox "Falha no teste de conexão com o banco de bados", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' TIMER RELOGIO - DATA/HORA
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub tmrRelogio_Timer()
    lblRelogio.Caption = Now
    
End Sub




