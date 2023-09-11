VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRegistros 
   ClientHeight    =   11520
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   28365
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   28365
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ptrCupom 
      Appearance      =   0  'Flat
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
      Height          =   8655
      Left            =   23040
      ScaleHeight     =   8625
      ScaleWidth      =   4965
      TabIndex        =   12
      Top             =   960
      Width           =   5000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   10560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   9720
      Width           =   22695
      Begin VB.CommandButton cmdDownload 
         Caption         =   "Download Registros"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   15960
         TabIndex        =   14
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmdPrintCupom 
         Caption         =   "Imprimir para Cupom"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmdExportCsv 
         Caption         =   "Exportar para CSV"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12600
         TabIndex        =   11
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar (Desativado)"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   19320
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmdPrintImpressora 
         Caption         =   "Imprimir para Impressora"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9240
         TabIndex        =   9
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   22695
      Begin VB.CommandButton cmdPeriodo 
         Caption         =   "Período"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   18000
         TabIndex        =   6
         Top             =   240
         Width           =   4600
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13320
         TabIndex        =   5
         Top             =   240
         Width           =   4600
      End
      Begin VB.ComboBox cboName 
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
         ItemData        =   "Form3.frx":0000
         Left            =   120
         List            =   "Form3.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   5115
      End
      Begin VB.CommandButton cmdNome 
         Caption         =   " Nome"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8640
         TabIndex        =   1
         Top             =   240
         Width           =   4600
      End
      Begin MSComCtl2.DTPicker dtpInicial 
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54394881
         CurrentDate     =   45156
      End
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   315
         Left            =   6960
         TabIndex        =   4
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   54394881
         CurrentDate     =   45156
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   11040
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
      Connect         =   $"Form3.frx":0004
      OLEDBString     =   $"Form3.frx":00C9
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":018E
      Height          =   8655
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   22695
      _ExtentX        =   40031
      _ExtentY        =   15266
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   36
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DATA"
         Caption         =   "DATA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "HORA"
         Caption         =   "HORA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "NOME"
         Caption         =   "NOME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "T1-E"
         Caption         =   "E"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "T1-1"
         Caption         =   "1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "T1-2"
         Caption         =   "2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "T1-3"
         Caption         =   "3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "T1-4"
         Caption         =   "4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "T1-5"
         Caption         =   "5"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "T1-Total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "T2-E"
         Caption         =   "E"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "T2-1"
         Caption         =   "1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "T2-2"
         Caption         =   "2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "T2-3"
         Caption         =   "3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "T2-4"
         Caption         =   "4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "T2-5"
         Caption         =   "5"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column17 
         DataField       =   "T2-Total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "T1T2-SubTotal"
         Caption         =   "SubTotal"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column19 
         DataField       =   "T3-1"
         Caption         =   "E"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "T3-1"
         Caption         =   "1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column21 
         DataField       =   "T3-2"
         Caption         =   "2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "T3-4"
         Caption         =   "3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column23 
         DataField       =   "T3-4"
         Caption         =   "4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column24 
         DataField       =   "T3-5"
         Caption         =   "5"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column25 
         DataField       =   "T3-Total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column26 
         DataField       =   "T4-E"
         Caption         =   "E"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column27 
         DataField       =   "T4-1"
         Caption         =   "1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column28 
         DataField       =   "T4-2"
         Caption         =   "2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column29 
         DataField       =   "T4-3"
         Caption         =   "3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column30 
         DataField       =   "T4-4"
         Caption         =   "4"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column31 
         DataField       =   "T4-5"
         Caption         =   "5"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column32 
         DataField       =   "T4-Total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column33 
         DataField       =   "T3T4-SubTotal"
         Caption         =   "SubTotal"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column34 
         DataField       =   "TotalFinal"
         Caption         =   "TotalFinal"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column35 
         DataField       =   "HORA-FIM"
         Caption         =   "Hora-Fim"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   4004,788
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   404,787
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   1005,165
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim query As String

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
' INICIO DO FORMULÁRIO REGISTROS ( PRINCIPAL )
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Form_Load()
    ' Atualização inicial
    Call ReadNameRegisters
    Call ReadAllRegisters
    
    ' Atualiza data atual
    dtpInicial.value = Date
    dtpFinal.value = Date
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' CRITÉRIO DE BUSCA POR NOME
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdNome_Click()
On Error GoTo Erro
    ' Verifica se nome em branco
    If cboName.Text = Empty Then
        cboName.BackColor = vbYellow
        MsgBox "Nenhum nome selecionado.", vbInformation, "DALCOQUIO AUTOMAÇÃO"
        cboName.BackColor = vbWhite
        Exit Sub
    End If
    
    ' Configurações para Registros
    query = "SELECT * FROM TabelaTreino WHERE Nome LIKE '%" & cboName.Text & "%'"
    Call queryString(query)
    
    ' Atualiza DataGrid1
    Set DataGrid1.DataSource = Adodc1.Recordset

Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' CRITÉRIO DE BUSCA POR DATA
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdData_Click()
On Error GoTo Erro
    ' Verifica se nome em branco
    If cboName.Text = Empty Then
        cboName.BackColor = vbYellow
        MsgBox "Nenhum nome selecionado.", vbInformation, "DALCOQUIO AUTOMAÇÃO"
        cboName.BackColor = vbWhite
        Exit Sub
    End If
    
    ' Configurações para Registros
    query = "SELECT * FROM TabelaTreino WHERE Nome LIKE '%" & cboName.Text & "%' AND Data = #" & dtpInicial.value & "#"
    Call queryString(query)
    
    ' Atualzia DataGrid1
    Set DataGrid1.DataSource = Adodc1.Recordset

Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' CRITÉRIO DE BUSCA POR PERÍODO
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdPeriodo_Click()
On Error GoTo Erro
    ' Verifica se nome em branco
    If cboName.Text = Empty Then
        cboName.BackColor = vbYellow
        MsgBox "Nenhum nome selecionado.", vbInformation, "DALCOQUIO AUTOMAÇÃO"
        cboName.BackColor = vbWhite
        Exit Sub
    End If
    
    ' Configurações para Registros
    query = "SELECT * FROM TabelaTreino WHERE Nome LIKE '%" & cboName.Text & "%' AND Data >= #" & dtpInicial.value & "# AND Data <= #" & dtpFinal.value & "#"
    Call queryString(query)
    
   ' Atualzia DataGrid1
    Set DataGrid1.DataSource = Adodc1.Recordset

Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' COMANDO PARA O BOTÃO EDITAR
' Ativa e Desativa a Lista de Registros para Edição
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdEditar_Click()
    If cmdEditar.Caption = "Editar (Desativado)" Then
        DataGrid1.AllowDelete = True
        DataGrid1.AllowUpdate = True
        cmdEditar.Caption = "Editar (Ativado)"
        cmdEditar.BackColor = vbYellow
    Else
        DataGrid1.AllowDelete = False
        DataGrid1.AllowUpdate = False
        cmdEditar.Caption = "Editar (Desativado)"
        cmdEditar.BackColor = &H8000000F
    End If
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' CRITÉRIO DE BUSCA POR REGISTRO DE NOME
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub ReadNameRegisters()
On Error GoTo Erro

    ' Configurações para Registros
    query = "SELECT * FROM TabelaCadastro ORDER by Nome ASC "
    Call queryString(query)
    
    ' Atualiza lista
    cboName.Clear
    Do While Not Adodc1.Recordset.EOF
        cboName.AddItem Adodc1.Recordset("NOME")
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
' CRITÉRIO DE BUSCA POR TODOS OS REGISTROS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub ReadAllRegisters()
On Error GoTo Erro
    
    ' Configurações para Registros
    query = "SELECT * FROM TabelaTreino ORDER by ID ASC "
    Call queryString(query)
    
    ' Atualzia DataGrid1
    Set DataGrid1.DataSource = Adodc1.Recordset
    
Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' ATUALIZA LISTA DE REGISTROS (REFRESH)
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdDownload_Click()
    Call ReadAllRegisters
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' PRINT SCREEN
' Cria um print screen do form de registros
' Impressora deverá ser configurada para Folha A3 e Modo Paisagem
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdPrintImpressora_Click()
    frmPrint.Show
        
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' IMPRESSÃO PARA CUPOM DOS RESULTADOS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdPrintCupom_Click()

On Error GoTo Erro

    ' Verifica nome selecionado na lista
    Dim selectedID As String
    selectedID = DataGrid1.Columns(3).Text

     ' Se nenhum nome selecionada na lista
    If selectedID = Empty Then
        MsgBox "Nenhum nome selecionado na lista.", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
        Exit Sub
    End If
    
    ' Configurações para Registros
    query = "SELECT * FROM TabelaCadastro ORDER by Nome ASC "
    Call queryString(query)
    
     ' Busca no registro nome selecionado na lista
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset("NOME") = selectedID Then
            Exit Do ' Localizado
        End If
        Adodc1.Recordset.MoveNext
    Loop
    
    If Adodc1.Recordset.EOF = True Then
        MsgBox "Nome não encontrado !!!", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
        Exit Sub
    Else
    
        ' Busca os registros
        Dim ptrData As String
        Dim ptrNome As String
        Dim ptrT1 As Integer
        Dim ptrT2 As Integer
        Dim ptrT1T2 As Integer
        Dim ptrT3 As Integer
        Dim ptrT4 As Integer
        Dim ptrT3T4 As Integer
        Dim ptrTotal As Integer
         
        ' Critério de busca
        query = "SELECT * FROM TabelaTreino WHERE NOME LIKE '%" & selectedID & "%'"
        Call queryString(query)
        If Not Adodc1.Recordset.EOF Then
            ptrData = Adodc1.Recordset.Fields("DATA")
            ptrNome = Adodc1.Recordset.Fields("NOME")
            ptrT1 = Adodc1.Recordset.Fields("T1-Total")
            ptrT2 = Adodc1.Recordset.Fields("T2-Total")
            ptrT1T2 = Adodc1.Recordset.Fields("T1T2-SubTotal")
            ptrT3 = Adodc1.Recordset.Fields("T3-Total")
            ptrT4 = Adodc1.Recordset.Fields("T4-Total")
            ptrT3T4 = Adodc1.Recordset.Fields("T3T4-SubTotal")
            ptrTotal = Adodc1.Recordset.Fields("TotalFinal")
        
            ' Monta Cupom
            ptrCupom.AutoRedraw = True
            ptrCupom.Cls
            
            Fonte 10, False, False
            ptrCupom.Print String(45, " ") 'Pula uma Linha
            ptrCupom.Print String(45, "-") 'Faz uma Linha
            ptrCupom.Print "Data: " & ptrData
            ptrCupom.Print String(45, "-") 'Faz uma Linha
            ptrCupom.Print "Nome: " & ptrNome
            ptrCupom.Print String(45, " ") 'Pula uma Linha
            ptrCupom.Print "ToTal 1: " & ptrT1
            ptrCupom.Print "ToTal 2: " & ptrT2
            ptrCupom.Print "Sub Total: " & ptrT1T2
            ptrCupom.Print "ToTal 3: " & ptrT3
            ptrCupom.Print "ToTal 4: " & ptrT4
            ptrCupom.Print "Sub Total: " & ptrT3T4
            Fonte 12, True, False
            ptrCupom.Print "Total Final: " & ptrTotal
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
        
        End If
    
    End If
    
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
' ABRE CAIXA DE DIÁLOGO PARA SALVAR REGISTROS EM ARQUIVO.CSV
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdExportCsv_Click()
    On Error GoTo Erro
    Dim i As Long
    Dim x As Long
    Dim Cols As Integer
    Dim sLine As String
    Dim filePath As String

    ' Abre o diálogo de "Salvar como"
    CommonDialog1.CancelError = True
    CommonDialog1.DialogTitle = "Salvar Arquivo CSV"
    CommonDialog1.Filter = "Arquivos CSV (*.csv)|*.csv|Todos os arquivos (*.*)|*.*"

    CommonDialog1.ShowSave

    If Len(CommonDialog1.FileName) > 0 Then ' Verifica se o usuário selecionou "Salvar"
        filePath = CommonDialog1.FileName

        Open filePath For Output As #1
        Cols = DataGrid1.Columns.Count
        
        ' Escreva os nomes das colunas no arquivo CSV
        Dim colNames As String
        For i = 0 To Cols - 1
            colNames = colNames & DataGrid1.Columns(i).Caption & IIf(i < Cols - 1, ", ", "")
        Next i
        Print #1, colNames

        ' Escreva os dados das linhas no arquivo CSV
        For x = 0 To DataGrid1.VisibleRows - 1
            DataGrid1.Row = x
            sLine = ""
            For i = 0 To Cols - 1
                DataGrid1.Col = i
                sLine = sLine & DataGrid1.Text & IIf(i < Cols - 1, ", ", "")
            Next i
            Print #1, sLine
        Next x

        Close #1

        Beep
        MsgBox "Arquivo.csv exportado com sucesso...", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
    End If

Exit Sub

Erro:
    ' Sem tratamento de erro

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' SALVA ARQUIVO.CSV NA PASTA ATUAL DO PROJETO
'//////////////////////////////////////////////////////////////////////////////////////////////

'Private Sub cmdExportCsv_Click()
'    On Error GoTo Erro
'    Dim i As Long
'    Dim x As Long
'    Dim Cols As Integer
'    Dim sLine As String
'
'    Open App.Path & "\Arquivo.csv" For Output As #1
'    Cols = DataGrid1.Columns.Count
'
'    For x = 0 To DataGrid1.VisibleRows - 1
'        DataGrid1.Row = x
'        sLine = ""
'        For i = 0 To Cols - 1 ' Note que os índices das colunas começam em 0
'            DataGrid1.Col = i ' Defina a coluna atual
'            sLine = sLine & DataGrid1.Text & IIf(i < Cols - 1, ", ", "")
'        Next i
'        Print #1, sLine
'    Next x
'
'    Close #1
'    Beep
'    MsgBox "Arquivo.csv exportado com sucesso...", vbInformation, "DALÇÓQUIO AUTOMAÇÃO"
'
'Exit Sub
'
'Erro:
'    ' Sem tratamento de erro
'
'End Sub

