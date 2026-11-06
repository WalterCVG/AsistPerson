VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ProgAdmin 
   Caption         =   "Programar"
   ClientHeight    =   10515
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15075
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "ProgAdmin.frx":0000
   ScaleHeight     =   10515
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdServidor 
      Caption         =   "Prog. Servidor"
      Height          =   495
      Left            =   12360
      TabIndex        =   180
      ToolTipText     =   "Recalculo Servidor"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton CmdLIMPIA 
      Caption         =   "Limpia"
      Height          =   615
      Left            =   14160
      Picture         =   "ProgAdmin.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   179
      ToolTipText     =   "Limpia Plantilla"
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox Txfecnac 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   178
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxGrpGen 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10920
      MaxLength       =   4
      TabIndex        =   176
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdClear 
      Height          =   615
      Left            =   14280
      Picture         =   "ProgAdmin.frx":80798
      Style           =   1  'Graphical
      TabIndex        =   175
      ToolTipText     =   "Limpia Seleccion"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Txdscccosto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   174
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   173
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Txbuscar 
      Height          =   285
      Left            =   10080
      MaxLength       =   5
      TabIndex        =   168
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Txbuscar1 
      Height          =   285
      Left            =   12000
      MaxLength       =   12
      TabIndex        =   167
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton CmdPrint 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   13200
      Picture         =   "ProgAdmin.frx":81062
      Style           =   1  'Graphical
      TabIndex        =   164
      ToolTipText     =   "Prog. Mensual"
      Top             =   9000
      Width           =   645
   End
   Begin VB.ComboBox Cbo_ccosto 
      Height          =   315
      Left            =   6480
      TabIndex        =   162
      Top             =   240
      Width           =   3855
   End
   Begin VB.TextBox Tx_ccosto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10440
      MaxLength       =   2
      TabIndex        =   161
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Turnos"
      Height          =   615
      Left            =   240
      TabIndex        =   160
      Top             =   8760
      Width           =   2535
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "A:Diurno       *:Libre"
         Height          =   255
         Left            =   480
         TabIndex        =   165
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.CommandButton CmdPrograma 
      BackColor       =   &H00C0C0FF&
      Caption         =   "PROGRAMA Todos/Und.Org"
      Height          =   615
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   159
      ToolTipText     =   "Todos / Und.Org."
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox TxCargo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10320
      MaxLength       =   4
      TabIndex        =   158
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxPlH 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   157
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox Txccosto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9120
      MaxLength       =   2
      TabIndex        =   156
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Txtrabajador 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      MaxLength       =   35
      TabIndex        =   4
      Top             =   5520
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "ProgAdmin.frx":8192C
      Style           =   1  'Graphical
      TabIndex        =   155
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox TxTurno 
      Enabled         =   0   'False
      Height          =   195
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   150
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboTurDir 
      Height          =   315
      Left            =   4080
      TabIndex        =   151
      Top             =   9120
      Width           =   4695
   End
   Begin VB.TextBox TxDscCargo 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   35
      TabIndex        =   6
      Top             =   5520
      Width           =   3135
   End
   Begin VB.TextBox TxCarEst 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9960
      MaxLength       =   2
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   14160
      Picture         =   "ProgAdmin.frx":8256E
      Style           =   1  'Graphical
      TabIndex        =   148
      ToolTipText     =   "Cerrar"
      Top             =   9720
      Width           =   735
   End
   Begin VB.CommandButton mnuFicImpSelec 
      Height          =   495
      Left            =   12360
      Picture         =   "ProgAdmin.frx":829B0
      Style           =   1  'Graphical
      TabIndex        =   147
      ToolTipText     =   "Selec.Impresora"
      Top             =   9000
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   9720
      TabIndex        =   146
      Top             =   8760
      Width           =   1935
      Begin VB.CommandButton cmdElimina 
         Caption         =   "Elimina"
         Height          =   615
         Left            =   1080
         Picture         =   "ProgAdmin.frx":8318E
         Style           =   1  'Graphical
         TabIndex        =   172
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdGraba 
         Caption         =   "Graba"
         Height          =   615
         Left            =   120
         Picture         =   "ProgAdmin.frx":83300
         Style           =   1  'Graphical
         TabIndex        =   171
         ToolTipText     =   "Grabar"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox TxHorario 
      Enabled         =   0   'False
      Height          =   195
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   152
      Top             =   9720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox CboHorario 
      Height          =   315
      Left            =   4080
      TabIndex        =   154
      Top             =   9480
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   360
      TabIndex        =   9
      Top             =   5880
      Width           =   13575
      Begin VB.TextBox TxMT 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9840
         TabIndex        =   181
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox Txt 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox TxM 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   130
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox TxN 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10560
         Locked          =   -1  'True
         TabIndex        =   128
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox TxD 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   126
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   30
         Left            =   10080
         MaxLength       =   4
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   1800
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   29
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   118
         Text            =   "Text1"
         Top             =   1800
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   28
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   114
         Text            =   "Text1"
         Top             =   1800
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   27
         Left            =   6480
         MaxLength       =   4
         TabIndex        =   110
         Text            =   "Text1"
         Top             =   1800
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   26
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   106
         Text            =   "Text1"
         Top             =   1800
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   25
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   144
         Text            =   "Text1"
         Top             =   1800
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   24
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   140
         Text            =   "Text1"
         Top             =   1800
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   23
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   136
         Text            =   "Text1"
         Top             =   1800
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   22
         Left            =   480
         MaxLength       =   4
         TabIndex        =   100
         Text            =   "Text1"
         Top             =   1800
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   21
         Left            =   12480
         MaxLength       =   4
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   20
         Left            =   11280
         MaxLength       =   4
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   19
         Left            =   10080
         MaxLength       =   4
         TabIndex        =   88
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   18
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   80
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   6480
         MaxLength       =   4
         TabIndex        =   76
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   64
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   60
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   480
         MaxLength       =   4
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   1200
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   12480
         MaxLength       =   4
         TabIndex        =   52
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   11280
         MaxLength       =   4
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   10080
         MaxLength       =   4
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   6480
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   480
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.TextBox TxHoras 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   134
         Top             =   2400
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   30
         Left            =   9720
         TabIndex        =   120
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   29
         Left            =   8520
         TabIndex        =   116
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   28
         Left            =   7320
         TabIndex        =   112
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   27
         Left            =   6120
         TabIndex        =   108
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   26
         Left            =   4920
         TabIndex        =   104
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   25
         Left            =   3720
         TabIndex        =   142
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   24
         Left            =   2520
         TabIndex        =   138
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   23
         Left            =   1320
         TabIndex        =   102
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   22
         Left            =   120
         TabIndex        =   98
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   21
         Left            =   12120
         TabIndex        =   94
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   20
         Left            =   10920
         TabIndex        =   90
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   19
         Left            =   9720
         TabIndex        =   86
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   18
         Left            =   8520
         TabIndex        =   82
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   17
         Left            =   7320
         TabIndex        =   78
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   16
         Left            =   6120
         TabIndex        =   74
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   15
         Left            =   4920
         TabIndex        =   70
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   14
         Left            =   3720
         TabIndex        =   66
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   13
         Left            =   2520
         TabIndex        =   62
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   12
         Left            =   1320
         TabIndex        =   58
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   54
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   10
         Left            =   12120
         TabIndex        =   50
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   9
         Left            =   10920
         TabIndex        =   46
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   8
         Left            =   9720
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   7
         Left            =   8520
         TabIndex        =   38
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   7320
         TabIndex        =   34
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   5
         Left            =   6120
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   3
         Left            =   3720
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   2
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   19
         Text            =   "Text3"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   3
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "Text4"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   4
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   5
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "Text6"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   6
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   35
         Text            =   "Text7"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   7
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   39
         Text            =   "Text8"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   8
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   43
         Text            =   "Text9"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   9
         Left            =   10920
         MaxLength       =   2
         TabIndex        =   47
         Text            =   "Text10"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   10
         Left            =   12120
         MaxLength       =   2
         TabIndex        =   51
         Text            =   "Text11"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   11
         Left            =   120
         MaxLength       =   2
         TabIndex        =   55
         Text            =   "Text12"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   12
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   59
         Text            =   "Text13"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   13
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   63
         Text            =   "Text14"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   14
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   67
         Text            =   "Text15"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   15
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   71
         Text            =   "Text16"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   16
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   75
         Text            =   "Text17"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   17
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   79
         Text            =   "Text18"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   18
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   83
         Text            =   "Text19"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   19
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   87
         Text            =   "Text20"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   20
         Left            =   10920
         MaxLength       =   2
         TabIndex        =   91
         Text            =   "Text21"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   21
         Left            =   12120
         MaxLength       =   2
         TabIndex        =   95
         Text            =   "Text22"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   22
         Left            =   120
         MaxLength       =   2
         TabIndex        =   99
         Text            =   "Text23"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   23
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   135
         Text            =   "Text24"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   24
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   139
         Text            =   "Text25"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   25
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   143
         Text            =   "Text26"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   26
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   105
         Text            =   "Text27"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   27
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   109
         Text            =   "Text28"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   28
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   113
         Text            =   "Text29"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Textd 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   29
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   117
         Text            =   "Text30"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   30
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   121
         Text            =   "Text31"
         Top             =   1800
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MT"
         Height          =   255
         Left            =   9840
         TabIndex        =   182
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "M"
         Height          =   255
         Left            =   10920
         TabIndex        =   129
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label11 
         Caption         =   "T"
         Height          =   255
         Left            =   11400
         TabIndex        =   131
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "N"
         Height          =   255
         Left            =   10560
         TabIndex        =   127
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "D"
         Height          =   255
         Left            =   10200
         TabIndex        =   125
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "TOTALES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   124
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   30
         Left            =   10440
         TabIndex        =   123
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   29
         Left            =   9240
         TabIndex        =   119
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   28
         Left            =   8040
         TabIndex        =   115
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   27
         Left            =   6840
         TabIndex        =   111
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   5640
         TabIndex        =   107
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   25
         Left            =   4440
         TabIndex        =   103
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   3240
         TabIndex        =   141
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   23
         Left            =   2040
         TabIndex        =   137
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   22
         Left            =   840
         TabIndex        =   101
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   12840
         TabIndex        =   97
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   11640
         TabIndex        =   93
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   10440
         TabIndex        =   89
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   9240
         TabIndex        =   85
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   8040
         TabIndex        =   81
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   6840
         TabIndex        =   77
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   5640
         TabIndex        =   73
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   4440
         TabIndex        =   69
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   3240
         TabIndex        =   65
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   2040
         TabIndex        =   61
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   840
         TabIndex        =   57
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   12840
         TabIndex        =   53
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   11640
         TabIndex        =   49
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   10440
         TabIndex        =   45
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   9240
         TabIndex        =   41
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   8040
         TabIndex        =   37
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   6840
         TabIndex        =   33
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   5640
         TabIndex        =   29
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4440
         TabIndex        =   25
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   21
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   17
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   13
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label3 
         Caption         =   "HORAS Programadas"
         Height          =   255
         Left            =   11760
         TabIndex        =   133
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   615
      Left            =   15360
      Picture         =   "ProgAdmin.frx":83472
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ubicar Registros"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Txtarjeta 
      Height          =   285
      Left            =   960
      MaxLength       =   5
      TabIndex        =   3
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      MaxLength       =   2
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   166
      Top             =   960
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   6800
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
      Caption         =   "Programacion Mensual"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label LblONOMASTICO 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Onomastico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   177
      Top             =   5160
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarjeta"
      Height          =   255
      Left            =   9480
      TabIndex        =   170
      Top             =   4920
      Width           =   615
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Apel. Paterno"
      Height          =   255
      Left            =   10920
      TabIndex        =   169
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad Org."
      Height          =   255
      Left            =   5520
      TabIndex        =   163
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Turno "
      Height          =   255
      Left            =   3480
      TabIndex        =   153
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione y Pulse Intro "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      TabIndex        =   149
      Top             =   8880
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Horario"
      Height          =   255
      Left            =   3480
      TabIndex        =   145
      Top             =   9480
      Width           =   615
   End
   Begin VB.Label Label_3 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "ProgAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim rD As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim rt1 As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw, f_sigue, F_serv, f_Fer As Boolean
Dim nposicion, dfi, dff, z, ndia, nDiaPrg, nson, nhoras As Integer
Dim sql, ccadena, cvalor, c_a, ccampo, cTarjeta, cCadRPrg, cTurnoDAnt, cTurnoDSig, cTurnoAnt As String
Dim unafecha As Date
Private Sub Command1_Click()
     c_grupo = "02": c_secuencia = "022"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub
Private Sub Form_Load()
  ProgAdministrativo.Caption = "Programa MENSUAL Administrativos - " & genaño
  DataGrid1.Enabled = False
  For i = 1 To 31
    z = i - 1
    Option1(z).Value = False
    Option1(z).Enabled = False
    Option1(z).Visible = False
    Textd(z).Text = "":   Textd(z).ForeColor = &H0&
    Textp(z).Text = "":   Textp(z).ForeColor = &H0&
    Lblhr(z).Caption = "":    Lblhr(z).ForeColor = &H0&
    Textd(z).Locked = True
    Textp(z).Locked = True
  Next
  CboHorario.Enabled = False
  Set rs = New Recordset
  '-Meses
  rs.Open "SELECT * FROM cptablas WHERE tabla='32' and codigo <>'..' ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboMes.ListCount - 1
           If CboMes.List(i) = rs.Fields("Descripcion") & " :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboMes.AddItem rs.Fields("Descripcion") & " :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  Txmes.Text = genmes
  For i = 0 To CboMes.ListCount - 1
     npos = InStr(CboMes.List(i), ":")
     nlen = Len(CboMes.List(i))
     dato = Trim(Right(CboMes.List(i), nlen - npos))
     If dato = Trim(Txmes.Text) Then
        CboMes.ListIndex = i
        Exit For
     End If
  Next i
  Call Cbomes_KeyPress(13)
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpCcosto ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To Cbo_ccosto.ListCount - 1
         If Cbo_ccosto.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("Codigo") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        Cbo_ccosto.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("Codigo")
     End If
     rs.MoveNext
  Wend
  rs.Close
  '-Turnos Disponibles   */L/V/A
  CboTurDir.Enabled = False
  Set rs = New Recordset
  rs.Open "SELECT * FROM cptablas WHERE tabla='31' and (Patindex('%* %',Codigo)>0 or codigo='L' or  codigo='V' or codigo='A' )  ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboTurDir.ListCount - 1
           If CboTurDir.List(i) = rs.Fields("Descripcion") & " :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboTurDir.AddItem rs.Fields("Descripcion") & " :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  '-Horarios administrativos Disponibles
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpHorarioNew WHERE asistencial<>'S' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly    'CpHorarios
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboHorario.ListCount - 1
           If CboHorario.List(i) = Left(rs.Fields("descripcion"), 30) & " (" & Left(rs.Fields("horas"), 2) + ":" + Right(rs.Fields("horas"), 2) & ") =>" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboHorario.AddItem Left(rs.Fields("descripcion"), 30) & " (" & Left(rs.Fields("horas"), 2) + ":" + Right(rs.Fields("horas"), 2) & ") =>" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  If Wccosto <> "" Then
     Tx_ccosto.Text = Wccosto
     For i = 0 To Cbo_ccosto.ListCount - 1
        npos = InStr(Cbo_ccosto.List(i), ":")
        nlen = Len(Cbo_ccosto.List(i))
        dato = Trim(Right(Cbo_ccosto.List(i), nlen - npos))
        If dato = Trim(Wccosto) Then
           Cbo_ccosto.ListIndex = i
           Exit For
        End If
     Next i
  End If
End Sub
Private Sub Cbomes_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboMes.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbomes_LostFocus()
   If IsEmpty(CboMes.Text) Then
   Else
      Call Cbomes_KeyPress(13)
      Exit Sub
   End If
End Sub
Private Sub Cbomes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboMes.Text) Then
         CboMes.SetFocus
      Else
         j = CboMes.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboMes.List(j), ":")
            nlen = Len(CboMes.List(j))
            Txmes.Text = Right(CboMes.List(j), nlen - nposicion)
            TxNombMes.Text = Trim(Left(CboMes.List(j), nposicion - 1))
            unafecha = "01/" + Txmes.Text + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)
            Call Cargadias       'primer/ultimo
            fecCierre = ultimo
            
            Set rs = New Recordset
            rs.Open "SELECT * FROM CpCronograma WHERE ano='" & genaño & "' and Mes='" & Txmes.Text & "'", Cnx, adOpenStatic, adLockReadOnly   'CpHorarios
            If rs.EOF Then
               MsgBox "No Existe Cronograma del MES... Verifique", vbExclamation
            Else
               If IsDate(rs!fecha) Then fecCierre = rs!fecha
               If fecCierre <= Date Then
                  Me.Caption = "Programa de ADMINISTRATIVOS... CERRADO.  "
               Else
                  Me.Caption = "Programa de ADMINISTRATIVOS"
               End If
            End If
            rs.Close
         
         End If
      End If
      Exit Sub
   End If
End Sub
Private Sub Cbo_ccosto_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Cbo_ccosto.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbo_ccosto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If IsEmpty(Cbo_ccosto.Text) Then
     Else
        Txtarjeta.Text = "": Txtrabajador.Text = "": Txfecnac.Text = "": TxCarEst.Text = "": TxDscCargo.Text = ""
        DataGrid1.Enabled = True
        j = Cbo_ccosto.ListIndex
        If j <> -1 Then
           nposicion = InStr(Cbo_ccosto.List(j), ":")
           nlen = Len(Cbo_ccosto.List(j))
           Tx_ccosto.Text = Right(Cbo_ccosto.List(j), nlen - nposicion)
           If Wccosto <> "" Then
              If Tx_ccosto.Text <> Wccosto Then
                 MsgBox "Centro de Costo,No autorizado..", vbExclamation
                 Exit Sub
              End If
           End If
           Txdscccosto.Text = Left(Cbo_ccosto.List(j), nposicion - 1)
           Call cmdcarga_Click
        End If
        Exit Sub
     End If
   End If
End Sub
Private Sub Cbo_ccosto_LostFocus()
   If IsEmpty(Cbo_ccosto.Text) Then
   Else
      j = Cbo_ccosto.ListIndex
      If j <> -1 Then
         nposicion = InStr(Cbo_ccosto.List(j), ":")
         nlen = Len(Cbo_ccosto.List(j))
         Tx_ccosto.Text = Right(Cbo_ccosto.List(j), nlen - nposicion)
         If Wccosto <> "" Then
            If Tx_ccosto.Text <> Wccosto Then
               MsgBox "Centro de Costo,No autorizado..", vbExclamation
               Exit Sub
            End If
         End If
         Txdscccosto.Text = Left(Cbo_ccosto.List(j), nposicion - 1)
         Call cmdcarga_Click
      End If
      Exit Sub
   End If
End Sub
Private Sub CmdPrograma_Click()
   If MsgBox("Confirma, PROGRAMAR servidores ?   ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   cTarjeta = "":  cNombre = "": cCadRPrg = "": ccCosto = "": cCarest = "": ccargo = "": cdscCargo = ""
   cGrpGen = "":      cplh = "": cSabado = "": nhoras = 0: ntHoras = 0: nDiaPrg = 0: ndia = 0: nson = 0
   Set rs = New Recordset
   If Trim(Tx_ccosto.Text) <> "" Then
      Txdscccosto.Text = ""
      Cbo_ccosto.Text = ""
      SQ1 = "SELECT * FROM Cpgeneral  WHERE estado='A' and TipoPla='ADM' and cCosto ='" & Tx_ccosto.Text & "'"
   Else
      SQ1 = "SELECT * FROM Cpgeneral  WHERE estado='A' and TipoPla='ADM' ORDER By CCosto"
   End If
   rs.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   f_sigue = False
   '-Proceso ADMINISTRATIVOS
   Do While Not rs.EOF
     cTarjeta = rs!tarjeta
     CmdPrograma.Caption = cTarjeta
     f_sigue = True
     nDiaPrg = 0      'programa turnos UPR
     cNombre = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 25)
     cCarest = IIf(IsNull(rs!carest), "", rs!carest)
     ccargo = IIf(IsNull(rs!cargo), "", rs!cargo)
     cGrpGen = IIf(IsNull(rs!GrpGen), "", rs!GrpGen)
     cplh = IIf(IsNull(rs!plh), "", rs!plh)
     ccCosto = IIf(IsNull(rs!cCosto), "", rs!cCosto)
     cSabado = "": cHorario = ""
     If cGrpGen <> "" Then
        Set rs_d = New Recordset
        sq_l = "SELECT * FROM CpGrupos WHERE Grupo='" & cGrpGen & "'"
        rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs_d.EOF Then
           cSabado = IIf(IsNull(rs_d!sabado), "N", rs_d!sabado)
           cHorario = IIf(IsNull(rs_d!horario), "N", rs_d!horario)
        End If
        rs_d.Close
     End If
     Set rs_d = New Recordset
     sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & cHorario & "'"     'CpHorarios
     rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
     If Not rs_d.EOF Then
        cvalor = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
        nhoras = Val(Left(cvalor, 2)) + (Val(Right(cvalor, 2)) / 60)
     End If
     rs_d.Close
     Set rt = New Recordset
     sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & cTarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
     rt.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
     If rt.EOF Then
        rt.AddNew
        rt!tarjeta = cTarjeta
        rt!ano = genaño
        rt!mes = Txmes.Text
        rt!usuario = wUsuario      'Firma Usuario
        rt!fecha = Date
     Else
        '-Verifica No exista Programa del titular
        For i = Day(primer) To Day(ultimo)
           z = i - 1
           chaytur = IIf(IsNull(rt.Fields(z + 4)), "", Trim(rt.Fields(z + 4)))
           'If chaytur <> "" And InStr("L_V_R_A", chaytur) = 0 Then
           If chaytur <> "" And chaytur = "A" Then
              nDiaPrg = nDiaPrg + 1
           End If
        Next i
     End If
     If nDiaPrg > 0 Then
        '-<E>,Existe programa anterior para el servidor
        cCadRPrg = cCadRPrg + (cTarjeta + "_")
        nson = nson + 1
     End If 'else
     
        ntHoras = 0:      ndia = 0:
        For i = Day(primer) To Day(ultimo)
           cPatron = "":   Sw = True: chaytur = ""
           z = i - 1
           FechaB = Format(primer + z, "yyyymmdd")
           cPatron = Mid(ccadena, i, 1)     ' F/O
           If Weekday(primer + i - 1) = 1 Then
              'NO Programa DOMINGO
              cPatron = "F"
           End If
           chaytur = IIf(IsNull(rt.Fields(z + 4)), "", Trim(rt.Fields(z + 4)))
           If cPatron = "F" Then
              rt.Fields(z + 4) = ""
              Sw = False           'NO Programa Feriado /DOMINGO
           Else
              If Weekday(primer + i - 1) = 7 Then
              ' Left(WeekdayName(Weekday(primer + i - 1)), 3) = "Sáb"
                 If cSabado = "S" Then
                    If chaytur = "" Or (Trim(chaytur) <> "" And InStr("L_V_R", chaytur) = 0) Then
                       rt.Fields(z + 4) = "A"
                       ntHoras = ntHoras + nhoras
                       ndia = ndia + 1
                    End If
                 Else
                    rt.Fields(z + 4) = ""
                    Sw = False      'Horario no aplica SABADO - NO Programa
                 End If
              Else
                 If chaytur = "" Or (Trim(chaytur) <> "" And InStr("L_V_R", chaytur) = 0) Then
                    rt.Fields(z + 4) = "A"
                    ntHoras = ntHoras + nhoras
                    ndia = ndia + 1
                 End If
              End If
           End If
           If Sw And (chaytur = "" Or (Trim(chaytur) <> "" And InStr("L_V_R", chaytur) = 0)) Then
              Set rt1 = New Recordset
              SQ1 = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & FechaB & "'"
              rt1.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
              If rt1.EOF Then
                 rt1.AddNew
                 rt1!tarjeta = cTarjeta
                 rt1!fecha = primer + z
                 rt1!usuario = wUsuario      'Firma Usuario
                 rt1!fecRegistro = Date
              End If
              rt1!horario = cHorario
              rt1.Update
              rt1.Close
           Else
              If cPatron = "F" Or chaytur = "" Or (Weekday(primer + i - 1) = 7 And cSabado = "") Then
                 Cnx.Execute "DELETE FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "'AND fecha ='" & FechaB & "'"
              End If
           End If
        Next i
        rt!Horas = ntHoras
        rt!m = ndia
        rt!carest = IIf(IsNull(cCarest), "", cCarest)
        rt!cCosto = IIf(IsNull(ccCosto), "", ccCosto)
        rt!cargo = IIf(IsNull(ccargo), "", ccargo)
        rt!Flag = "1"       'ADMINISTRATIVO
        rt!plh = IIf(IsNull(cplh), "", cplh)
        rt.Update
        rt.Close
     rs.MoveNext
   Loop
   CmdPrograma.Caption = "Programa"
   If nson > 0 Then
      MsgBox "<E> Servidores R-programados : " + Trim(Str(nson)), vbExclamation
   End If
   Call cmdcarga_Click
End Sub
Private Sub Cmdservidor_Click()
   cCadUpr = "": nDiaPrg = 0
   Set rs = New Recordset
   SQ1 = "SELECT * FROM Cpgeneral  WHERE estado='A' and TipoPla='ADM' and tarjeta='" & Txtarjeta.Text & "'"
   rs.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   f_sigue = False
   Do While Not rs.EOF
     cTarjeta = rs!tarjeta
     CmdServidor.Caption = cTarjeta
     f_sigue = True
     cNombre = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 25)
     cCarest = IIf(IsNull(rs!carest), "", rs!carest)
     ccargo = IIf(IsNull(rs!cargo), "", rs!cargo)
     cGrpGen = IIf(IsNull(rs!GrpGen), "", rs!GrpGen)
     cplh = IIf(IsNull(rs!plh), "", rs!plh)
     ccCosto = IIf(IsNull(rs!cCosto), "", rs!cCosto)
     Tx_ccosto.Text = ccCosto
     cSabado = "": cHorario = ""
     If cGrpGen <> "" Then
        Set rs_d = New Recordset
        sq_l = "SELECT * FROM CpGrupos WHERE Grupo='" & cGrpGen & "'"
        rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs_d.EOF Then
           cSabado = IIf(IsNull(rs_d!sabado), "N", rs_d!sabado)
           cHorario = IIf(IsNull(rs_d!horario), "N", rs_d!horario)
        End If
        rs_d.Close
     End If
     Set rs_d = New Recordset
     sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & cHorario & "'"
     rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
     If Not rs_d.EOF Then
        cvalor = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
        nhoras = Val(Left(cvalor, 2)) + (Val(Right(cvalor, 2)) / 60)
     End If
     rs_d.Close
     
     Set rt = New Recordset
     sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & cTarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
     rt.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
     If rt.EOF Then
        rt.AddNew
        rt!tarjeta = cTarjeta
        rt!ano = genaño
        rt!mes = Txmes.Text
        rt!usuario = wUsuario      'Firma Usuario
        rt!fecha = Date
     Else
        '-Verifica No exista Programa del titular
        For i = Day(primer) To Day(ultimo)
           z = i - 1
           chaytur = IIf(IsNull(rt.Fields(z + 4)), "", Trim(rt.Fields(z + 4)))
           'If chaytur <> "" And InStr("L_V_R_A", chaytur) = 0 Then
           If chaytur <> "" And chaytur = "A" Then
              nDiaPrg = nDiaPrg + 1
           End If
        Next i
     End If
     nson = 0
     If nDiaPrg > 0 Then
        '-<E>,Existe programa anterior para el servidor
        cCadRPrg = cCadRPrg + (cTarjeta + "_")
        nson = nson + 1
     End If
     
        ntHoras = 0:      ndia = 0:
        For i = Day(primer) To Day(ultimo)
           cPatron = "":   Sw = True: chaytur = ""
           z = i - 1
           chaytur = IIf(IsNull(rt.Fields(z + 4)), "", Trim(rt.Fields(z + 4)))
           FechaB = Format(primer + z, "yyyymmdd")
           cPatron = Mid(ccadena, i, 1)     ' F/O
           If Weekday(primer + i - 1) = 1 Then
              cPatron = "F"   ' NO Programa DOMINGO
           End If
           If cPatron = "F" Then
              rt.Fields(z + 4) = ""
              Sw = False           'NO Programa Feriado/Domingo
           Else
             If Weekday(primer + i - 1) = 7 Then
                'Left(WeekdayName(Weekday(primer + i - 1)), 3) = "Sáb"
                If cSabado = "S" Then
                   If chaytur = "" Or (Trim(chaytur) <> "" And InStr("L_V_R", chaytur) = 0) Then
                      rt.Fields(z + 4) = "A"
                      ntHoras = ntHoras + nhoras
                      ndia = ndia + 1
                   End If
                Else
                   rt.Fields(z + 4) = ""
                   Sw = False       'Horario no aplica SABADO - NO Programa
                End If
             Else
                If chaytur = "" Or (Trim(chaytur) <> "" And InStr("L_V_R", chaytur) = 0) Then
                   rt.Fields(z + 4) = "A"
                   ntHoras = ntHoras + nhoras
                   ndia = ndia + 1
                End If
             End If
           End If
           If Sw And (chaytur = "" Or (Trim(chaytur) <> "" And InStr("L_V_R", chaytur) = 0)) Then
              Set rt1 = New Recordset
              SQ1 = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & FechaB & "'"
              rt1.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
              If rt1.EOF Then
                 rt1.AddNew
                 rt1!tarjeta = cTarjeta
                 rt1!fecha = primer + z
                 rt1!usuario = wUsuario      'Firma Usuario
                 rt1!fecRegistro = Date
              End If
              rt1!horario = cHorario
              rt1.Update
              rt1.Close
           Else
              If cPatron = "F" Or chaytur = "" Then
                 Cnx.Execute "DELETE FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "'AND fecha ='" & FechaB & "'"
              End If
           End If
        Next i
        rt!Horas = ntHoras
        rt!m = ndia
        rt!carest = IIf(IsNull(cCarest), "", cCarest)
        rt!cCosto = IIf(IsNull(ccCosto), "", ccCosto)
        rt!cargo = IIf(IsNull(ccargo), "", ccargo)
        rt!Flag = "1"       'ADMINISTRATIVO
        rt!plh = IIf(IsNull(cplh), "", cplh)
        rt.Update
        rt.Close
     
     rs.MoveNext
   Loop
   CmdServidor.Caption = "Prog. Servidor"
   If nDiaPrg > 0 Then
      MsgBox "<!> Servidor Reprogramado... ", vbExclamation
   End If
   Call cmdcarga_Click
   Call CmdLimpia_Click
End Sub
Private Sub cmdcarga_Click()
   Set DD = New Recordset
   With DD
     .Fields.Append "Tarjeta", adChar, 5
     .Fields.Append "Nombre", adChar, 25
     .Fields.Append "Cargo", adChar, 15
     .Fields.Append "D", adChar, 2
     .Fields.Append "M", adChar, 2
     .Fields.Append "T", adChar, 2
     .Fields.Append "N", adChar, 2
     .Fields.Append "Horas", adDouble
     For i = dfi To dff
        ccampo = "D" + Trim(Str(i))
        .Fields.Append ccampo, adChar, 2
     Next i
     .Fields.Append "MT", adChar, 2
     .Open
   End With
   Set DataGrid1.DataSource = DD
   For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
        Select Case i
          Case 0
            .Caption = "Tarjeta"
            .Width = Me.TextWidth(String(7, "S"))
          Case 1
            '.Caption =
            .Width = Me.TextWidth(String(22, "S"))
          Case 2
            '.Caption =
            .Width = Me.TextWidth(String(15, "S"))
          Case 3
            .Width = Me.TextWidth(String(2, "S"))
          Case 4
            .Width = Me.TextWidth(String(2, "S"))
          Case 5
            .Width = Me.TextWidth(String(2, "S"))
          Case 6
            .Width = Me.TextWidth(String(2, "S"))
          Case 7
            .Width = Me.TextWidth(String(5, "S"))
          Case Else
            .Width = Me.TextWidth(String(3, "S"))
        End Select
        '.NumberFormat = "##,##0.0"
        .Locked = True
     End With
   Next
   Fecha1 = Format(primer, "yyyymmdd")
   fecha2 = Format(ultimo, "yyyymmdd")
   '-DIRECTORIO Administrativo
   sql = "SELECT * FROM CpDirectorio  WHERE ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "' AND flag='1' ORDER BY Tarjeta"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   Do While Not rs.EOF
      cTarjeta = rs!tarjeta:  cNombre = "":   cCarest = "":  ccargo = "":  cdscCargo = ""
      cGrpGen = "":      cplh = ""
      SQ1 = "SELECT * FROM Cpgeneral  WHERE tarjeta='" & cTarjeta & "'"
      Set rs_d = New Recordset
      rs_d.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      f_sigue = False
      If rs_d.EOF Then
      Else
         If rs_d!cCosto = Tx_ccosto.Text Then
            f_sigue = True
            cNombre = Left(Trim(rs_d!apelpat) + " " + Trim(rs_d!apelmat) + " " + Trim(rs_d!nombres), 25)
            cCarest = IIf(IsNull(rs_d!carest), "", rs_d!carest)
            ccargo = IIf(IsNull(rs_d!cargo), "", rs_d!cargo)
            cGrpGen = IIf(IsNull(rs_d!GrpGen), "", rs_d!GrpGen)
            cplh = IIf(IsNull(rs_d!plh), "", rs_d!plh)
          End If
      End If
      rs_d.Close
      If f_sigue Then
         Set rs_d = New Recordset
         rs_d.Open "SELECT * FROM Cpcargos WHERE codigo='" & ccargo & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs_d.EOF Then
            cdscCargo = Left(rs_d!descripcion, 15)
         End If
         rs_d.Close
         DD.AddNew
         DD!tarjeta = rs!tarjeta
         DD!nombre = cNombre
         DD!cargo = cdscCargo
         DD!d = IIf(IsNull(rs!d), "", rs!d)
         DD!n = IIf(IsNull(rs!n), "", rs!n)
         DD!m = IIf(IsNull(rs!m), "", rs!m)
         DD!T = IIf(IsNull(rs!T), "", rs!T)
         DD!Horas = IIf(IsNull(rs!Horas), 0, rs!Horas)
         For i = Day(primer) To Day(ultimo)
            z = i - 1
            '- directorio rs.fields(Z+4)  ==>    dd.fields(Z+8)
            '   Textd(z) = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
            'cCampo = rs.Fields(z + 4).Name
            DD.Fields(z + 8) = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
         Next i
         DD!MT = IIf(IsNull(rs!MT), "", rs!MT)
         DD.Update
      End If
      rs.MoveNext
   Loop
   rs.Close
   DataGrid1.Refresh
   Command1.SetFocus
   '   DataGrid1.SetFocus
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Txtarjeta.Text = DataGrid1.Columns(0)
   LblONOMASTICO.Caption = "Onomastico"
   LblONOMASTICO.Visible = False
   Call Txtarjeta_KeyPress(13)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   With DD
      If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
         .Sort = .Fields(ColIndex).[Name] & " Desc"
      Else
         .Sort = .Fields(ColIndex).[Name] & " Asc"
      End If
   End With
   Call Limpiacampos
End Sub
Private Sub TxBUscar_KeyPress(KeyAscii As Integer)
  If nposicion = 0 Then Exit Sub
  If KeyAscii = 13 Then
     Dim cBuscar As String
     cpuntero = DataGrid1.Bookmark
     If Txbuscar.Text = "" Then Exit Sub
     cBuscar = "tarjeta like '" & Txbuscar.Text & "*'"
     DD.MoveFirst
     DD.Find (cBuscar)
     If DD.EOF Then
        Txbuscar.BackColor = &HFF&
     Else
        Txbuscar.BackColor = &H80000005
        DataGrid1.Refresh
        Txbuscar.Text = ""
        DataGrid1_Click
     End If
  End If
End Sub
Private Sub Txbuscar1_KeyPress(KeyAscii As Integer)
   If nposicion = 0 Then Exit Sub
   If KeyAscii = 13 Then
     Dim cBuscar As String
     cpuntero = DataGrid1.Bookmark
     If Txbuscar1.Text = "" Then Exit Sub
     cBuscar = "Nombre like '" & Txbuscar1.Text & "*'"
     DD.MoveFirst
     DD.Find (cBuscar)
     If DD.EOF Then
        Txbuscar1.BackColor = &HFF&
     Else
        Txbuscar1.BackColor = &H80000005
        DataGrid1.Refresh
        Txbuscar1.Text = ""
        DataGrid1_Click
     End If
  End If
End Sub

Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Frame1.Caption = "Programa - TRABAJADOR"
      LblONOMASTICO.Caption = "Onomastico"
      LblONOMASTICO.Visible = False
      If Trim(Txtarjeta.Text) = "" Or (Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta)) Then
         If Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta) Then
            Ayuda1.TxUbiAP = Trim(Txtarjeta)
         End If
         Load Ayuda1
         Ayuda1.Tag = 6.1
         Ayuda1.Show 1
      Else
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE tarjeta = '" & Txtarjeta & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "Tarjeta, No Ubicada", vbExclamation
            Txtarjeta = ""
         Else
            Txtrabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
            Txccosto.Text = IIf(IsNull(rs!cCosto), "", rs!cCosto)
            TxCarEst.Text = IIf(IsNull(rs!carest), "", rs!carest)
            TxPlH.Text = IIf(IsNull(rs!plh), "", rs!plh)
            TxCargo.Text = IIf(IsNull(rs!cargo), "", rs!cargo)
            TxGrpGen.Text = IIf(IsNull(rs!GrpGen), "", rs!GrpGen)
            If IsDate(rs!fecnac) Then Txfecnac.Text = rs!fecnac
         End If
         rs.Close
      End If
      If Txtarjeta <> "" Then
         If IsDate(Txfecnac.Text) Then
            If Month(Txfecnac.Text) = Val(Txmes.Text) Then
               LblONOMASTICO.Visible = True
               LblONOMASTICO.Caption = "Onomastico el " + Str(Day(Txfecnac.Text))
            End If
         End If
         Set rs_d = New Recordset
         rs_d.Open "SELECT * FROM Cpcargos WHERE  codigo='" & TxCargo & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs_d.EOF Then
            TxDscCargo.Text = Left(rs_d!descripcion, 15)
         End If
         rs_d.Close
         '-Caso Filtro Operador, pertenece a un wccosto
         f_sigue = True
         If Trim(Wccosto) <> "" Then
            If Txccosto.Text <> Wccosto Then
               f_sigue = False
            End If
         End If
         If f_sigue Then
            If IsDate(Txfecnac) Then    '<-----
            If Month(Txfecnac.Text) = Val(Txmes.Text) Then
               '-Valida y genera autorización ONO en dia, Jef UPR debe programar 1/2turno
               If Mid(ccadena, Day(DateValue(Txfecnac.Text)), 1) = "F" Then
               Else
                  Set rs_d = New Recordset
                  onofecha = DateValue(Right("0" + Trim(Day(Txfecnac.Text)), 2) + "/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño)))
                  FechaB = Format(onofecha, "yyyymmdd")
                  rs_d.Open "SELECT * FROM CpAutori" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND DiaHora='D' AND fechaini<='" & FechaB & "'AND fechaFin>='" & FechaB & "' ", Cnx, adOpenStatic, adLockReadOnly
                  If rs_d.EOF Then
                     rs_d.Close
                     rs_d.Open "SELECT * FROM CpAutori" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND codigo='ONO' AND DiaHora='D' AND fechaini='" & FechaB & "'", Cnx, adOpenStatic, adLockOptimistic
                     If rs_d.EOF Then
                        rs_d.AddNew
                        rs_d!tarjeta = Txtarjeta
                        rs_d!Codigo = "ONO"
                        rs_d!diahora = "D"
                        rs_d!fechaIni = onofecha
                        rs_d!fechaFin = onofecha
                        rs_d!documento = "autorizacion automatica"
                        rs_d!usuario = wUsuario      'Firma Usuario
                        rs_d!fecha = Date
                        rs_d.Update
                     End If
                  End If
                  rs_d.Close
               End If
            End If
            End If   '----->
         Else
            MsgBox "Trabajador No pertenece al C.Costo..!", vbExclamation
            Txtarjeta.Text = "": Txtrabajador.Text = "": Txfecnac.Text = ""
            TxCarEst.Text = "": Txccosto.Text = "": TxPlH.Text = "": TxCargo.Text = "":  TxDscCargo.Text = ""
            TxGrpGen.Text = ""
            LblONOMASTICO.Caption = "Onomastico"
            LblONOMASTICO.Visible = False
            Exit Sub
         End If
      End If
      For i = dfi To dff
        z = i - 1
        Option1(z).Value = False
        Textd(z).Visible = True: Textp(z).Visible = True
        Textd(z).Text = "": Textd(z).ForeColor = &H80000012
        Textp(z).Text = "":  Textp(z).ForeColor = &H80000012
        Lblhr(z).Caption = ""
      Next
      Call cmdcargTRA_Click
   End If
End Sub
Private Sub cmdcargTRA_Click()
   TxHoras.Text = "": TxD.Text = "": TxN.Text = "": TxM.Text = "": Txt.Text = "": TxMT.Text = ""
   If Txtarjeta.Text = "" Then Exit Sub
   For i = dfi To dff
       z = i - 1
       Option1(z).Value = False
       Textd(z).Visible = True: Textp(z).Visible = True
       Textd(z).Text = "": Textd(z).ForeColor = &H80000012
       Textp(z).Text = "":  Textp(z).ForeColor = &H80000012
       Lblhr(z).Caption = ""
   Next
   Fecha1 = Format(primer, "yyyymmdd")
   fecha2 = Format(ultimo, "yyyymmdd")
   '-carga  DIRECTORIO
   sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      For i = Day(primer) To Day(ultimo)
         z = i - 1
         'For i = 0 To rs.Fields.Count - 1
         'cCampo = rs.Fields(z + 4).Name
         'If InStr("C1049_C1053_C1054_C1056_C1064_C1160", cCampo) > 0 Then
         Textd(z) = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
         Select Case Trim(Textd(z).Text)
              Case "D"
                 TxD.Text = Val(TxD.Text) + 1
              Case "D1"
                 TxD.Text = Val(TxD.Text) + 1
               Case "D2"
                 TxD.Text = Val(TxD.Text) + 1
              Case "MT"
                 TxMT.Text = Val(TxMT.Text) + 1
                 
              Case "N"
                 TxN.Text = Val(TxN.Text) + 1
              Case "N2"
                 TxN.Text = Val(TxN.Text) + 1
              Case "A"
                 TxM.Text = Val(TxM.Text) + 1
              Case "M"
                 TxM.Text = Val(TxM.Text) + 1
              Case "M1"
                 TxM.Text = Val(TxM.Text) + 1
              Case "M2"
                 TxM.Text = Val(TxM.Text) + 1
              Case "M3"
                 TxM.Text = Val(TxM.Text) + 1
              Case "T"
                 Txt.Text = Val(Txt.Text) + 1
              Case "T1"
                 Txt.Text = Val(Txt.Text) + 1
              Case "T2"
                 Txt.Text = Val(Txt.Text) + 1
                Case "H"
                   TxM.Text = Val(TxM.Text) + 1
                   TxN.Text = Val(TxN.Text) + 1
                Case "H1"
                   TxM.Text = Val(TxM.Text) + 1
                   TxN.Text = Val(TxN.Text) + 1
                Case "C"
                   Txt.Text = Val(Txt.Text) + 1
                   TxN.Text = Val(TxN.Text) + 1
                Case "C1"
                   Txt.Text = Val(Txt.Text) + 1
                   TxN.Text = Val(TxN.Text) + 1
                Case "P"
                   TxD.Text = Val(TxD.Text) + 1
                   TxN.Text = Val(TxN.Text) + 1
                Case "P1"
                   TxD.Text = Val(TxD.Text) + 1
                   TxN.Text = Val(TxN.Text) + 1
         End Select
      Next i
   End If
   rs.Close
   '-Carga HORARIO
   sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha >='" & Fecha1 & "' AND fecha<='" & fecha2 & "' ORDER BY Fecha,tarjeta"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      rs.MoveFirst
      Do While Not rs.EOF
         i = Day(rs!fecha)
         z = i - 1
         Textp(z) = IIf(IsNull(rs!horario), "", rs!horario)
         If Textp(z) <> "" Then
            Set rs_d = New Recordset
            sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & Trim(Textp(z)) & "'"     'CpHorarios
            rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs_d.EOF Then
               cvalor = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
               TxHoras.Text = Val(TxHoras.Text) + Val(Left(cvalor, 2)) + (Val(Right(cvalor, 2)) / 60)
               If Val(cvalor) > 0 Then
                  Lblhr(z).Caption = Left(cvalor, 2) + ":" + Right(cvalor, 2)
               Else
                  Lblhr(z).Caption = ""
               End If
            End If
            rs_d.Close
         End If
         rs.MoveNext
      Loop
   End If
   rs.Close
End Sub
Private Sub CmdClear_Click()
    Txtarjeta.Text = "": Txtrabajador.Text = "": Txfecnac.Text = ""
    Txccosto.Text = "": TxPlH.Text = "": TxCarEst.Text = "": TxCargo.Text = "": TxDscCargo.Text = "": TxGrpGen.Text = ""
    TxD.Text = "": TxN.Text = "": TxM.Text = "": Txt.Text = "": TxHoras.Text = "":  TxMT.Text = ""
    CboTurDir.Text = "": TxTurno.Text = ""
    CboHorario.Text = "": TxHorario.Text = ""
    Txbuscar.Text = "": Txbuscar1.Text = ""
    For i = dfi To dff
       z = i - 1
       Option1(z).Value = False
       Textd(z).Visible = True: Textp(z).Visible = True
       Textd(z).Text = "": Textd(z).ForeColor = &H80000012  '&HFF0000
       Textp(z).Text = "":  Textp(z).ForeColor = &H80000012 '&HFF0000
       Lblhr(z).Caption = ""
    Next
    Cbo_ccosto.Text = "":     Tx_ccosto.Text = ""
    Txdscccosto.Text = ""
    DD.MoveFirst
    Do While Not DD.EOF
       DD.Delete
       DD.MoveNext
    Loop
    Cbo_ccosto.SetFocus
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
      Case 1
   End Select
   CboTurDir.Enabled = True
   TxTurno.Text = ""
   cTurnoAnt = ""
   CboHorario.Enabled = True
   TxHorario.Text = ""
   dfi = Day(primer)
   dff = Day(ultimo)
   For i = dfi To dff
     z = i - 1
     If Option1(z).Value = True Then
        If InStr(Textd(z).Text, "L_V_R") > 0 And Trim(Textd(z).Text) <> "" Then
           '-No Editable
           Textd(z).ForeColor = &H0&
           Option1(z).Value = False
           Textd(z).Locked = True:                  Textp(z).Locked = True
           Exit Sub
        End If
        cTurnoAnt = Textd(z).Text
        TxTurno.Text = Textd(z).Text
        TxHorario.Text = Textp(z).Text
        Textd(z).ForeColor = &HFF0000
        Textd(z).Locked = False
        Textd(z).SetFocus
        Exit Sub
     End If
   Next
End Sub
Private Sub Textd_LostFocus(Index As Integer)
      Textd(Index) = UCase(Textd(Index))
      cTurnoDAnt = "OK": cTurnoDSig = "OK"
      If Trim(Textd(Index).Text) <> "" Then
         For i = 0 To CboTurDir.ListCount - 1
            npos = InStr(CboTurDir.List(i), ":")
            nlen = Len(CboTurDir.List(i))
            dato = Trim(Right(CboTurDir.List(i), nlen - npos))
            If dato = Trim(Textd(Index).Text) Then
               CboTurDir.ListIndex = i
               Exit For
            End If
         Next i
         
         '-Bloqueo sujeto ha :    N-N , N-D, N-M, N-T
         If InStr("NN*N2N3_DD1D2D3_D*_MM1M2M3_TT1T2_HCP_H*H1_C*C1_P*P1_MT", Textd(Index)) > 0 And Textd(Index) <> "*" Then
            '-Verifica Dia anterior
            If Index = 0 Then
               dfecha = primer - 1    'ubica ult dia mes anterior
               ndia = Day(dfecha)
               z = ndia - 1
               cmes = Right("0" + Trim(Str(Month(dfecha))), 2)
               cano = Trim(Str(Year(dfecha)))
               sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & cano & "' AND mes='" & cmes & "'"
               Set rs = New Recordset
               rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not rs.EOF Then
                  cTurnoDAnt = IIf(IsNull(rs.Fields(z + 4)), cTurnoDAnt, rs.Fields(z + 4))
               End If
               rs.Close
            Else
               cTurnoDAnt = IIf(IsNull(Textd(Index - 1)), cTurnoDAnt, Textd(Index - 1))
            End If
            If InStr("NN*N2N3_4_HCP_H*H1_C*C1_P*P1", cTurnoDAnt) > 0 And cTurnoDAnt <> "" And cTurnoDAnt <> "*" Then
               MsgBox "Turno No es Valido, verifique Dia Anterior... ", vbExclamation
               Textd(Index).ForeColor = &H0&
               Textd(Index) = cTurnoAnt
               Exit Sub
            End If
         End If
         If InStr("NN*N2N3_4_HCP_H*H1_C*C1_P*P1", Textd(Index)) > 0 And Textd(Index) <> "*" Then
            '-Verifica Dia Siguiente
            If Index + 1 = ultimo Then  'ubica dia mes siguiente
               dfecha = ultimo + 1
            Else
               dfecha = primer + Index + 1
            End If
            ndia = Day(dfecha)
            z = ndia - 1
            cmes = Right("0" + Trim(Str(Month(dfecha))), 2)
            cano = Trim(Str(Year(dfecha)))
            If Txmes.Text = cmes Then
               '-Next day
               cTurnoDSig = IIf(IsNull(Textd(Index + 1)), cTurnoDSig, Textd(Index + 1))
            Else
              '-Next Month
               sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & cano & "' AND mes='" & cmes & "'"
               Set rs = New Recordset
               rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not rs.EOF Then
                  cTurnoDSig = IIf(IsNull(rs.Fields(z + 4)), cTurnoDSig, rs.Fields(z + 4))
               End If
               rs.Close
            End If
            If InStr("NN*N2N3_DD1D2D3_D*_MM1M2M3_TT1T2_HCP_H*H1_C*C1_P*P1_MT", cTurnoDSig) > 0 And cTurnoDSig <> "" And cTurnoDSig <> "*" Then
               MsgBox "Turno No es valido, Verifique Dia Siguiente... ", vbExclamation
               Textd(Index).ForeColor = &H0&
               Textd(Index) = cTurnoAnt
               Exit Sub
            End If
         End If
         If InStr("NN*N2N3_HCP_H*H1_C*C1_P*P1", Textd(Index)) > 0 And Textd(Index) <> "*" Then
            If Val(TxN.Text) = GenMaxTN Then
               MsgBox "Turno No es valido, Excede " + Trim(Str(GenMaxTN)) + " Noches Permitidas... ", vbExclamation
               Textd(Index).ForeColor = &H0&
               Textd(Index) = cTurnoAnt
               Exit Sub
            End If
         End If
         '-Valida Max dias Libres  GENMAXDL
         If InStr("ANN*_DD1D2D3_D*_MM1M2M3_TT1T2_HCP_H*H1_C*C1_P*P1_MT", Textd(Index)) > 0 And Textd(Index) <> "*" Then
            '-Valida Max dias Libres
            nposicion = 0
            cvalor = ""
            dfecha = primer + Index    'ubica dia base
            For i = 1 To GenMaxDL          '4
               ndia = Day(dfecha - i)
               cmes = Right("0" + Trim(Str(Month(dfecha - i))), 2)
               cano = Trim(Str(Year(dfecha - i)))
               If Index - i >= 0 Then
                  If Textd(Index - i) = "*" Or Len(Trim(Textd(Index - i))) = 0 Then
                      nposicion = nposicion + 1
                      cvalor = cvalor + Trim(Str(Index - i + 1)) + "-"
                  Else
                      nposicion = 0
                      cvalor = ""
                  End If
               Else
                  sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & cano & "' AND mes='" & cmes & "'"
                  Set rs = New Recordset
                  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  z = ndia - 1
                  If Not rs.EOF Then
                     If rs.Fields(z + 4) = "*" Or IsEmpty(rs.Fields(z + 4)) Then
                        nposicion = nposicion + 1
                        cvalor = cvalor + Trim(Str(ndia)) + "-"
                     Else
                        nposicion = 0
                        cvalor = ""
                     End If
                  End If
                  rs.Close
               End If
            Next i
         End If
         If nposicion > GenMaxDL Then
            MsgBox "VerifiquE BLOQUE de Dias (" + Trim(Str(nposicion)) + "), Libres.. " + cvalor, vbExclamation
            Textd(Index).ForeColor = &H0&
            Textd(Index) = cTurnoAnt
            Option1(Index).Value = False
            Textd(Index).Locked = True:                  Textp(Index).Locked = True
            cmdGraba.SetFocus
            Exit Sub
         End If
           TxTurno = UCase(Textd(Index))
         If Trim(UCase(Textd(Index).Text)) = "V" Or Trim(UCase(Textd(Index).Text)) = "*" Then
            Textp(Index).Text = "O"
            TxHorario = "O"
         Else
            If Trim(UCase(TxTurno)) <> "A" Then
               Set rs_d = New Recordset
               sq_l = "SELECT * FROM CpcCostoHor WHERE Ccosto='" & Txccosto & "' And turno='" & Trim(Textd(Index).Text) & "'"
               rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not rs_d.EOF Then
                  Textp(Index).Text = rs_d!horario
                  TxHorario = rs_d!horario
               Else
                  Textp(Index).Text = ""
               End If
               rs_d.Close
            Else
               If Trim(UCase(TxTurno)) = "A" Then
                   Select Case TxPlH.Text
                     Case "CAS"
                        TxHorario = "10"
                     Case "HEP"
                     
                        If TxGrpGen <> "" Then
                            Set rs_d = New Recordset
                            sq_l = "SELECT * FROM CpGrupos WHERE Grupo='" & TxGrpGen & "'"
                            rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                            If Not rs_d.EOF Then
                                TxHorario = IIf(IsNull(rs_d!horario), "N", rs_d!horario)
                            End If
                            rs_d.Close
                        End If
                     
                        'If TxGrpGen = "G01" Then
                        '   TxHorario = "09"
                        'ElseIf TxGrpGen = "G03" Then
                         '      TxHorario = "10"
                        'End If
                   End Select
               End If
            End If
         End If
         '- Calcula acumulados
         Select Case Trim(Textd(Index).Text)
           Case "D"
              TxD.Text = Val(TxD.Text) + 1
           Case "D*"
              TxD.Text = Val(TxD.Text) + 1
           Case "D1"
              TxD.Text = Val(TxD.Text) + 1
           Case "D2"
              TxD.Text = Val(TxD.Text) + 1
           Case "D3"
              TxD.Text = Val(TxD.Text) + 1
              
           Case "MT"
              TxMT.Text = Val(TxMT.Text) + 1
              
           Case "N"
              TxN.Text = Val(TxN.Text) + 1
           Case "N*"
              TxN.Text = Val(TxN.Text) + 1
           Case "N2"
              TxN.Text = Val(TxN.Text) + 1
           Case "N3"
              TxN.Text = Val(TxN.Text) + 1
              
           Case "A"
              TxM.Text = Val(TxM.Text) + 1
           Case "M"
              TxM.Text = Val(TxM.Text) + 1
           Case "M1"
              TxM.Text = Val(TxM.Text) + 1
           Case "M2"
              TxM.Text = Val(TxM.Text) + 1
           Case "M3"
              TxM.Text = Val(TxM.Text) + 1
           Case "T"
              Txt.Text = Val(Txt.Text) + 1
           Case "T1"
              Txt.Text = Val(Txt.Text) + 1
           Case "T2"
              Txt.Text = Val(Txt.Text) + 1
              
            Case "H"
              TxM.Text = Val(TxM.Text) + 1
              TxN.Text = Val(TxN.Text) + 1
            Case "H1"
              TxM.Text = Val(TxM.Text) + 1
              TxN.Text = Val(TxN.Text) + 1
            Case "H*"
              TxM.Text = Val(TxM.Text) + 1
              TxN.Text = Val(TxN.Text) + 1
           Case "C"
              Txt.Text = Val(Txt.Text) + 1
              TxN.Text = Val(TxN.Text) + 1
            Case "C1"
              Txt.Text = Val(Txt.Text) + 1
              TxN.Text = Val(TxN.Text) + 1
            Case "C*"
              Txt.Text = Val(Txt.Text) + 1
              TxN.Text = Val(TxN.Text) + 1
            Case "P"
              TxD.Text = Val(TxD.Text) + 1
              TxN.Text = Val(TxN.Text) + 1
            Case "P1"
              TxD.Text = Val(TxD.Text) + 1
              TxN.Text = Val(TxN.Text) + 1
            Case "P*"
              TxD.Text = Val(TxD.Text) + 1
              TxN.Text = Val(TxN.Text) + 1
         End Select
         Select Case Trim(cTurnoAnt)
           Case "D"
              TxD.Text = Val(TxD.Text) - 1
           Case "D*"
              TxD.Text = Val(TxD.Text) - 1
           Case "D1"
              TxD.Text = Val(TxD.Text) - 1
           Case "D2"
              TxD.Text = Val(TxD.Text) - 1
           Case "D3"
              TxD.Text = Val(TxD.Text) - 1
              
           Case "MT"
              TxMT.Text = Val(TxMT.Text) - 1
              
           Case "N"
              TxN.Text = Val(TxN.Text) - 1
           Case "N2"
              TxN.Text = Val(TxN.Text) - 1
           Case "N3"
              TxN.Text = Val(TxN.Text) - 1
           Case "M"
              TxM.Text = Val(TxM.Text) - 1
           Case "M1"
              TxM.Text = Val(TxM.Text) - 1
           Case "M2"
              TxM.Text = Val(TxM.Text) - 1
           Case "M3"
              TxM.Text = Val(TxM.Text) - 1
              
           Case "A"
              TxM.Text = Val(TxM.Text) - 1
           Case "T"
              Txt.Text = Val(Txt.Text) - 1
           Case "T1"
              Txt.Text = Val(Txt.Text) - 1
           Case "T2"
              Txt.Text = Val(Txt.Text) - 1
              
            Case "H"
               TxM.Text = Val(TxM.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
            Case "H1"
               TxM.Text = Val(TxM.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
            Case "H*"
               TxM.Text = Val(TxM.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
            Case "C"
               Txt.Text = Val(Txt.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
            Case "C1"
               Txt.Text = Val(Txt.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
            Case "C*"
               Txt.Text = Val(Txt.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
            Case "P"
               TxD.Text = Val(TxD.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
            Case "P1"
               TxD.Text = Val(TxD.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
            Case "P*"
               TxD.Text = Val(TxD.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
         End Select
         If Val(TxN.Text) > GenMaxTN Then
            MsgBox "Exede Noches permitidas, Verifique Programa..!", vbExclamation
            cmdGraba.SetFocus
            Exit Sub
         End If
              Textp(Index).Text = TxHorario
         CboTurDir_KeyPress (13)
      Else
         Select Case Trim(cTurnoAnt)
           Case "D"
              TxD.Text = Val(TxD.Text) - 1
           Case "D*"
              TxD.Text = Val(TxD.Text) - 1
           Case "D1"
              TxD.Text = Val(TxD.Text) - 1
           Case "D2"
              TxD.Text = Val(TxD.Text) - 1
           Case "D3"
              TxD.Text = Val(TxD.Text) - 1
              
           Case "MT"
              TxMT.Text = Val(TxMT.Text) - 1
              
           Case "N"
              TxN.Text = Val(TxN.Text) - 1
           Case "N*"
              TxN.Text = Val(TxN.Text) - 1
           Case "N2"
              TxN.Text = Val(TxN.Text) - 1
           Case "N3"
              TxN.Text = Val(TxN.Text) - 1
              
           Case "A"
              TxM.Text = Val(TxM.Text) - 1
           Case "M"
              TxM.Text = Val(TxM.Text) - 1
           Case "M1"
              TxM.Text = Val(TxM.Text) - 1
           Case "M2"
              TxM.Text = Val(TxM.Text) - 1
           Case "M3"
              TxM.Text = Val(TxM.Text) - 1
           Case "T"
              Txt.Text = Val(Txt.Text) - 1
           Case "T1"
              Txt.Text = Val(Txt.Text) - 1
           Case "T2"
              Txt.Text = Val(Txt.Text) - 1
              
             Case "H"
               TxM.Text = Val(TxM.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
             Case "H*"
               TxM.Text = Val(TxM.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
             Case "H1"
               TxM.Text = Val(TxM.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
             Case "C"
               Txt.Text = Val(Txt.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
             Case "C*"
               Txt.Text = Val(Txt.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
             Case "C1"
               Txt.Text = Val(Txt.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
             Case "P"
               TxD.Text = Val(TxD.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
             Case "P*"
               TxD.Text = Val(TxD.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
             Case "P1"
               TxD.Text = Val(TxD.Text) - 1
               TxN.Text = Val(TxN.Text) - 1
         End Select
         TxHoras.Text = Val(TxHoras.Text) - Val(Lblhr(Index).Caption)
         Textd(Index).ForeColor = &H0&: Textp(Index).ForeColor = &H0&
         Option1(Index).Value = False
         Textd(Index).Locked = True:                  Textp(Index).Locked = True
         Textd(Index).Text = "":     Textp(Index).Text = "": Lblhr(Index).Caption = ""
         TxTurno.Text = "":   CboTurDir.Text = "":   CboTurDir.Enabled = False
         TxHorario.Text = "":   CboHorario.Text = "":    CboHorario.Enabled = False
         cmdGraba.SetFocus
      End If
End Sub
Private Sub CboTurDir_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboTurDir.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboTurDir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If IsEmpty(CboTurDir.Text) Then
     Else
        j = CboTurDir.ListIndex
        If j <> -1 Then
           nposicion = InStr(CboTurDir.List(j), ":")
           nlen = Len(CboTurDir.List(j))
           TxTurno.Text = Right(CboTurDir.List(j), nlen - nposicion)
           If InStr(Trim(TxTurno.Text), "A_*") = 0 Then
              '-Turno No Valido
              Textd(z).ForeColor = &H0&
              Option1(z).Value = False
              Textd(z).Locked = True:                  Textp(z).Locked = True
              TxTurno.Text = "":   CboTurDir.Text = "":   CboTurDir.Enabled = False
              TxHorario.Text = "":   CboHorario.Text = "":    CboHorario.Enabled = False
              Exit Sub
           End If
           dfi = Day(primer)
           dff = Day(ultimo)
           For i = dfi To dff
              z = i - 1
              If Option1(z).Value = True Then
                 Textd(z).Text = TxTurno.Text
                 Textd(z).ForeColor = &HFF0000
                 Textp(z).SetFocus
                 If Trim(TxHorario.Text) <> "" Then      ' "O"
                    Set rs_d = New Recordset
                    sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & TxHorarioext & "'"
                    rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                    If Not rs_d.EOF Then
                       rs_d.Close
                       TxHorario_KeyPress (13)
                    Else
                       rs_d.Close
                    End If
                 End If
                 Exit Sub
              End If
           Next
        End If
        Exit Sub
     End If
   End If
End Sub
Private Sub TxHorario_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     TxHorario = UCase(TxHorario)
     If TxHorario.Text <> "  " Then
        'NOMBRADOS Medicos D, N, T, M, 1, 4; Enf-tec D,N,T,J,M
        'CAS       Medicos D,N,T,M,1,4;     Enf-tec D,N,T,M
       For i = 0 To CboHorario.ListCount - 1
         npos = InStr(CboHorario.List(i), ">")
         nlen = Len(CboHorario.List(i))
         dato = Trim(Right(CboHorario.List(i), nlen - npos))
         If dato = TxHorario Then
            CboHorario.ListIndex = i
            Exit For
         End If
       Next i
       Cbohorario_KeyPress (13)
     End If
  End If
End Sub
Private Sub Textp_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
     Textp(Index) = UCase(Textp(Index))
     If Textp(Index).Text <> "  " Then
       For i = 0 To CboHorario.ListCount - 1
         npos = InStr(CboHorario.List(i), ">")
         nlen = Len(CboHorario.List(i))
         dato = Trim(Right(CboHorario.List(i), nlen - npos))
         If dato = Textp(Index).Text Then
            CboHorario.ListIndex = i
            Exit For
         End If
       Next i
       Cbohorario_KeyPress (13)
     End If
  End If
End Sub

Private Sub Cbohorario_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboHorario.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbohorario_KeyPress(KeyAscii As Integer)
   Sw = True
   If KeyAscii = 13 Then
     If IsEmpty(CboHorario.Text) Then
     Else
        j = CboHorario.ListIndex
        If j <> -1 Then
           nposicion = InStr(CboHorario.List(j), ">")
           nlen = Len(CboHorario.List(j))
           TxHorario.Text = Right(CboHorario.List(j), nlen - nposicion)
           ' TODOS V-O,*-O
           'NOMBRADOS Medicos D,N,T,M,1,4; Enf-tec
           'CAS       Medicos D,N,T,M,1,4;      Enf-tec
           If (Trim(UCase(TxTurno.Text)) = "V" Or Trim(UCase(TxTurno.Text)) = "*") And Trim(UCase(TxHorario.Text)) <> "O" Then
              Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "A" Then
                  Select Case TxPlH.Text
                    Case "CAS"
                      If InStr("10", Trim(UCase(TxHorario.Text))) = 0 Then
                         Sw = False
                      End If
                    Case "HEP"
                          If TxGrpGen <> "" Then
                            Set rs_d = New Recordset
                            sq_l = "SELECT * FROM CpGrupos WHERE Grupo='" & TxGrpGen & "'"
                            rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                            If Not rs_d.EOF Then
                               If rs_d!horario <> Trim(TxHorario) Then
                                  Sw = False
                               End If
                            Else
                                  Sw = False
                            End If
                            rs_d.Close
                        End If
                  End Select
           ElseIf Trim(UCase(TxTurno.Text)) = "D" And InStr("01_05", Trim(UCase(TxHorario.Text))) = 0 Then
               Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "D1" And InStr("11", Trim(UCase(TxHorario.Text))) = 0 Then
               Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "MT" And InStr("52_53", Trim(UCase(TxHorario.Text))) = 0 Then
               Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "N" And InStr("02_06_12", Trim(UCase(TxHorario.Text))) = 0 Then
                  Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "T" And InStr("04_08", Trim(UCase(TxHorario.Text))) = 0 Then
                  Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "M" And InStr("03_07", Trim(UCase(TxHorario.Text))) = 0 Then
                  Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "1" And InStr("01", Trim(UCase(TxHorario.Text))) = 0 Then
                  Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "4" And InStr("02", Trim(UCase(TxHorario.Text))) = 0 Then
                  Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "M1" And InStr("13", Trim(UCase(TxHorario.Text))) = 0 Then
                  Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "T1" And InStr("14", Trim(UCase(TxHorario.Text))) = 0 Then
                  Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "H" And InStr("15_16", Trim(UCase(TxHorario.Text))) = 0 Then
                  Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "C" And InStr("17_18", Trim(UCase(TxHorario.Text))) = 0 Then
                  Sw = False
           ElseIf Trim(UCase(TxTurno.Text)) = "P" And InStr("19_20", Trim(UCase(TxHorario.Text))) = 0 Then
                  Sw = False
           End If
           If Sw Then
           Else
              MsgBox "HORARIO no Corresponde...  ", vbExclamation
              Exit Sub
           End If
           dfi = Day(primer)
           dff = Day(ultimo)
           TxHoras.Text = ""     'calcula Horas Prog
           ndia = 0
           For i = dfi To dff
              z = i - 1
              If Option1(z).Value = True Then
                 ndia = z
                 Textp(z).Text = TxHorario
                 Textp(z).ForeColor = &HFF0000
                 Option1(z).Value = False
                 Textd(z).Locked = True:                  Textp(z).Locked = True
                 TxHorario = "":   CboHorario.Text = "":    CboHorario.Enabled = False
                 Lblhr(z).Caption = ""
                 TxTurno = "":  CboTurDir.Text = "":   CboTurDir.Enabled = False
                 'cmdGraba.SetFocus
              End If
              If Textp(z).Text <> "" Then
                 Set rs_d = New Recordset
                 sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & Trim(Textp(z)) & "'"     'CpHorarios
                 rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                 If Not rs_d.EOF Then
                    cvalor = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
                    TxHoras = Val(TxHoras) + Val(Left(cvalor, 2)) + (Val(Right(cvalor, 2)) / 60)
                    If Val(cvalor) > 0 Then
                       Lblhr(z).Caption = Left(cvalor, 2) + ":" + Right(cvalor, 2)
                    Else
                       Lblhr(z).Caption = ""
                    End If
                 End If
                 rs_d.Close
              End If
           Next
           
           cmdGraba.SetFocus
        End If
        Exit Sub
     End If
   End If
End Sub
Private Sub CmdLimpia_Click()
    Txtarjeta = "": Txtrabajador = "": Txfecnac = ""
    Txccosto = "": TxPlH = "": TxCarEst = "":  TxCargo = "": TxDscCargo = "": TxGrpGen = ""
    TxD = "": TxN = "": TxM = "": Txt = "": TxHoras = "":  TxMT = ""
    CboTurDir.Text = "": TxTurno = ""
    CboHorario.Text = "": TxHorario = ""
    LblONOMASTICO.Caption = "Onomastico"
    LblONOMASTICO.Visible = False
    For i = dfi To dff
       z = i - 1
       Option1(z).Value = False
       Textd(z).Visible = True: Textp(z).Visible = True
       Textd(z).Text = "": Textd(z).ForeColor = &H80000012  '&HFF0000
       Textp(z).Text = "":  Textp(z).ForeColor = &H80000012 '&HFF0000
       Lblhr(z).Caption = ""
    Next
End Sub
Private Sub Limpiacampos()
   Txtarjeta.Text = "": Txtrabajador.Text = "": Txfecnac.Text = "": TxGrpGen.Text = ""
   TxCarEst.Text = "": Txccosto.Text = "": TxPlH.Text = "": TxCargo.Text = "": TxDscCargo.Text = ""
   TxTurno.Text = "": CboTurDir.Text = ""
   TxHoras.Text = "":  TxHorario.Text = "": CboHorario.Text = ""
   TxD.Text = "": TxN.Text = "": TxM.Text = "": Txt.Text = "":  TxMT.Text = ""
   For i = 1 To 31
      z = i - 1
      Option1(z).Value = False
      Textd(z).Text = "":      Textd(z).ForeColor = &H80000008
      Textd(z).Visible = True
      Textp(z).Text = "":      Textp(z).ForeColor = &H80000008
      Lblhr(z).Caption = "":   Lblhr(z).ForeColor = &H80000008
   Next
End Sub

Private Sub mnuFicImpSelec_Click()
    ' Seleccionar la impresora a usar                               (23/Ene/00)
    ' La detección de errores es por si no hay impresora instalada
    On Error Resume Next
    With CommonDialog1
        .DialogTitle = "Seleccionar impresora"
        .Flags = cdlPDPrintSetup
        .ShowPrinter
    End With
    Err = 0
End Sub
Private Sub cmdNuevo_Click()
   For i = 1 To 31
      z = i - 1
      Textd(z).Text = "":      Textd(z).ForeColor = &H80000008
      Textp(z).Text = "":      Textp(z).ForeColor = &H80000008
      Lblhr(z).Caption = ""
   Next
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   'DataGrid1.Row = 0
   'DataGrid1.Col = 0
End Sub
Private Sub cmdGraba_Click()
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If IsNull(Txmes) Then
      MsgBox "defina PERIODO...  ", vbExclamation
      Exit Sub
   End If
   If IsNull(Txtarjeta) Then   'And IsNull(Txgrupo.Text)
      MsgBox "defina trabajador...  ", vbExclamation
      Exit Sub
   End If
   If Val(TxN) > GenMaxTN Then
      MsgBox "Excede " + Trim(Str(GenMaxTN)) + " Noches permitidas, Verifique Programa..!", vbExclamation
      Exit Sub
   End If
   
   'Valida Max dias Libres  GENMAXDL
   nposicion = 0
   cvalor = ""
   nhoras = 0
   For i = Day(primer) To Day(ultimo)
       z = i - 1
       If Val(Lblhr(z).Caption) = 0 And (Textd(z) = "*" Or Trim(Textd(z)) = "") Then
          nposicion = nposicion + 1
          cvalor = cvalor + Trim(Str(i)) + "-"
       ElseIf Val(Lblhr(z).Caption) > 0 And InStr("ADD1D2D3_D*_MM1M2M3_TT1T2_NN*N2N3_V_L_HCP_H*H1_P*P1_C*C1_MT", Textd(z)) > 0 And (Textd(z) <> "*" Or Trim(Textd(z)) <> "") Then
       '-ElseIf Val(Lblhr(z).Caption) > 0 And InStr("A_V_L", Trim(Textd(z))) > 0 And (Textd(z) <> "*" Or Trim(Textd(z)) <> "") Then
           nposicion = 0
           cvalor = ""
       End If
       If nposicion > GenMaxDL Then
          MsgBox "Verifique BLOQUE de Dias (" + Trim(Str(nposicion)) + "), Libres.. " + cvalor, vbExclamation
          Exit Sub
       End If
   Next i
   On Error GoTo final
   Cnx.BeginTrans
   '-Actualiza DIRECTORIO "TURNO"
   Sw = False
   If Txtarjeta <> "" Then
      TxHoras = nhoras
      
      sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      If rs_d.EOF Then
         Sw = True
         rs_d.AddNew
         rs_d!tarjeta = Txtarjeta
         rs_d!ano = genaño
         rs_d!mes = Txmes
         rs_d!usuario = wUsuario      'Firma Usuario
         rs_d!fecha = Date
      End If
      For i = Day(primer) To Day(ultimo)
          z = i - 1
          'For i = 0 To rs.Fields.Count - 1
          'cCampo = rs.Fields(z + 4).Name
          rs_d.Fields(z + 4) = IIf(IsNull(Textd(z)), "", Textd(z))
      Next i
      rs_d!carest = IIf(IsNull(TxCarEst), "", TxCarEst)
      rs_d!cCosto = IIf(IsNull(Txccosto), "", Txccosto)
      rs_d!cargo = IIf(IsNull(TxCargo), "", TxCargo)
      rs_d!d = IIf(IsNull(TxD), "", TxD)
      rs_d!n = IIf(IsNull(TxN), "", TxN)
      rs_d!m = IIf(IsNull(TxM), "", TxM)
      rs_d!T = IIf(IsNull(Txt), "", Txt)
      rs_d!MT = IIf(IsNull(TxMT), "", TxMT)
      rs_d!Horas = IIf(IsNull(TxHoras), 0, Val(TxHoras))
      rs_d!Flag = "1"       'ADMINISTRATIVO
      rs_d!plh = IIf(IsNull(TxPlH), "", TxPlH)
      rs_d!usuarioM = wUsuario      'Firma Usuario
      rs_d!fechaM = Date
      rs_d.Update
      rs_d.Close
      If Sw Then
         DD.AddNew
      End If
      DD!tarjeta = Txtarjeta
      DD!nombre = Left(Txtrabajador, 25)
      DD!cargo = Left(TxDscCargo, 15)
      DD!d = Val(TxD)
      DD!m = Val(TxM)
      DD!T = Val(Txt)
      DD!n = Val(TxN)
      DD!Horas = Val(TxHoras)
      For i = dfi To dff
        z = i - 1
        DD.Fields(z + 8) = IIf(IsNull(Textd(z)), "", Textd(z))
      Next i
      DD!MT = Val(TxMT)
      DD.Update
   End If
   '-Actualiza Programa de HORARIOS del grupo
   For i = dfi To dff
      z = i - 1
      f_echa = Format(primer + z, "yyyymmdd")
      If Textp(z).Text <> "" Then
         Set rs = New Recordset
         If Txtarjeta.Text <> "" Then
            sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & f_echa & "'"
         End If
         rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If rs.EOF Then
            rs.AddNew
            rs!tarjeta = IIf(IsNull(Txtarjeta), "", Txtarjeta)
            rs!fecha = primer + z
            rs!usuario = wUsuario      'Firma Usuario
         End If
         rs!horario = Textp(z).Text
         rs.Update
         rs.Close
      Else
         If Txtarjeta <> "" Then
            Cnx.Execute "DELETE FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "'AND fecha ='" & f_echa & "'"
         End If
      End If
   Next
   Cnx.CommitTrans
   Call Limpiacampos
   Frame1.Caption = "Programa"
   DataGrid1.Refresh
   GoTo Final1
final:
   MsgBox Err.Description & " Error: " & Err.Number & vbCr
   Cnx.RollbackTrans
Final1:
End Sub
Private Sub CmdElimina_Click()
    If wrights = "CON" Then
       MsgBox "No Autorizado...  ", vbExclamation
       Exit Sub
    End If
    If (Txmes.Text <> "" And IsNull(Txtarjeta.Text)) Then Exit Sub
    '-Valida retiro
    Set rt = New Recordset
    sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
    rt.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    If Not rt.EOF Then
       ndia = 0
       For i = Day(primer) To Day(ultimo)
          z = i - 1
          TxTurno = Trim(IIf(IsNull(rt.Fields(z + 4)), "", rt.Fields(z + 4)))
          If TxTurno <> "" Then
             ndia = ndia + IIf(InStr("L_V_R", TxTurno) = 0, 0, 1)
          End If
       Next i
       rt.Close
       If ndia > 0 Then
          MsgBox "<E> Existe Autorizaciones previas.." + Str(ndia), vbExclamation
          Exit Sub
       End If
    Else
       rt.Close
       Exit Sub
    End If
    If MsgBox("Eliminar Directorio y Programa del mes ?      ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       Fecha1 = Format(primer, "yyyymmdd")
       fecha2 = Format(ultimo, "yyyymmdd")
       If Txtarjeta.Text <> "" Then
          Cnx.Execute "DELETE FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "'AND fecha >='" & Fecha1 & "' AND fecha<='" & fecha2 & "'"
          '-Actualiza DIRECTORIO "TURNO"
          Cnx.Execute "DELETE FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
       End If
       cBuscar = "tarjeta like '" & Txtarjeta & "'"
       DD.MoveFirst
       DD.Find (cBuscar)
       If Not DD.EOF Then
          DD.Delete
       End If
    End If
    Call Limpiacampos
 End Sub
Private Sub Cmdsalir_Click()
   Set rs_d = Nothing
   Set rs = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        '-Form Abierto
        If UCase(Forms(cont).Name) <> "CPMENU" And Forms(cont).Name <> "PROGADMINISTRATIVO" Then
           Unload Forms(cont)
        End If
      End If
   Next
   Me.Hide
   Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set rs_d = Nothing
   Set rs = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        '- Form Abierto
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "ProgAdministrativo" Then
           Unload Forms(cont)
        End If
      End If
   Next
   Set r_d = Nothing
   Me.Hide
   Unload Me
End Sub
Private Sub Cargadias()
   For i = 1 To 31
      z = i - 1
      Textd(z).Text = "":      Textd(z).ForeColor = &H80000008
      Textp(z).Text = "":      Textp(z).ForeColor = &H80000008
      Lblhr(z).Caption = ""
      Textd(z).Visible = False
      Textp(z).Visible = False
      Option1(z).Visible = False
   Next
   unafecha = primer
   dfi = Day(primer)
   dff = Day(ultimo)
   
   ccadena = ""    'Ubica Feriados/Domingo
   For i = dfi To dff
       f_Fer = False
       z = i - 1
       Textd(z).Visible = True
       Textp(z).Visible = True
       Option1(z).Enabled = True
       Option1(z).Visible = True
       Option1(z).Caption = Left(WeekdayName(Weekday(unafecha + i - 1)), 3) + " " & i
       Option1(z).ForeColor = &H80000012      'negro
       If Weekday(unafecha + i - 1) = 1 Then
          '-UCase(Left(WeekdayName(Weekday(unafecha + i - 1)), 3)) = "DOM"
          Option1(z).ForeColor = &HFF               'rojo
          f_Fer = True
       Else
          Set rs = New Recordset
          rs.Open "SELECT * FROM Feriados ORDER BY codigo", Cnx0, adOpenStatic, adLockReadOnly
          'rs.Open "SELECT * FROM CpTablasAux WHERE grupo='FER' ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
          rs.MoveFirst
          While Not rs.EOF
             If Len(Trim(rs!Codigo)) = 6 Then
                c_a = Right(rs!Codigo, 2)
             ElseIf Len(Trim(rs!Codigo)) = 4 Then
                c_a = Right(Year(unafecha), 2)
             End If
             If Val(Left(rs!Codigo, 2)) = Month(unafecha) And Val(Mid(rs!Codigo, 3, 2)) = i And Right(Year(unafecha), 2) = c_a Then
                Option1(z).ForeColor = &HFF       'rojo
                f_Fer = True
             End If
             rs.MoveNext
          Wend
          rs.Close
       End If
       ccadena = ccadena + IIf(f_Fer, "F", "O")
       Textp(z).Enabled = True
   Next
End Sub
Private Sub CmdPrint_Click()
   If Txmes <> "" And Tx_ccosto <> "" Then
      '-Programa mensual CCosto
      CargaProgCC
      nTipRpt = 11.2
      InfoCpRS.Left = 0: InfoCpRS.Top = 0
      InfoCpRS.Show
      Call Limpiacampos
      CmdPrint.Caption = ""
   End If
   Exit Sub
End Sub
Private Sub CargaProgCC()
'***  CARGA  CC-MES
    On Error Resume Next
    Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmpDCosto"
    '-USO TEMPORAL
     Cnx.Execute ("SELECT * Into #Tabla  FROM CpTmpDCosto WHERE 1 = 2")
     Cnx.Execute "DELETE  FROM #tabla"
     Cnx.CommitTrans
    Set r_d = New ADODB.Recordset
    nposicion = 0
    '-Directorio Mensual de CC
    sql = "SELECT  * FROM CpDirectorio WHERE ano='" & genaño & "' and mes='" & Txmes.Text & "' and ccosto='" & Tx_ccosto.Text & "' and flag='1'" & _
    "ORDER BY ccosto,Tarjeta "
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Set rs_d = New Recordset
    'SQ_L = "SELECT * FROM CpTmpDcosto"
      sq_l = "SELECT * FROM #Tabla"
    rs_d.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    unafecha = "01/" + Trim(r_d!mes) + "/" + Trim(Str(genaño))
    Call ObtenerPrimerUltimoDia(unafecha)
    unafecha = primer
    Do While Not r_d.EOF
       rs_d.AddNew
       cTarjeta = r_d!tarjeta
       CmdPrint.Caption = cTarjeta
       Set rs = New Recordset
       rs.Open "SELECT * FROM cpgeneral WHERE tarjeta='" & cTarjeta & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
       cNombre = ""
       ndia = 0
       If Not rs.EOF Then
          cNombre = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 30)
          If IsDate(rs!fecnac) Then
             If Month(rs!fecnac) = Val(Txmes.Text) Then
                ndia = Day(rs!fecnac)
             End If
          End If
       End If
       rs.Close
       rs_d!tarjeta = cTarjeta
       rs_d!ano = r_d!ano
       rs_d!mes = r_d!mes
       cCarest = IIf(IsEmpty(r_d!carest), "", r_d!carest)
       ccCosto = IIf(IsEmpty(r_d!cCosto), "", r_d!cCosto)
       ccargo = IIf(IsEmpty(r_d!cargo), "", r_d!cargo)
       rs_d!cargo = ccargo
       rs_d!cCosto = ccCosto
       For i = Day(primer) To Day(ultimo)
          z = i - 1
          If ndia = i Then   'Onomastico
             rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "¹", "¹" + r_d.Fields(z + 4))
          Else
             rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
          End If
          rs_d.Fields(z + 6 + 31) = Left(WeekdayName(Weekday(unafecha + i - 1)), 1)   'LU/MA/MI/JU/VI/SA/DO
       Next i
       rs_d!nombre = cNombre
       If ccargo <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM cpcargos WHERE codigo='" & ccargo & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If rs.EOF Then
          Else
             rs_d!DscCargo = IIf(IsNull(rs!descripcion), "", Left(rs!descripcion, 30))
          End If
          rs.Close
       End If
       rs_d!d = IIf(IsEmpty(r_d!d), "", r_d!d)
       rs_d!n = IIf(IsEmpty(r_d!n), "", r_d!n)
       rs_d!m = IIf(IsEmpty(r_d!m), "", r_d!m)
       rs_d!T = IIf(IsEmpty(r_d!T), "", r_d!T)
       rs_d!MT = IIf(IsEmpty(r_d!MT), "", r_d!MT)   '<=====
       rs_d!Horas = IIf(IsEmpty(r_d!Horas), 0, r_d!Horas)
       rs_d!Dia = ndia
       rs_d.Update
       r_d.MoveNext
    Loop
    'Cnx.CommitTrans
    rs_d.Close
    r_d.Close
    Exit Sub
End Sub

