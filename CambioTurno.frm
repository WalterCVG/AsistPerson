VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CambioTurno 
   Caption         =   "Boleta Cambio de Turno"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "CambioTurno.frx":0000
   ScaleHeight     =   8175
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Tx_Cargo 
      Enabled         =   0   'False
      Height          =   195
      Left            =   10080
      TabIndex        =   205
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Tx_DscCargo 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   204
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox Tx_Plh 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   8280
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   194
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   13320
      Picture         =   "CambioTurno.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   193
      ToolTipText     =   "Limpia seleccion"
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton CmdLIMPIA 
      Caption         =   "Limpia"
      Height          =   615
      Left            =   13920
      Picture         =   "CambioTurno.frx":80B30
      Style           =   1  'Graphical
      TabIndex        =   192
      ToolTipText     =   "Limpia Plantilla"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Tx_codupr 
      Enabled         =   0   'False
      Height          =   195
      Left            =   14160
      MaxLength       =   2
      TabIndex        =   190
      Top             =   6120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Cbo_UPR 
      Height          =   315
      Left            =   11880
      TabIndex        =   189
      Top             =   6360
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13080
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton mnuFicImpSelec 
      Height          =   495
      Left            =   12480
      Picture         =   "CambioTurno.frx":81062
      Style           =   1  'Graphical
      TabIndex        =   186
      ToolTipText     =   "Selec.Impresora"
      Top             =   6720
      Width           =   615
   End
   Begin VB.CommandButton cmdchgreg 
      Height          =   615
      Left            =   13320
      Picture         =   "CambioTurno.frx":81840
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Ubicar Registros"
      Top             =   480
      Width           =   735
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   1800
      TabIndex        =   30
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   195
      Left            =   3840
      MaxLength       =   2
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Directorio"
      Height          =   2055
      Left            =   0
      TabIndex        =   27
      Top             =   2640
      Width           =   15015
      Begin VB.TextBox TxPA 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   12480
         TabIndex        =   207
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox TxNA 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   12120
         TabIndex        =   206
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox TxMt 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   11760
         TabIndex        =   201
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   600
         MaxLength       =   4
         TabIndex        =   171
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   170
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   169
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   168
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   167
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   6960
         MaxLength       =   4
         TabIndex        =   166
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   8280
         MaxLength       =   4
         TabIndex        =   165
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   9480
         MaxLength       =   4
         TabIndex        =   164
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   10800
         MaxLength       =   4
         TabIndex        =   163
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   12120
         MaxLength       =   4
         TabIndex        =   162
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   13320
         MaxLength       =   4
         TabIndex        =   161
         Text            =   "Text1"
         Top             =   480
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   600
         MaxLength       =   4
         TabIndex        =   160
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   159
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   158
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   157
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   156
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   6960
         MaxLength       =   4
         TabIndex        =   155
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         Left            =   8280
         MaxLength       =   4
         TabIndex        =   154
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   18
         Left            =   9480
         MaxLength       =   4
         TabIndex        =   153
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   19
         Left            =   10800
         MaxLength       =   4
         TabIndex        =   152
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   20
         Left            =   12120
         MaxLength       =   4
         TabIndex        =   151
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   21
         Left            =   13320
         MaxLength       =   4
         TabIndex        =   150
         Text            =   "Text1"
         Top             =   1080
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   22
         Left            =   600
         MaxLength       =   4
         TabIndex        =   149
         Text            =   "Text1"
         Top             =   1680
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   23
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   148
         Text            =   "Text1"
         Top             =   1680
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   24
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   147
         Text            =   "Text1"
         Top             =   1680
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   25
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   146
         Text            =   "Text1"
         Top             =   1680
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   26
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   145
         Text            =   "Text1"
         Top             =   1680
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   27
         Left            =   6960
         MaxLength       =   4
         TabIndex        =   144
         Text            =   "Text1"
         Top             =   1680
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   28
         Left            =   8280
         MaxLength       =   4
         TabIndex        =   143
         Text            =   "Text1"
         Top             =   1680
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   29
         Left            =   9480
         MaxLength       =   4
         TabIndex        =   142
         Text            =   "Text1"
         Top             =   1680
         Width           =   355
      End
      Begin VB.TextBox Textp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   30
         Left            =   10800
         MaxLength       =   4
         TabIndex        =   141
         Text            =   "Text1"
         Top             =   1680
         Width           =   355
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   28
         Left            =   7920
         TabIndex        =   137
         Top             =   1440
         Width           =   855
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
         Left            =   14385
         Locked          =   -1  'True
         TabIndex        =   130
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox TxD 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   12840
         Locked          =   -1  'True
         TabIndex        =   129
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox TxN 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   128
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox TxM 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   13560
         Locked          =   -1  'True
         TabIndex        =   127
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox Txt 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   13920
         Locked          =   -1  'True
         TabIndex        =   126
         Top             =   1680
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   94
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   30
         Left            =   10440
         MaxLength       =   2
         TabIndex        =   93
         Text            =   "Text31"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Textd 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   29
         Left            =   9120
         MaxLength       =   2
         TabIndex        =   92
         Text            =   "Text30"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   28
         Left            =   7920
         MaxLength       =   2
         TabIndex        =   91
         Text            =   "Text29"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   27
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   90
         Text            =   "Text28"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   26
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   89
         Text            =   "Text27"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   25
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   88
         Text            =   "Text26"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   24
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   87
         Text            =   "Text25"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   23
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   86
         Text            =   "Text24"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   22
         Left            =   240
         MaxLength       =   2
         TabIndex        =   85
         Text            =   "Text23"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   21
         Left            =   12960
         MaxLength       =   2
         TabIndex        =   84
         Text            =   "Text22"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   20
         Left            =   11760
         MaxLength       =   2
         TabIndex        =   83
         Text            =   "Text21"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   19
         Left            =   10440
         MaxLength       =   2
         TabIndex        =   82
         Text            =   "Text20"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   18
         Left            =   9120
         MaxLength       =   2
         TabIndex        =   81
         Text            =   "Text19"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   17
         Left            =   7920
         MaxLength       =   2
         TabIndex        =   80
         Text            =   "Text18"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   16
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   79
         Text            =   "Text17"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   15
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   78
         Text            =   "Text16"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   14
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   77
         Text            =   "Text15"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   13
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   76
         Text            =   "Text14"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   12
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   75
         Text            =   "Text13"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   11
         Left            =   240
         MaxLength       =   2
         TabIndex        =   74
         Text            =   "Text12"
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   10
         Left            =   12960
         MaxLength       =   2
         TabIndex        =   73
         Text            =   "Text11"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   9
         Left            =   11760
         MaxLength       =   2
         TabIndex        =   72
         Text            =   "Text10"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   8
         Left            =   10440
         MaxLength       =   2
         TabIndex        =   71
         Text            =   "Text9"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   7
         Left            =   9120
         MaxLength       =   2
         TabIndex        =   70
         Text            =   "Text8"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   6
         Left            =   7920
         MaxLength       =   2
         TabIndex        =   69
         Text            =   "Text7"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   5
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   68
         Text            =   "Text6"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   4
         Left            =   5280
         MaxLength       =   2
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   3
         Left            =   4080
         MaxLength       =   2
         TabIndex        =   66
         Text            =   "Text4"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   2
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   65
         Text            =   "Text3"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   1
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   64
         Text            =   "Text2"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   2
         TabIndex        =   63
         Text            =   "Text1"
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   62
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   61
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   3
         Left            =   4080
         TabIndex        =   60
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   59
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   5
         Left            =   6600
         TabIndex        =   58
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   7920
         TabIndex        =   57
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   7
         Left            =   9120
         TabIndex        =   56
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   8
         Left            =   10440
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   9
         Left            =   11760
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   10
         Left            =   12960
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   52
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   12
         Left            =   1440
         TabIndex        =   51
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   13
         Left            =   2760
         TabIndex        =   50
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   14
         Left            =   4080
         TabIndex        =   49
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   15
         Left            =   5280
         TabIndex        =   48
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   16
         Left            =   6600
         TabIndex        =   47
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   17
         Left            =   7920
         TabIndex        =   46
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   18
         Left            =   9120
         TabIndex        =   45
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   19
         Left            =   10440
         TabIndex        =   44
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   20
         Left            =   11760
         TabIndex        =   43
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   21
         Left            =   12960
         TabIndex        =   42
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   22
         Left            =   240
         TabIndex        =   41
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   23
         Left            =   1440
         TabIndex        =   40
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   24
         Left            =   2760
         TabIndex        =   39
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   25
         Left            =   4080
         TabIndex        =   38
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   26
         Left            =   5280
         TabIndex        =   37
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   27
         Left            =   6600
         TabIndex        =   36
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   29
         Left            =   9120
         TabIndex        =   35
         Top             =   1440
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   30
         Left            =   10440
         TabIndex        =   34
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label26 
         Caption         =   "PA"
         Height          =   255
         Left            =   12480
         TabIndex        =   210
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label25 
         Caption         =   "NA"
         Height          =   255
         Left            =   12120
         TabIndex        =   209
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Horas"
         Height          =   255
         Left            =   14280
         TabIndex        =   208
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MT"
         Height          =   255
         Left            =   11760
         TabIndex        =   202
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "T"
         Height          =   255
         Left            =   13920
         TabIndex        =   136
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   15
         Left            =   0
         TabIndex        =   135
         Top             =   0
         Width           =   135
      End
      Begin VB.Label Label17 
         Caption         =   "HORAS "
         Height          =   255
         Left            =   15240
         TabIndex        =   134
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "D"
         Height          =   255
         Left            =   12840
         TabIndex        =   133
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "N"
         Height          =   255
         Left            =   13200
         TabIndex        =   132
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "M"
         Height          =   255
         Left            =   13560
         TabIndex        =   131
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   125
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   124
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3480
         TabIndex        =   123
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4800
         TabIndex        =   122
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6000
         TabIndex        =   121
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   7320
         TabIndex        =   120
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   8640
         TabIndex        =   119
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   9840
         TabIndex        =   118
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   11160
         TabIndex        =   117
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   12480
         TabIndex        =   116
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   13680
         TabIndex        =   115
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   960
         TabIndex        =   114
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   2160
         TabIndex        =   113
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   3480
         TabIndex        =   112
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   4800
         TabIndex        =   111
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   6000
         TabIndex        =   110
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   7320
         TabIndex        =   109
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   8640
         TabIndex        =   108
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   9840
         TabIndex        =   107
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   11160
         TabIndex        =   106
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   12480
         TabIndex        =   105
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   13680
         TabIndex        =   104
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   22
         Left            =   1080
         TabIndex        =   103
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   23
         Left            =   2160
         TabIndex        =   102
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   3480
         TabIndex        =   101
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   25
         Left            =   4800
         TabIndex        =   100
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   6000
         TabIndex        =   99
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   27
         Left            =   7320
         TabIndex        =   98
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   28
         Left            =   8640
         TabIndex        =   97
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   29
         Left            =   9840
         TabIndex        =   96
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Lblhr 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   30
         Left            =   11160
         TabIndex        =   95
         Top             =   1680
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   14160
      Picture         =   "CambioTurno.frx":81A1B
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cerrar"
      Top             =   7320
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "[Registro]"
      Height          =   2655
      Left            =   480
      TabIndex        =   9
      Top             =   4920
      Width           =   11175
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   7440
         TabIndex        =   198
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7440
         TabIndex        =   197
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Txplh 
         BorderStyle     =   0  'None
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   8640
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   196
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Txcodupr 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         MaxLength       =   2
         TabIndex        =   185
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxDeTur 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   184
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox TxDetur0 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   183
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox TxAhrs 
         Enabled         =   0   'False
         Height          =   195
         Left            =   9600
         TabIndex        =   182
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TxAhrs0 
         Enabled         =   0   'False
         Height          =   195
         Left            =   9600
         TabIndex        =   181
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TxDeHrs 
         Enabled         =   0   'False
         Height          =   195
         Left            =   4800
         TabIndex        =   180
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxdeHrs0 
         Enabled         =   0   'False
         Height          =   195
         Left            =   4800
         TabIndex        =   179
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxAtur 
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   9120
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   178
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox TxAtur0 
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   9120
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   177
         Top             =   690
         Width           =   375
      End
      Begin VB.ComboBox CboAHor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7080
         TabIndex        =   176
         Top             =   1560
         Width           =   3975
      End
      Begin VB.ComboBox CboDeHor 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   175
         Top             =   1560
         Width           =   3375
      End
      Begin VB.ComboBox CboAhor0 
         Height          =   315
         Left            =   7080
         TabIndex        =   174
         Top             =   960
         Width           =   3975
      End
      Begin VB.ComboBox Cbodehor0 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         TabIndex        =   173
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox Txtipo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8040
         MaxLength       =   1
         TabIndex        =   172
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxAhor 
         Enabled         =   0   'False
         Height          =   195
         Left            =   10080
         MaxLength       =   4
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxdeHor 
         Enabled         =   0   'False
         Height          =   195
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Txafecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TxdeHor0 
         Enabled         =   0   'False
         Height          =   195
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxDocumento 
         Height          =   285
         Left            =   8400
         MaxLength       =   30
         TabIndex        =   26
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox Txdefecha 
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Txtarjeta 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Txtrabajador 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   14
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox TxAHor0 
         Enabled         =   0   'False
         Height          =   195
         Left            =   10080
         MaxLength       =   4
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Txcargo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Txccosto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         MaxLength       =   2
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxCarEst 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         MaxLength       =   2
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Txtipopla 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         MaxLength       =   3
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "JG"
         Height          =   255
         Left            =   7080
         TabIndex        =   200
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label20 
         Caption         =   "JG"
         Height          =   255
         Left            =   7080
         TabIndex        =   199
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "A :"
         Height          =   255
         Left            =   6840
         TabIndex        =   140
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "de :"
         Height          =   255
         Left            =   2520
         TabIndex        =   139
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha  Nueva "
         Height          =   255
         Left            =   120
         TabIndex        =   138
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "A :"
         Height          =   255
         Left            =   6840
         TabIndex        =   33
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "de :"
         Height          =   255
         Left            =   2520
         TabIndex        =   28
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "Documento"
         Height          =   255
         Left            =   7200
         TabIndex        =   24
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha BASE  "
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Trabajador"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
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
      Left            =   13320
      Picture         =   "CambioTurno.frx":81B65
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir"
      Top             =   6720
      Width           =   765
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "Graba"
      Height          =   615
      Left            =   12000
      Picture         =   "CambioTurno.frx":8242F
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Grabar"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox Tx_trabajador 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   4935
   End
   Begin VB.TextBox Tx_tarjeta 
      Height          =   285
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox TxDia 
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
      Left            =   7200
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   0
      Picture         =   "CambioTurno.frx":825A1
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   1560
      TabIndex        =   6
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2990
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
      Caption         =   "Registro Cambios"
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
   Begin VB.Label LblDS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LblDS"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   11640
      TabIndex        =   211
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "CARGO:"
      Height          =   255
      Left            =   9120
      TabIndex        =   203
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Plh"
      Height          =   255
      Left            =   7920
      TabIndex        =   195
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "U.P.R."
      Height          =   255
      Left            =   11880
      TabIndex        =   191
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo U.P.R."
      Height          =   375
      Left            =   0
      TabIndex        =   188
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo U.P.R."
      Height          =   375
      Left            =   0
      TabIndex        =   187
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   960
      TabIndex        =   31
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label_2 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "CambioTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DD As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim r_t As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim Sw, FlgExcep, FlgXGuar As Boolean
Dim nson, nSonD, nSonM, nSonT, nSonN, nposicion, ndifdias, nmesA, nmesI, nmesF, nano As Integer
Dim ceditor, sql, SqC, cmes, Chor, cvalor, Zde, ZA, msgExcep, TarjOk, cTrab As String
Dim ntHoras, nhrFin, nhrIni As Double
Dim dfi, dff As Integer
Dim fecCierre As Date
Dim unafecha As Date
Dim l_ok As String
Private Sub Command1_Click()
    c_grupo = "02": c_secuencia = "025"
    Load AydAcerca
    AydAcerca.Left = 0: AydAcerca.Top = 0
    AydAcerca.Show
End Sub
Private Sub Form_Load()
  Me.Caption = "CAMBIO de TURNO - " & genaño & "  -   Sesión: " + wUsuNombre
  LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
  For i = 1 To 31
     z = i - 1
     Option1(z).Value = False
     Textd(z).Text = "":   Textd(z).ForeColor = &H0&
     Textp(z).Text = "":   Textp(z).ForeColor = &H0&
     Lblhr(z).Caption = "":    Lblhr(z).ForeColor = &H0&
     Textd(z).Locked = True
     Textp(z).Locked = True
  Next
  Set rs = New Recordset
  rs.Open "SELECT * FROM cptablas WHERE tabla='32' and codigo<>'..' ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboMes.ListCount - 1
           If CboMes.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboMes.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
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
  rs.Open "SELECT * FROM CpUndPrgRol ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To Cbo_UPR.ListCount - 1
         If Cbo_UPR.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        Cbo_UPR.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR")
     End If
     rs.MoveNext
  Wend
  rs.Close
  If Wccosto <> "" Then
     CmdPrint.Enabled = False
     '-Ubica titular UPR
     If wCadenaUpr <> "" Then
        Tx_codupr.Text = Left(wCadenaUpr, 2)
        For i = 0 To Cbo_UPR.ListCount - 1
           npos = InStr(Cbo_UPR.List(i), ":")
           nlen = Len(Cbo_UPR.List(i))
           dato = Trim(Right(Cbo_UPR.List(i), nlen - npos))
           If dato = Trim(Tx_codupr.Text) Then
              Cbo_UPR.ListIndex = i
              CmdPrint.Enabled = True
              Exit For
           End If
        Next i
     End If
  End If
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM cphorarionew  ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To Cbodehor0.ListCount - 1
           If Cbodehor0.List(i) = rs.Fields("Descripcion") & "   =>" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          Cbodehor0.AddItem rs.Fields("Descripcion") & "    =>" & rs.Fields("codigo")
          CboAhor0.AddItem rs.Fields("Descripcion") & "    =>" & rs.Fields("codigo")
          CboDeHor.AddItem rs.Fields("Descripcion") & "    =>" & rs.Fields("codigo")
          CboAHor.AddItem rs.Fields("Descripcion") & "    =>" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  Set DD = New Recordset
  With DD
     .Fields.Append "tarjeta", adChar, 5
     .Fields.Append "Fecha", adDate
     .Fields.Append "dehor0", adChar, 4
     .Fields.Append "aHor0", adChar, 4
     .Fields.Append "NFecha", adDate
     .Fields.Append "DeHor", adChar, 4
     .Fields.Append "AHor", adChar, 4
     .Fields.Append "tipo", adChar, 1
     .Fields.Append "Documento", adChar, 30
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
        Select Case i
          Case 0
          '.Caption = "Tarjeta"
          .Width = Me.TextWidth(String(8, "S"))
          Case 1
          '.Caption = "fecha"
          .Width = Me.TextWidth(String(10, "S"))
          Case 2
          '.Caption ="dehor0"
          .Width = Me.TextWidth(String(6, "S"))
          Case 3
          '.Caption ="ahor0"
          .Width = Me.TextWidth(String(6, "S"))
        End Select
        '.NumberFormat = "##,##0.0"
        .Locked = True
     End With
  Next
  DataGrid1.Enabled = False
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
      Else
         j = CboMes.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboMes.List(j), ":")
            nlen = Len(CboMes.List(j))
            Txmes = Right(CboMes.List(j), nlen - nposicion)
            unafecha = "01/" + Txmes + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)
            Call Cargadias       'primer/ultimo
         End If
      End If
      fecCierre = ultimo
      
      Set rs = New Recordset
      rs.Open "SELECT * FROM CpCronograma WHERE ano='" & genaño & "' and Mes='" & Txmes.Text & "'", Cnx, adOpenStatic, adLockReadOnly   'CpHorarios
      If rs.EOF Then
         MsgBox "No Existe CRONOGRAMA del MES... Verifique", vbExclamation
      Else
         If IsDate(rs!cierremes) Then fecCierre = rs!cierremes
         If fecCierre <= Date Then
            Beep 500, 350
            Me.Caption = "Cambio de Turno, Programa CERRADO. - " & genaño & "  -   Sesión: " + wUsuNombre
         Else
            Me.Caption = "Cambio de Turno - " & genaño & "  -   Sesión: " + wUsuNombre
         End If
      End If
      rs.Close
      Exit Sub
   End If
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
   For i = dfi To dff
       z = i - 1
       Option1(z).Enabled = True
       Option1(z).Visible = True
       Option1(z).Caption = Left(WeekdayName(Weekday(primer + z)), 3) + " " & i
       Option1(z).ForeColor = &H80000012      'negro
       '''If Weekday(unafecha + i - 1) = 1 Then
       If UCase(Left(WeekdayName(Weekday(primer + z)), 3)) = "DOM" Then
          Option1(z).ForeColor = &HFF               'rojo
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
                Option1(z).ForeColor = &HFF
             End If
             rs.MoveNext
          Wend
          rs.Close
       End If
       Textd(z).Visible = True
       Textp(z).Visible = True
   Next
End Sub

Private Sub Tx_tarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Tx_tarjeta.Text) = "" Or (Trim(Tx_tarjeta.Text) <> "" And EsTexto(Tx_tarjeta.Text)) Then
         If Trim(Tx_tarjeta.Text) <> "" And EsTexto(Tx_tarjeta.Text) Then
            Ayuda1.TxUbiAP = Trim(Tx_tarjeta.Text)
         End If
         Load Ayuda1
         Ayuda1.Tag = 5
         Ayuda1.Show 1
      Else
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Tx_tarjeta & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "Tarjeta, No Ubicada", vbExclamation
            Tx_tarjeta = "":               Tx_trabajador = ""
            Tx_Plh = ""
            Tx_Cargo = "": Tx_DscCargo = ""
         Else
            Tx_trabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
            Txccosto = IIf(IsNull(rs!CCosto), "", rs!CCosto)
            Txcodupr = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)
            Tx_Plh = IIf(IsNull(rs!PLH), "", rs!PLH)
            Tx_Cargo = IIf(IsNull(rs!Cargo), "", Trim(rs!Cargo))
         End If
      End If
      
      If Tx_Plh = "" Then
         MsgBox "Tarjeta, No indica REGIMEN LABORAL...", vbExclamation
         Tx_tarjeta = "":               Tx_trabajador = ""
         Tx_Plh = ""
         Tx_Cargo = "": Tx_DscCargo = ""
         Exit Sub
      End If
      
      ceditor = "X"
      If Wccosto = "" Then
         ceditor = "S"
      Else
         If Txccosto <> "" Then
            '-Verifica Acceso a UPR
            If InStr(wCadenaUpr, Txcodupr) = 0 Then
               MsgBox "Tarjeta, No Autorizada...", vbExclamation
               Tx_tarjeta = "":               Tx_trabajador = ""
               Tx_Plh = ""
               Tx_Cargo = "": Tx_DscCargo = ""
               Exit Sub
            End If
         End If
         'Derechos en Upr
         Set rs = New Recordset
         rs.Open "SELECT * FROM CpCcostoUPR WHERE ccosto='" & Txccosto & "' and Codupr='" & Txcodupr & "'", Cnx, adOpenStatic, adLockReadOnly
         If rs.EOF Then
         Else
            If rs!Dni1 = wUsuario Then
               ceditor = rs!editor1
            ElseIf rs!Dni2 = wUsuario Then
                ceditor = rs!editor2
            ElseIf rs!dni3 = wUsuario Then
                ceditor = rs!editor3
            End If
         End If
         rs.Close
         If ceditor <> "S" Then
            MsgBox "Operador, No autorizado..", vbExclamation
            Tx_tarjeta = "":            Tx_trabajador = ""
            Tx_Plh = ""
            Tx_Cargo = "": Tx_DscCargo = ""
            Exit Sub
         End If
      End If
      
      If Tx_Cargo <> "" Then
         Set r_t = New Recordset
         r_t.Open "SELECT * FROM CpCargos where codigo='" & Tx_Cargo + "'", Cnx, adOpenStatic, adLockReadOnly
         If Not r_t.EOF Then
            Tx_DscCargo = Trim(r_t!descripcion)
         End If
         r_t.Close
      End If
      
      dfi = Day(primer)
      dff = Day(ultimo)
      For i = dfi To dff
        z = i - 1
        Option1(z).Value = False
        Textd(z).Visible = True: Textp(z).Visible = True
        Textd(z).Text = "": Textd(z).ForeColor = &H80000012
        Textp(z).Text = "":  Textp(z).ForeColor = &H80000012
        Lblhr(z).Caption = ""
      Next
      Call cmdcargTRA_Click
      Call CmdChgReg_Click
      
      fecha1 = Format(primer, "yyyymmdd")
      fecha2 = Format(ultimo, "yyyymmdd")
      Set rs = New Recordset
      sql = "SELECT count(*) AS Tot FROM CpChgTur" & genaño & " WHERE tarjeta='" + Tx_tarjeta.Text + "' and fecha>='" & fecha1 & "' AND fecha<='" & fecha2 & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      nson = rs!tot
      rs.Close
      cmdGraba.Enabled = True
      If nson >= 2 And Trim(wDerEspec) <> "S" And InStr("RES_3RO", Tx_Plh.Text) = 0 Then
         'bloquea registros
         MsgBox "Completo cambios autorizados..", vbExclamation       '1
         cmdGraba.Enabled = False
      End If
      
      Exit Sub
   End If
End Sub

Private Sub cmdcargTRA_Click()
   TxHoras = "": TxD = "": TxN = "": TxM = "": Txt = ""
   If Tx_tarjeta = "" Then Exit Sub
   For i = dfi To dff
       z = i - 1
       Option1(z).Value = False
       Textd(z).Visible = True: Textp(z).Visible = True
       Textd(z).Text = "": Textd(z).ForeColor = &H80000012
       Textp(z).Text = "":  Textp(z).ForeColor = &H80000012
       Lblhr(z).Caption = ""
   Next
   fecha1 = Format(primer, "yyyymmdd")
   fecha2 = Format(ultimo, "yyyymmdd")
      
   '-Carga  DIRECTORIO
   sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Tx_tarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   cDirectHoy = ""
   If Not rs.EOF Then
      For i = Day(primer) To Day(ultimo)
         z = i - 1
         Textd(z) = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
         
         cDirectHoy = cDirectHoy + IIf(IsNull(Textd(z)), "==", Left(Trim(Textd(z)) + "==", 2))
         '-D N P MTMXD*V ==N*...
         ' 1234567890123456
         
         Select Case Trim(Textd(z).Text)
              Case "D"
                 TxD = Val(TxD) + 1
              Case "D1"
                 TxD = Val(TxD) + 1
              Case "D2"
                 TxD = Val(TxD) + 1
              Case "D*"
                 TxD = Val(TxD) + 1
              Case "D3"
                 TxD = Val(TxD) + 1
              Case "MT"
                 TxMt = Val(TxMt) + 1
              Case "MX"
                 TxMt = Val(TxMt) + 1
                                  
              Case "N"
                 TxN = Val(TxN) + 1
              Case "N2"
                 TxN = Val(TxN) + 1
              Case "N*"
                 TxN = Val(TxN) + 1
              Case "N3"
                 TxN = Val(TxN) + 1
              Case "NA"
                 TxNA = Val(TxNA) + 1
                 
              Case "M"
                 TxM = Val(TxM) + 1
              Case "M1"
                 TxM = Val(TxM) + 1
              Case "M2"
                 TxM = Val(TxM) + 1
              Case "M3"
                 TxM = Val(TxM) + 1
                 
              Case "T"
                 Txt = Val(Txt) + 1
              Case "T1"
                 Txt = Val(Txt) + 1
              Case "T2"
                 Txt = Val(Txt) + 1
                 
              Case "H"
                   TxM = Val(TxM) + 1
                   TxN = Val(TxN) + 1
              Case "H1"
                   TxM = Val(TxM) + 1
                   TxN = Val(TxN) + 1
              Case "H*"
                   TxM = Val(TxM) + 1
                   TxN = Val(TxN) + 1
              Case "C"
                   Txt = Val(Txt) + 1
                   TxN = Val(TxN) + 1
              Case "C1"
                   Txt = Val(Txt) + 1
                   TxN = Val(TxN) + 1
              Case "C*"
                   Txt = Val(Txt) + 1
                   TxN = Val(TxN) + 1
                   
              Case "P"
                   TxD = Val(TxD) + 1
                   TxN = Val(TxN) + 1
              Case "P1"
                   TxD = Val(TxD) + 1
                   TxN = Val(TxN) + 1
              Case "P*"
                   TxD = Val(TxD) + 1
                   TxN = Val(TxN) + 1
              Case "PA"
                   TxPA = Val(TxPA) + 1
                   
         End Select
      Next i
      '---  MsgBox "DirHoy(" + cDirectHoy + ")"
      
   End If
   rs.Close
   '-Carga HORARIO al DIRECTORIO
   sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Tx_tarjeta & "' AND fecha >='" & fecha1 & "' AND fecha<='" & fecha2 & "' ORDER BY Fecha,tarjeta"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      rs.MoveFirst
      Do While Not rs.EOF
         i = Day(rs!fecha)
         z = i - 1
         Textp(z) = IIf(IsNull(rs!Horario), "", rs!Horario)
         If Textp(z) <> "" Then
            Set rs_d = New Recordset
            sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & Trim(Textp(z)) & "'"      'CpHorarios
            rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If rs_d.EOF Then
               rs_d.Close
               Set rs_d = New Recordset
               sq_l = "SELECT * FROM CpHorarios WHERE Codigo='" & Trim(Textp(z)) & "'"      'CpHorarios
               rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            End If
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

Private Sub CmdChgReg_Click()
   If Txmes.Text = "" Then Exit Sub
   If Trim(Tx_tarjeta) <> "" And Trim(Tx_trabajador) = "" Then
      Set rs = New Recordset
      sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and  tarjeta = '" & Tx_tarjeta & "'"
      rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "No Ubicado", vbExclamation
         Tx_tarjeta = ""
         Tx_Plh = ""
         Exit Sub
      Else
         Tx_trabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
      End If
      rs.Close
   End If
   If Trim(Txmes) <> "" Then
      fecha1 = Format(primer, "yyyymmdd")
      fecha2 = Format(ultimo, "yyyymmdd")
   End If
   If Trim(Tx_tarjeta) = "" Then
      sql = "SELECT * FROM CpChgTur" & genaño & " WHERE fecha>='" & fecha1 & "' And  fecha<='" & fecha2 & "' ORDER BY tarjeta,fecha "
      '-SELECT   tb1.tarjeta AS codigo, tb1.FECHA, tb2.CodUPR AS Upr, tb1.AHor0, tb1.AHor
      ' FROM     CpChgTur2013 tb1 INNER JOIN
      '          CpGeneral tb2 ON tb2.Tarjeta = tb1.tarjeta AND tb1.FECHA >= '01/08/2013' AND tb1.FECHA <= '30/08/2013' AND CHARINDEX(tb2.CodUPR,'10_11_13_14') > 0
      'ORDER BY tb1.tarjeta, tb1.FECHA
   Else
      sql = "SELECT * FROM CpChgTur" & genaño & " WHERE fecha>='" & fecha1 & "' AND fecha<='" & fecha2 & "' AND tarjeta='" & Tx_tarjeta.Text & "' ORDER BY tarjeta,fecha "
   End If
   DataGrid1.Enabled = True
   Call cargaDD
End Sub
Private Sub cargaDD()
   nson = 0
   Do While Not DD.EOF
      DD.Delete
      DD.MoveFirst
   Loop
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      rs.MoveFirst
      Do While Not rs.EOF
         TarjOk = "S"
         If wCadenaUpr <> "" Then
            Set r_t = New Recordset
            sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta ='" & rs!tarjeta & "'"
            r_t.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If Not r_t.EOF Then
               If InStr(wCadenaUpr, r_t!CodUpr) > 0 Then
               Else
                  TarjOk = "N"
               End If
            End If
            r_t.Close
         End If
         If TarjOk = "S" Then
            '-Es dominio de UPR
            nson = nson + 1
            DD.AddNew
            DD!tarjeta = IIf(IsNull(rs!tarjeta), "", rs!tarjeta)
            If IsDate(rs!fecha) Then DD!fecha = rs!fecha
            DD!dehor0 = IIf(IsNull(rs!dehor0), "", rs!dehor0)
            DD!Ahor0 = IIf(IsNull(rs!Ahor0), "", rs!Ahor0)
            If IsDate(rs!nfecha) Then DD!nfecha = rs!nfecha
            DD!dehor = IIf(IsNull(rs!dehor), "", rs!dehor)
            DD!aHor = IIf(IsNull(rs!aHor), "", rs!aHor)
            DD!tipo = IIf(IsNull(rs!tipo), "", rs!tipo)
            DD!documento = IIf(IsNull(rs!documento), "", rs!documento)
            DD.Update
         End If
         rs.MoveNext
      Loop
   End If
   rs.Close
   Call Limpiacampos
End Sub
Private Sub CmdClear_Click()
   dfi = Day(primer)
   dff = Day(ultimo)
   For i = dfi To dff
     z = i - 1
     Option1(z).Value = False
     Textd(z).Text = ""
     Textp(z).Text = ""
     Lblhr(z).Caption = ""
   Next i
   Do While Not DD.EOF
      DD.Delete
      DD.MoveFirst
   Loop
   Call Limpiacampos
   Tx_tarjeta = "": Tx_trabajador = "": Tx_Plh = ""
   Tx_Cargo = "": Tx_DscCargo = ""
   Tx_tarjeta.SetFocus
End Sub

Private Sub Limpiacampos()
   Txtarjeta = "": Txtrabajador = "": Txtarjeta.Locked = True
   Cbodehor0 = "":    CboAhor0.Text = ""
   Txdefecha = "": TxdeHor0 = "": TxAHor0 = "": TxDetur0 = "": TxAtur0 = "": TxdeHrs0 = "": TxAhrs0 = ""
   CboDeHor.Text = "":    CboAHor.Text = ""
   Txafecha = "": TxdeHor = "": TxAhor.Text = "": TxDeTur = "": TxAtur = "": TxDeHrs = "": TxAhrs = ""
   Txtipo = "": TxDocumento = ""
   Check1.Value = 0: Check2.Value = 0
   Check1.Enabled = False: Check2.Enabled = False
   Txccosto = "": Txcodupr = "": Txcargo = "": TxCarEst = "": Txtipopla = ""
   'DataGrid1.SetFocus
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   With DD      ' Adodc1.Recordset
      If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
         .Sort = .Fields(ColIndex).[Name] & " Desc"
      Else
         .Sort = .Fields(ColIndex).[Name] & " Asc"
      End If
   End With
   Call Limpiacampos
End Sub
Private Sub DataGrid1_Click()
    If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
    On Error GoTo final
    Call Limpiacampos
    Txtarjeta = IIf(IsEmpty(DataGrid1.Columns(0)), "", DataGrid1.Columns(0))
    If IsDate(DataGrid1.Columns(1)) Then Txdefecha.Text = DataGrid1.Columns(1)
    TxdeHor0 = IIf(IsEmpty(DataGrid1.Columns(2)), "", DataGrid1.Columns(2))
    TxAHor0 = IIf(IsEmpty(DataGrid1.Columns(3)), "", DataGrid1.Columns(3))
    If IsDate(DataGrid1.Columns(4)) Then Txafecha.Text = DataGrid1.Columns(4)
    TxdeHor = IIf(IsEmpty(DataGrid1.Columns(5)), "", DataGrid1.Columns(5))
    TxAhor = IIf(IsEmpty(DataGrid1.Columns(6)), "", DataGrid1.Columns(6))
    Txtipo = IIf(IsEmpty(DataGrid1.Columns(7)), "", DataGrid1.Columns(7))
    TxDocumento = IIf(IsEmpty(DataGrid1.Columns(8)), "", DataGrid1.Columns(8))
    If Trim(Txtarjeta) <> "" Then
       Set rs = New Recordset
       rs.Open "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta='" & Txtarjeta.Text & "'", Cnx, adOpenStatic, adLockReadOnly
       If Not rs.EOF Then
          Txtrabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
          Txccosto = IIf(IsNull(rs!CCosto), "", rs!CCosto)
          Txcodupr = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)
          Txcargo = IIf(IsNull(rs!Cargo), "", rs!Cargo)
          TxCarEst = IIf(IsNull(rs!carEst), "", rs!carEst)
          Txtipopla = IIf(IsNull(rs!tipoPla), "", rs!tipoPla)
          Txplh = IIf(IsNull(rs!PLH), "", rs!PLH)
       End If
       rs.Close
    End If
    If Trim(TxdeHor0) <> "" Then
       For i = 0 To Cbodehor0.ListCount - 1
          npos = InStr(Cbodehor0.List(i), ">")
          nlen = Len(Cbodehor0.List(i))
          dato = Trim(Right(Cbodehor0.List(i), nlen - npos))
          If dato = Trim(TxdeHor0) Then
             Cbodehor0.ListIndex = i
             Exit For
          End If
       Next i
    End If
    If Trim(TxAHor0) <> "" Then
       For i = 0 To CboAhor0.ListCount - 1
          npos = InStr(CboAhor0.List(i), ">")
          nlen = Len(CboAhor0.List(i))
          dato = Trim(Right(CboAhor0.List(i), nlen - npos))
          If dato = Trim(TxAHor0) Then
             CboAhor0.ListIndex = i
             Exit For
          End If
       Next i
    End If
    If Trim(TxdeHor) <> "" Then
       For i = 0 To CboDeHor.ListCount - 1
          npos = InStr(CboDeHor.List(i), ">")
          nlen = Len(CboDeHor.List(i))
          dato = Trim(Right(CboDeHor.List(i), nlen - npos))
          If dato = Trim(TxdeHor) Then
             CboDeHor.ListIndex = i
             Exit For
          End If
       Next i
    End If
    If Trim(TxAhor) <> "" Then
       For i = 0 To CboAHor.ListCount - 1
          npos = InStr(CboAHor.List(i), ">")
          nlen = Len(CboAHor.List(i))
          dato = Trim(Right(CboAHor.List(i), nlen - npos))
          If dato = Trim(TxAhor) Then
             CboAHor.ListIndex = i
             Exit For
          End If
       Next i
    End If
    Sw = False
    Frame1.Enabled = False
    
    Exit Sub
final:
   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
End Sub
Private Sub dataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   '''On Error GoTo final
   If KeyCode = 13 Then
      Call DataGrid1_Click
   End If
'''final:
   '''MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
End Sub

Private Sub cmdGraba_Click()
   If wrights = "CON" Or ceditor <> "S" Then
      MsgBox "Operador, No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   ''If Format(fecCierre, "yyyymmdd") <= Format(Date, "yyyymmdd") Then    'And wDerEspec <> "S"
   If fecCierre <= Date Then
      If (Txplh = "3RO" And Txtipopla = "ASI") Then
      Else
         MsgBox "Programa CERRADO...  " + Format(fecCierre, "yyyymmdd") + "<=" + Format(Date, "yyyymmdd"), vbExclamation
         Exit Sub
      End If
   End If
   
          
   If Trim(TxAHor0) <> "" Then
      '-Cambio en Dia Base
      Txtipo = "0"
   End If
   If Trim(TxAhor) <> "" Then
      '-Cambio en Dia Base y Dia Nuevo
      Txtipo = "1"
   End If
   If Not IsDate(Txdefecha) Or DateValue(Txdefecha) < primer Or DateValue(Txdefecha) > ultimo Then
      MsgBox "<E> Fecha Base,  verifique contenido..  ", vbExclamation
      Exit Sub
   End If
   If Txtipo = "1" Then
      If IsDate(Txafecha) Then
         If DateValue(Txafecha) > ultimo Then
            MsgBox "<E> Fecha Nueva, verifique...  ", vbExclamation, "[IMPORTANTE]"
            Exit Sub
         End If
      Else
         MsgBox "<E> Fecha Nueva, verifique...  ", vbExclamation, "[IMPORTANTE]"
         Exit Sub
      End If
   End If
   If IsNull(Txtarjeta) Then
      MsgBox "Ingrese Nro tarjeta...  ", vbCritical, "[IMPORTANTE"
      Exit Sub
   Else
      Set rs = New Recordset
      sql = "SELECT * FROM CpGeneral  WHERE  (baja is NULL) and tarjeta='" & Txtarjeta & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "Error, Tarjeta No existe..! ", vbCritical, "[IMPORTANTE]"
         rs.Close
         Exit Sub
      Else
         rs.Close
      End If
   End If
   If IsDate(Txafecha) Then
      If DateValue(Txdefecha) > DateValue(Txafecha) Then
         MsgBox "Verifique fechas...  ", vbExclamation
         Exit Sub
      End If
   End If
   If IsEmpty(TxAHor0) Then
      MsgBox "<E> CAMBIO horario, Dia BASE ", vbExclamation
      Exit Sub
   End If
   If IsDate(Txafecha) Then
      If IsEmpty(TxAhor) Then
         MsgBox "<E> CAMBIO, Dia NUEVO ", vbExclamation
         Exit Sub
      End If
   End If
   If Trim(TxDocumento) = "" Then
      MsgBox "Indique Doc. que Autoriza ..", vbExclamation
      Exit Sub
   End If
   
   '-carga detalles del cambio
   Dim rt As Recordset
   Set rt = New Recordset
   rt.Open "SELECT * FROM cphorarionew  order by codigo", Cnx, adOpenStatic, adLockReadOnly
   Do While Not rt.EOF
      If Trim(TxdeHor0) <> "" Then
         If rt!Codigo = TxdeHor0 Then
            '''-TxDetur0 = IIf(IsNull(rt!turno), "", rt!turno)
            TxdeHrs0 = IIf(IsNull(rt!Horas), 0, Val(Left(rt!Horas, 2)) + (Val(Right(rt!Horas, 2)) / 60))
         End If
      End If
      If Trim(TxAHor0) <> "" Then
         If rt!Codigo = TxAHor0.Text Then
            If Check1.Value = 1 And InStr("N_D_H_C_P", Left(rt!turno, 1)) > 0 Then
               TxAtur0 = IIf(IsNull(rt!turno), "", Trim(rt!turno)) + "*"
            Else
               TxAtur0 = IIf(IsNull(rt!turno), "", rt!turno)
            End If
            TxAhrs0 = IIf(IsNull(rt!Horas), 0, Val(Left(rt!Horas, 2)) + (Val(Right(rt!Horas, 2)) / 60))
         End If
      End If
      If Trim(TxdeHor) <> "" Then
         If rt!Codigo = TxdeHor Then
            '''-TxDeTur = IIf(IsNull(rt!turno), "", rt!turno)
            TxDeHrs = IIf(IsNull(rt!Horas), 0, Val(Left(rt!Horas, 2)) + (Val(Right(rt!Horas, 2)) / 60))
         End If
      End If
      If Trim(TxAhor) <> "" Then
         If rt!Codigo = TxAhor Then
            If Check2.Value = 1 And InStr("N_D_H_C_P", Left(rt!turno, 1)) > 0 Then
               TxAtur = IIf(IsNull(rt!turno), "", Trim(rt!turno)) + "*"
            Else
               TxAtur = IIf(IsNull(rt!turno), "", rt!turno)
            End If
            TxAhrs = IIf(IsNull(rt!Horas), 0, Val(Left(rt!Horas, 2)) + (Val(Right(rt!Horas, 2)) / 60))
         End If
      End If
      rt.MoveNext
   Loop
   rt.Close
   If Trim(TxdeHor0) = Trim(TxAHor0) And Trim(TxDetur0) = Trim(TxAtur0) Then
      MsgBox "No modifico TURNO del dìa Base ..", vbCritical
      Exit Sub
   End If
   If Trim(Trim(TxdeHor)) <> "" And Trim(TxdeHor) = Trim(TxAhor) And Trim(TxDeTur) = Trim(TxAtur) Then
      MsgBox "No modifico TURNO dìa Nuevo ..", vbCritical
      Exit Sub
   End If
   If Txtipo = "0" Then
      If Trim(TxdeHor0) = Trim(TxAHor0) And Trim(TxDetur0) = Trim(TxAtur0) Then
         MsgBox "Transacciòn No es VALIDA...", vbCritical
         Exit Sub
      End If
   Else
      If Trim(TxdeHor0) = Trim(TxAhor) And Trim(TxDetur0) = Trim(TxAtur) And DateValue(Txdefecha) = DateValue(Txafecha) Then
         MsgBox "Transacciòn No es VALIDA...", vbCritical
         Exit Sub
      End If
  End If
   FlgExcep = False
   FlgXGuar = False      'No Exceda Guardias Permitidas   12 - 150hras
   msgExcep = ""
   Call ChkExcepcion
   
   If FlgXGuar Then
      MsgBox "[Cambio No AUTORIZADO], Excede Horas permitidas...", vbExclamation
      Exit Sub
   End If
   If FlgExcep And wDerEspec <> "S" Then
      MsgBox msgExcep + "' [Cambio No AUTORIZADO] '", vbExclamation
      Exit Sub
   End If
   
   fecha1 = Format(Txdefecha, "yyyymmdd")
   fecha2 = Format(Txafecha, "yyyymmdd")
   On Error GoTo final
   Cnx.BeginTrans
   
   Set rt = New Recordset
   sq2 = "SELECT * FROM CpChgTur" & genaño & " WHERE tarjeta ='" & Trim(Txtarjeta) & "' and fecha='" & fecha1 & "' and DeHor0='" & TxdeHor0.Text & "'"
   rt.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If rt.EOF Then
      rt.AddNew
      rt!tarjeta = Txtarjeta
      If IsDate(Txdefecha) Then rt!fecha = Txdefecha
      rt!dehor0 = TxdeHor0
   End If
    rt!usuario = wNick    'wUsuario      Firma Usuario
    rt!fecRegistro = Date
    rt!HorRegistro = Format(Time, "hh:mm:ss")
    rt!estacion = w_clienName
    
   rt!Ahor0 = TxAHor0
   Txtipo = "0"
   If IsDate(Txafecha) Then
      rt!nfecha = Txafecha
      rt!dehor = IIf(IsNull(TxdeHor), "", TxdeHor)
      rt!aHor = IIf(IsNull(TxAhor), "", TxAhor)
      Txtipo = "1"
   End If
   rt!tipo = Txtipo
   rt!documento = IIf(IsNull(TxDocumento), "", TxDocumento)
   rt.Update
   rt.Close
   If Sw Then
      DD.AddNew
      DD!tarjeta = Txtarjeta
      If IsDate(Txdefecha) Then DD!fecha = Txdefecha
      DD!dehor0 = TxdeHor0
   End If
   DD!Ahor0 = TxAHor0
   If IsDate(Txafecha) Then DD!nfecha = Txafecha
   DD!dehor = IIf(IsNull(TxdeHor), "", TxdeHor)
   DD!aHor = IIf(IsNull(TxAhor), "", TxAhor)
   DD!tipo = Txtipo.Text
   DD!documento = IIf(IsNull(TxDocumento), "", TxDocumento)
   DD.Update
   '-DE Fecha (origen)
   sql = "SELECT * FROM CpChgTurM" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND Fecha ='" & fecha1 & "' AND horario='" & TxdeHor0.Text & "'"
   Set rs_d = New Recordset
   rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If rs_d.EOF Then
      rs_d.AddNew
      rs_d!tarjeta = Txtarjeta
      rs_d!fecha = Txdefecha
      rs_d!Horario = TxdeHor0
   End If
   rs_d!CamHor = TxAHor0
   rs_d!documento = IIf(IsNull(TxDocumento), "", TxDocumento)
   
   rs_d!usuario = wNick
   rs_d!fecRegistro = Date
   rs_d!HorRegistro = Format(Time, "hh:mm:ss")
   rs_d!estacion = w_clienName
   rs_d.Update
   rs_d.Close
   If Txtipo = "1" Then
      '-A Fecha (destino)
      sql = "SELECT * FROM CpChgTurM" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND Fecha ='" & fecha2 & "' AND horario='" & TxdeHor.Text & "'"
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      If rs_d.EOF Then
         rs_d.AddNew
         rs_d!tarjeta = Txtarjeta
         rs_d!fecha = Txafecha
         rs_d!Horario = TxdeHor
      End If
      rs_d!CamHor = TxAhor
      rs_d!documento = IIf(IsNull(TxDocumento), "", TxDocumento)
      
      rs_d!usuario = wNick
      rs_d!fecRegistro = Date
      rs_d!HorRegistro = Format(Time, "hh:mm:ss")
      rs_d!estacion = w_clienName
      rs_d.Update
      rs_d.Close
   End If
   
   '-Aplica sumario Directorio
   nSonD = IIf(IsNull(TxD), 0, Val(TxD))
   nSonN = IIf(IsNull(TxN), 0, Val(TxN))
   nSonM = IIf(IsNull(TxM), 0, Val(TxM))
   nSonT = IIf(IsNull(Txt), 0, Val(Txt))
    nSonMT = IIf(IsNull(TxMt), 0, Val(TxMt))
    nSonNA = IIf(IsNull(TxNA), 0, Val(TxNA))
    nSonPA = IIf(IsNull(TxPA), 0, Val(TxPA))
   Select Case TxDetur0.Text
     Case "D"
      nSonD = nSonD - 1
     Case "D1"
      nSonD = nSonD - 1
     Case "D2"
      nSonD = nSonD - 1
     Case "D3"
      nSonD = nSonD - 1
     Case "D*"
      nSonD = nSonD - 1
     Case "MT"
      nSonMT = nSonMT - 1
     Case "MX"
      nSonMT = nSonMT - 1
      
     Case "N"
      nSonN = nSonN - 1
     Case "N2"
      nSonN = nSonN - 1
     Case "N3"
      nSonN = nSonN - 1
     Case "N*"
      nSonN = nSonN - 1
     Case "NA"
      nSonNA = nSonNA - 1
      
     Case "M"
      nSonM = nSonM - 1
     Case "M1"
      nSonM = nSonM - 1
     Case "M2"
      nSonM = nSonM - 1
     Case "M3"
      nSonM = nSonM - 1
      
     Case "T"
      nSonT = nSonT - 1
     Case "T1"
      nSonT = nSonT - 1
     Case "T2"
      nSonT = nSonT - 1
      
     Case "H"
        nSonM = nSonM - 1
        nSonN = nSonN - 1
     Case "H1"
        nSonM = nSonM - 1
        nSonN = nSonN - 1
     Case "H*"
         nSonM = nSonM - 1
         nSonN = nSonN - 1
     Case "C"
        nSonT = nSonT - 1
        nSonN = nSonN - 1
     Case "C1"
        nSonT = nSonT - 1
        nSonN = nSonN - 1
     Case "C*"
        nSonT = nSonT - 1
        nSonN = nSonN - 1
        
     Case "P"
        nSonD = nSonD - 1
        nSonN = nSonN - 1
     Case "P1"
         nSonD = nSonD - 1
         nSonN = nSonN - 1
     Case "P*"
         nSonD = nSonD - 1
         nSonN = nSonN - 1
     Case "PA"
         nSonPA = nSonPA - 1
   End Select
   Select Case TxAtur0
     Case "D"
      nSonD = nSonD + 1
     Case "D1"
      nSonD = nSonD + 1
     Case "D2"
      nSonD = nSonD + 1
     Case "D3"
      nSonD = nSonD + 1
     Case "D*"
      nSonD = nSonD + 1
     Case "MT"
      nSonMT = nSonMT + 1
     Case "MX"
      nSonMT = nSonMT + 1
     Case "N"
      nSonN = nSonN + 1
     Case "N2"
      nSonD = nSonD + 1
     Case "N3"
      nSonD = nSonD + 1
     Case "N*"
      nSonN = nSonN + 1
     Case "NA"
      nSonNA = nSonNA + 1
      
     Case "M"
      nSonM = nSonM + 1
     Case "M1"
      nSonM = nSonM + 1
     Case "M2"
      nSonM = nSonM + 1
     Case "M3"
      nSonM = nSonM + 1
      
     Case "T"
      nSonT = nSonT + 1
     Case "T1"
      nSonT = nSonT + 1
     Case "T2"
      nSonT = nSonT + 1
      
     Case "H"
        nSonM = nSonM + 1
        nSonN = nSonN + 1
     Case "H1"
        nSonM = nSonM + 1
        nSonN = nSonN + 1
     Case "H*"
        nSonM = nSonM + 1
        nSonN = nSonN + 1
     Case "C"
        nSonT = nSonT + 1
        nSonN = nSonN + 1
     Case "C1"
        nSonT = nSonT + 1
        nSonN = nSonN + 1
     Case "C*"
        nSonT = nSonT + 1
        nSonN = nSonN + 1
        
     Case "P"
        nSonD = nSonD + 1
        nSonN = nSonN + 1
     Case "P1"
        nSonD = nSonD + 1
        nSonN = nSonN + 1
     Case "P*"
        nSonD = nSonD + 1
        nSonN = nSonN + 1
     Case "PA"
        nSonPA = nSonPA + 1
   End Select
   If Txtipo = "1" Then
      Select Case TxDeTur
        Case "D"
         nSonD = nSonD - 1
        Case "D1"
         nSonD = nSonD - 1
        Case "D2"
         nSonD = nSonD - 1
        Case "D3"
         nSonD = nSonD - 1
        Case "D*"
         nSonD = nSonD - 1
        Case "MT"
         nSonMT = nSonMT + 1
        Case "MX"
         nSonMT = nSonMT + 1

        Case "N"
         nSonN = nSonN - 1
        Case "N2"
         nSonN = nSonN - 1
        Case "N3"
         nSonN = nSonN - 1
        Case "N*"
         nSonN = nSonN - 1
        Case "NA"
         nSonNA = nSonNA - 1
         
        Case "M"
         nSonM = nSonM - 1
        Case "M1"
         nSonM = nSonM - 1
        Case "M2"
         nSonM = nSonM - 1
        Case "M3"
         nSonM = nSonM - 1
         
        Case "T"
         nSonT = nSonT - 1
        Case "T1"
         nSonT = nSonT - 1
        Case "T2"
         nSonT = nSonT - 1
         
        Case "H"
           nSonM = nSonM - 1
           nSonN = nSonN - 1
        Case "H1"
           nSonM = nSonM - 1
           nSonN = nSonN - 1
        Case "H*"
           nSonM = nSonM - 1
           nSonN = nSonN - 1
        Case "C"
           nSonT = nSonT - 1
           nSonN = nSonN - 1
        Case "C1"
           nSonT = nSonT - 1
           nSonN = nSonN - 1
        Case "C*"
           nSonT = nSonT - 1
           nSonN = nSonN - 1
           
        Case "P"
           nSonD = nSonD - 1
           nSonN = nSonN - 1
        Case "P1"
           nSonD = nSonD - 1
           nSonN = nSonN - 1
        Case "P*"
           nSonD = nSonD - 1
           nSonN = nSonN - 1
        Case "PA"
           nSonPA = nSonPA - 1
           
      End Select
      Select Case TxAtur
        Case "D"
         nSonD = nSonD + 1
        Case "D1"
         nSonD = nSonD + 1
        Case "D2"
         nSonD = nSonD + 1
        Case "D3"
         nSonD = nSonD + 1
        Case "D*"
         nSonD = nSonD + 1

        Case "MT"
         nSonMT = nSonMT + 1
        Case "MX"
         nSonMT = nSonMT + 1
      
        Case "N"
         nSonN = nSonN + 1
        Case "N2"
         nSonN = nSonN + 1
        Case "N3"
         nSonN = nSonN + 1
        Case "N*"
         nSonN = nSonN + 1
        Case "NA"
         nSonNA = nSonNA + 1
         
        Case "M"
         nSonM = nSonM + 1
        Case "M1"
         nSonM = nSonM + 1
        Case "M2"
         nSonM = nSonM + 1
        Case "M3"
         nSonM = nSonM + 1
        Case "T"
         nSonT = nSonT + 1
        Case "T1"
         nSonT = nSonT + 1
        Case "T2"
         nSonT = nSonT + 1
         
        Case "H"
           nSonM = nSonM + 1
           nSonN = nSonN + 1
        Case "H1"
           nSonM = nSonM + 1
           nSonN = nSonN + 1
        Case "H*"
           nSonM = nSonM + 1
           nSonN = nSonN + 1
        Case "C"
           nSonT = nSonT + 1
           nSonN = nSonN + 1
        Case "C1"
           nSonT = nSonT + 1
           nSonN = nSonN + 1
        Case "C*"
           nSonT = nSonT + 1
           nSonN = nSonN + 1
           
        Case "P"
           nSonD = nSonD + 1
           nSonN = nSonN + 1
        Case "P1"
           nSonD = nSonD + 1
           nSonN = nSonN + 1
        Case "P*"
           nSonD = nSonD + 1
           nSonN = nSonN + 1
        Case "PA"
           nSonPA = nSonPA + 1
      End Select
   End If
   ntHoras = nSonPA * 24 + (nSonD + nSonN + nSonMT + nSonNA) * 12 + (nSonM + nSonT) * 6
   '-Actualiza  cambio al directorio
   sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
   Set rs_d = New Recordset
   rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If rs_d.EOF Then
      rs_d.AddNew
      rs_d!tarjeta = Txtarjeta
      rs_d!ano = genaño
      rs_d!mes = Txmes
      rs_d!usuario = wNick      'Nick Usuario
      rs_d!fecha = Date
      rs_d!hora = Format(Time, "hh:mm:ss")
      rs_d!estacion = w_clienName
   Else
      rs_d!usuarioM = wNick
      rs_d!fechaM = Date
   End If
   X = Day(Txdefecha)
   Zde = X - 1
   rs_d.Fields(Zde + 4) = IIf(IsNull(TxAtur0), "", TxAtur0)
   If Txtipo = "1" Then
      X = Day(Txafecha)
      ZA = X - 1
      rs_d.Fields(ZA + 4) = IIf(IsNull(TxAtur), "", TxAtur)
   End If
   
   rs_d!d = nSonD
   rs_d!n = nSonN
   rs_d!m = nSonM
   rs_d!T = nSonT
    rs_d!mT = nSonMT
    rs_d!na = nSonNA
     rs_d!pa = nSonPA
   rs_d!Horas = ntHoras
   rs_d!carEst = IIf(IsNull(TxCarEst), "", TxCarEst)
   rs_d!CCosto = IIf(IsNull(Txccosto), "", Txccosto)
   rs_d!CodUpr = IIf(IsNull(Txcodupr), "", Txcodupr)
   rs_d!Cargo = IIf(IsNull(Txcargo), "", Txcargo)
   rs_d!Flag = "0"       'ASISTENCIAL
   rs_d!usuarioM = wNick      'Nick Usuario
   rs_d!fechaM = Date
   rs_d.Update
   rs_d.Close
   
   sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fecha1 & "'"
   Set rs_d = New Recordset
   rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If rs_d.EOF Then
      rs_d.AddNew
      rs_d!tarjeta = Txtarjeta
      rs_d!fecha = Txdefecha
   End If
   rs_d!usuario = wNick      'Nick Usuario
   rs_d!fecRegistro = Date
   rs_d!HorRegistro = Format(Time, "hh:mm:ss")
   rs_d!estacion = w_clienName
   rs_d!Horario = IIf(IsNull(TxAHor0), "", TxAHor0)
   rs_d!fecCambio = Date
   rs_d.Update
   rs_d.Close
   
   If Txtipo = "1" Then
      sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fecha2 & "'"
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      If rs_d.EOF Then
         rs_d.AddNew
         rs_d!tarjeta = Txtarjeta
         rs_d!fecha = Txafecha
      End If
      rs_d!usuario = wNick      'Nick Usuario
      rs_d!fecRegistro = Date
      rs_d!HorRegistro = Format(Time, "hh:mm:ss")
      rs_d!estacion = w_clienName
      rs_d!Horario = IIf(IsNull(TxAhor), "", TxAhor)
      rs_d!fecCambio = Date
      rs_d.Update
      rs_d.Close
   End If
   Cnx.CommitTrans
   
   '---  MsgBox "ini chkcambioDir " + Trim(Str(Day(primer))) + "-" + Trim(Str(Day(ultimo))) + "(" + cDirectHoy + ")"
      
   cPlantilla0 = ""
   sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
   Set rs_d = New Recordset
   rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If Not rs_d.EOF Then
     For i = Day(primer) To Day(ultimo)
       z = i - 1
       If primer > Date Then
          '-Actualiza Plantilla0
          cPlantilla0 = cPlantilla0 + IIf(IsNull(rs_d.Fields(z + 4)), "==", Left(Trim(rs_d.Fields(z + 4)) + "==", 2))
          '-D N P MTMXD*V ==N*...
          ' 1234567890123456
       Else
          '-verifica en cDirectHoy cambio a Textd(z) --> AudDirect
          cvalor = Trim(Mid(cDirectHoy, (i - 1) * 2 + 1, 2))
          c_a = IIf(IsNull(rs_d.Fields(z + 4)), "==", Left(Trim(rs_d.Fields(z + 4)) + "==", 2))
          If cvalor <> c_a Then
             '--  MsgBox Trim(Str(i)) + "<>de[" + cvalor + "]a[" + c_a + "]"
             SqC = "INSERT INTO CPAudDirect(ano,mes,fecha,ccosto,codupr,tarjeta,deTurno,aTurno,Detalle,registro,estacion,fecregis,horregis) "
             SqC = SqC + " values('" + genaño + "','" + Txmes + "','" + Format(DateValue(primer + z), "yyyymmdd") + "','" + Txccosto + "','" + Txcodupr + "','" + Txtarjeta + "','" & _
             cvalor + "','" + c_a + "','" + Trim(TxDocumento) + "','" + wNick + "','" + w_clienName + "','" + Format(Date, "yyyymmdd") + "','" + Format(Time, "hh:mm:ss") + "')"
             Cnx.Execute SqC
         Else
             '--MsgBox Trim(Str(i)) + "=hoy[" + cvalor + "]a[" + c_a + "]"
             
          End If
       End If
     Next i
     If primer > Date Then
        '--    MsgBox "G pl0(" + cPlantilla0 + ")"
        rs_d!plantilla0 = cPlantilla0
        rs_d.Update
     End If
   End If
   rs_d.Close
   
        
   '''If FecCierre < Date Or (TxPlH = "3RO" And Txtipopla = "ASI") Then
         
      '-TZJ 31.03.2025   autoriza liberar Programa 3RO Asistencial
      Call ProcProg3RO(Txtarjeta, genaño + Txmes)
   '''End If
   
   
   GoTo Final1
final:
  MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
  Cnx.RollbackTrans
Final1:
   Call Limpiacampos
   'DataGrid1.SetFocus
   
   Tx_tarjeta_KeyPress (13)
   
   Sw = False
End Sub

Private Sub CmdElimina_Click()
    MsgBox "<E> No Autorizado..  ", vbExclamation
    Exit Sub
    If IsNull(Txtarjeta) Then Exit Sub
    fecha1 = Format(TxfechaIni, "yyyymmdd")
    If MsgBox("Eliminar  registro ?      ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       Cnx.Execute "DELETE FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "'  AND codigo='" & Txcodigo.Text & "' AND DiaHora='" & TxDiaHora.Text & "'"
       DD.Delete
       If TxTurno.Text <> "" Then
          '-Retira Programa Directorio
          If Month(TxfechaIni.Text) = Val(Txmes.Text) And Month(TxfechaIni.Text) = Month(TxFechaFin.Text) Then
             sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
             Set rs_d = New Recordset
             rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
             If rs_d.EOF Then
             Else
                For i = Day(TxfechaIni) To Day(TxFechaFin)
                   z = i - 1
                   rs_d.Fields(z + 4) = ""
                Next i
             End If
             rs_d.Update
             rs_d.Close
          ElseIf Month(TxfechaIni.Text) <> Month(TxFechaFin.Text) Then
                 nmesI = Month(TxfechaIni.Text)
                 nmesF = Month(TxFechaFin.Text)
                 nmesA = 0
                 ndifdias = DateValue(TxFechaFin) - DateValue(TxfechaIni) + 1
                 For j = 0 To ndifdias - 1
                    unafecha = DateValue(TxfechaIni.Text) + j
                    nmes = Month(unafecha)
                    nano = Year(unafecha)
                    If nmesA <> nmes Then
                       nmesA = nmes
                       Call ObtenerPrimerUltimoDia(unafecha)
                       If nmes = nmesI Then
                          primer = DateValue(TxfechaIni)
                       End If
                       If nmes = nmesF Then
                          ultimo = DateValue(TxFechaFin)
                       End If
                       cmes = Right("00" + Trim(Str(nmes)), 2)
                       cano = Trim(Str(nano))
                       sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & cano & "' AND mes='" & cmes & "'"
                       Set rs_d = New Recordset
                       rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
                       If rs_d.EOF Then
                       End If
                    End If
                    nson = 0
                    For i = Day(primer) To Day(ultimo)
                       z = i - 1
                       nson = nson + 1
                       rs_d.Fields(z + 4) = ""
                    Next i
                    j = j + (nson - 1)
                    rs_d.Update
                    rs_d.Close
                 Next j
          End If
       End If
    End If
    Call Limpiacampos
End Sub
Private Sub CmdLimpia_Click()
   dfi = Day(primer)
   dff = Day(ultimo)
   For i = dfi To dff
     z = i - 1
     Option1(z).Value = False
   Next i
   Call Limpiacampos
End Sub

Private Sub cmdsalir_Click()
  Set DD = Nothing
  Set rs_d = Nothing
  Set rs = Nothing
  
  For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        '- Forms Abiertos
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "CambioTurno" Then
           Unload Forms(cont)
        End If
      End If
  Next
  
  Me.Hide
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set DD = Nothing
  Set rs_d = Nothing
  Set rs = Nothing
  For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        '- Forms Abiertos
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "CambioTurno" Then
           Unload Forms(cont)
        End If
      End If
  Next
  
  Me.Hide
  Unload Me
End Sub
Private Sub Cbo_UPR_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(Cbo_UPR.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbo_UPR_LostFocus()
  If IsEmpty(Cbo_UPR.Text) Then
  Else
     j = Cbo_UPR.ListIndex
     If j <> -1 Then
        nposicion = InStr(Cbo_UPR.List(j), ":")
        nlen = Len(Cbo_UPR.List(j))
        Tx_codupr = Right(Cbo_UPR.List(j), nlen - nposicion)
        If Wccosto <> "" Then
           '-Verifica derechos sobre UPR
           CmdPrint.Enabled = False
           If InStr(wCadenaUpr, Tx_codupr) = 0 Then
              MsgBox "UPR, No Autorizada...", vbExclamation
              Tx_codupr = ""
              Cbo_UPR.Text = ""
              Exit Sub
           End If
        End If
        CmdPrint.Enabled = True
     End If
     Exit Sub
  End If
End Sub
Private Sub mnuFicImpSelec_Click()
    '- Seleccionar la impresora a usar                               (23/Ene/00)
    '- La detección de errores es por si no hay impresora instalada
    On Error Resume Next
    With CommonDialog1
        .DialogTitle = "Seleccionar impresora"
        .Flags = cdlPDPrintSetup
        .ShowPrinter
    End With
    Err = 0
End Sub
Private Sub CmdPrint_Click()
   If Wccosto <> "" Then
      '-Verifica derechos sobre UPR
      If InStr(wCadenaUpr, Tx_codupr) = 0 Then
         MsgBox "UPR, No Autorizada...", vbExclamation
         Tx_codupr = ""
         Cbo_UPR.Text = ""
         Exit Sub
      End If
   End If

   nson = 0
   If IsDate(TxDia) Then
      CargaCambios (1)
   ElseIf Trim(Txmes) <> "" Then
          CargaCambios (0)
   End If
   If nson = 0 Then
      MsgBox "Seleccion, no registra Transacciones..", vbExclamation
      Exit Sub
   End If
   nTipRpt = 3.1
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Exit Sub
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
      Case 1
   End Select
   
   '-Proceso Nuevo
   Call Limpiacampos
   
   fecha1 = Format(primer, "yyyymmdd")
   fecha2 = Format(ultimo, "yyyymmdd")
   Set rs = New Recordset
   sql = "SELECT count(*) AS Tot FROM CpChgTur" & genaño & " WHERE tarjeta='" + Tx_tarjeta + "' and fecha>='" & fecha1 & "' AND fecha<='" & fecha2 & "'"
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   nson = rs!tot
   rs.Close
   
   If nson >= 2 And wDerEspec <> "S" And InStr("RES_3RO", Tx_Plh) = 0 Then
      '-bloquea registros
      MsgBox "Completo cambios autorizadOs..", vbExclamation     '2
      Option1(Index).Value = False
      CmdClear.SetFocus
      Exit Sub
   End If
   
   Frame1.Enabled = True
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If Tx_tarjeta <> "" Then
      Txtarjeta = Tx_tarjeta
      Txplh = Tx_Plh
      Txtarjeta_KeyPress (13)
   Else
      Exit Sub
   End If
   Sw = True
   
   TxAHor0 = ""
   cHorAnt = ""
   CboAhor0.Enabled = True
   Txafecha = "": TxdeHor = "": TxAhor = ""
   CboDeHor.Text = "": CboAHor.Text = "":
   dfi = Day(primer)
   dff = Day(ultimo)
   For i = dfi To dff
     z = i - 1
     If Option1(z).Value = True Then
        Txdefecha = primer + z
        TxdeHor0 = Textp(z).Text
        cHorAnt = Textp(z).Text
        TxAHor0 = ""
        'Textd(z).ForeColor = &HFF0000
        'Textd(z).Locked = False
        If Trim(TxdeHor0) <> "" Then
           For j = 0 To Cbodehor0.ListCount - 1
               npos = InStr(Cbodehor0.List(j), ">")
               nlen = Len(Cbodehor0.List(j))
               dato = Trim(Right(Cbodehor0.List(j), nlen - npos))
               If dato = Trim(TxdeHor0) Then
                  Cbodehor0.ListIndex = j
                  Exit For
               End If
           Next j
        End If
        Option1(z).Value = False
        CboAhor0.SetFocus
        
        Set rs = New Recordset
        sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
        Set rs = New Recordset
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs.EOF Then
           i = Day(Txdefecha)
           z = i - 1
           TxDetur0 = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
        End If
        rs.Close
        
        Exit Sub
     End If
   Next
End Sub





Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Txtarjeta) = "" Or (Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta)) Then
         If Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta) Then
            Ayuda1.TxUbiAP = Trim(Txtarjeta)
         End If
         Load Ayuda1
         Ayuda1.Tag = 5.1
         Ayuda1.Show 1
      Else
         Set rt = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
         rt.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rt.EOF Then
            MsgBox "No Ubicado", vbExclamation
            Txtarjeta.Text = ""
         Else
            Txtrabajador = Left(Trim(rt!apelpat) + " " + Trim(rt!apelmat) + " " + Trim(rt!nombres), 35)
            Txccosto = IIf(IsNull(rt!CCosto), "", rt!CCosto)
            Txcodupr = IIf(IsNull(rt!CodUpr), "", rt!CodUpr)
            Txcargo = IIf(IsNull(rt!Cargo), "", rt!Cargo)
            TxCarEst = rt!carEst
            Txtipopla = IIf(IsNull(rt!tipoPla), "", rt!tipoPla)
            Txplh = IIf(IsNull(rt!PLH), "", rt!PLH)
            Check1.Value = 0:  Check2.Value = 0
            Check1.Enabled = False:    Check2.Enabled = False
            If Txcodupr.Text = "08" Then 'servicio Emergencia
               Check1.Enabled = True
               Check2.Enabled = True
            End If
         End If
         rt.Close
      End If
      
      ceditor = "X"
      If Wccosto = "" Then
         ceditor = "S"
      Else
         If Txccosto <> "" Then
            '-Verifica  titular Prog ROL
            Set rs = New Recordset
            rs.Open "SELECT * FROM CpCcostoUPR WHERE ccosto='" & Txccosto & "' and Codupr='" & Txcodupr & "'", Cnx, adOpenStatic, adLockReadOnly
            If rs.EOF Then
            Else
               If rs!Dni1 = wUsuario Then
                  ceditor = rs!editor1
               ElseIf rs!Dni2 = wUsuario Then
                  ceditor = rs!editor2
               ElseIf rs!dni3 = wUsuario Then
                  ceditor = rs!editor3
               End If
            End If
            rs.Close
            If ceditor <> "S" Then
               MsgBox "operador No autorizado..", vbExclamation
               Txtarjeta = ""
            End If
         End If
      End If
      
      fecha1 = Format(primer, "yyyymmdd")
      fecha2 = Format(ultimo, "yyyymmdd")
      Set rs = New Recordset
      sql = "SELECT count(*) AS Tot FROM CpChgTur" & genaño & " WHERE tarjeta='" + Tx_tarjeta + "' and fecha>='" & fecha1 & "' AND fecha<='" & fecha2 & "' AND tarjeta='" & Txtarjeta & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      nson = rs!tot
      rs.Close
      If nson >= 2 And wDerEspec <> "S" And InStr("RES_3RO", Txplh) = 0 Then
         '-bloquea registros
         MsgBox "Completo cambios Autorizados..", vbExclamation       '3
         Limpiacampos
         cmdGraba.Enabled = False
      Else
         cmdGraba.Enabled = True
      End If
      
      Exit Sub
   End If
End Sub
Private Sub CbodeHor0_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Cbodehor0.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboDeHor0_LostFocus()
   If IsEmpty(Cbodehor0.Text) Then
   Else
      j = Cbodehor0.ListIndex
      If j <> -1 Then
         nposicion = InStr(Cbodehor0.List(j), ">")
         nlen = Len(Cbodehor0.List(j))
         TxdeHor0 = Trim(Right(Cbodehor0.List(j), nlen - nposicion))
      End If
      Exit Sub
   End If
End Sub
Private Sub Cbodehor0_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(Cbodehor0.Text) Then
      Else
         j = Cbodehor0.ListIndex
         If j <> -1 Then
            nposicion = InStr(Cbodehor0.List(j), ">")
            nlen = Len(Cbodehor0.List(j))
            TxdeHor0 = Right(Cbodehor0.List(j), nlen - nposicion)
         End If
      End If
      Exit Sub
   End If
End Sub
Private Sub CboAhor0_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboAhor0.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboAHor0_LostFocus()
   If IsEmpty(CboAhor0.Text) Then
   Else
      j = CboAhor0.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboAhor0.List(j), ">")
         nlen = Len(CboAhor0.List(j))
         TxAHor0 = Right(CboAhor0.List(j), nlen - nposicion)
         If TxdeHor0 = TxAHor0 Then
            If MsgBox("<CHK> no cambio HORARIO...  continua ?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               CboAhor0 = "":              TxAHor0.Text = ""
               Exit Sub
            End If
         End If
         If TxAHor0 <> "O" And Trim(TxAHor0) <> "" Then
            '-Verifica horario Permitido en Unid. Organica
            'cpccostohor-> ccosto+horario
            Set rs = New Recordset
            sql = "SELECT * FROM CpCcostoHor WHERE ccosto='" + Txccosto + "' and horario='" & TxAHor0 & "'"
            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If rs.EOF Then
               MsgBox "<E> Horario no corresponde..", vbExclamation
               CboAhor0.Text = ""
               TxAHor0 = ""
               rs.Close
               Exit Sub
            End If
            rs.Close
            
            If InStr("15_16_17_18_19_20_50_61_62", Trim(TxAHor0)) > 0 Then     'hrs 18 y 24
               'If InStr("RES_3RO", Trim(Txplh.Text)) = 0 And wDerEspec <> "S" Then
               '   MsgBox "<E> Horario no Autorizado.." + Txplh.Text, vbExclamation
               '   CboAhor0.Text = ""
               '   TxAHor0.Text = ""
               '   Exit Sub
               'End If
            End If
         End If
      End If
   End If
End Sub
Private Sub CboAhor0_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboAhor0.Text) Then
      Else
         j = CboAhor0.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboAhor0.List(j), ">")
            nlen = Len(CboAhor0.List(j))
            TxAHor0 = Right(CboAhor0.List(j), nlen - nposicion)
            If TxdeHor0 = TxAHor0 Then
               If MsgBox("<CHK> no cambio HORARIO...  continua ?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  CboAhor0.Text = "":                  TxAHor0.Text = ""
                  Exit Sub
               End If
            End If
            If TxAHor0 <> "O" And Trim(TxAHor0) <> "" Then
               '-Verifica horario Permitido en Unid. Organica
               'cpccostohor-> ccosto+horario
               Set rs = New Recordset
               sql = "SELECT * FROM CpCcostoHor WHERE ccosto='" + Txccosto + "' and horario='" & TxAHor0 & "'"
               rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If rs.EOF Then
                  MsgBox "<E> Horario no corresponde..", vbExclamation
                  CboAhor0 = ""
                  TxAHor0 = ""
                  rs.Close
                  Exit Sub
               End If
               rs.Close
               
               If InStr("15_16_17_18_19_20_50_61_62", Trim(TxAHor0)) > 0 Then      '18 y 24 hrs
                  'If InStr("RES_3RO", Trim(Txplh.Text)) = 0 And wDerEspec <> "S" Then
                  '   MsgBox "<E> Horario no Autorizado..." + Txplh.Text, vbExclamation
                  '   CboAhor0.Text = ""
                  '   TxAHor0.Text = ""
                  '   Exit Sub
                  'End If
               End If
            End If
         End If
      End If
   End If
End Sub
Private Sub TxAfecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsDate(Txafecha) And IsDate(Txdefecha) Then
         If CDate(Txafecha) >= CDate(Txdefecha) Then
            If CDate(Txafecha) >= CDate(primer) And CDate(Txafecha) <= CDate(ultimo) Then
               X = Day(Txafecha.Text)
               z = X - 1
               TxdeHor = Textp(z).Text
               If Trim(TxdeHor) <> "" Then
                  For i = 0 To CboDeHor.ListCount - 1
                     npos = InStr(CboDeHor.List(i), ">")
                     nlen = Len(CboDeHor.List(i))
                     dato = Trim(Right(CboDeHor.List(i), nlen - npos))
                     If dato = Trim(TxdeHor) Then
                        CboDeHor.ListIndex = i
                        Exit For
                     End If
                  Next i
               End If
               Cbodehor_KeyPress (13)
               CboAHor.Enabled = True
               CboAHor.Text = ""
               CboAHor.SetFocus
               
               Set rs = New Recordset
               sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
               Set rs = New Recordset
               rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not rs.EOF Then
                  i = Day(Txafecha)
                  z = i - 1
                  TxDeTur = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
               End If
               rs.Close
  
            Else
               MsgBox "Fuera de Rango...", vbExclamation
            End If
         Else
            MsgBox "fecha BASE, No puede ser mayor a fecha DESTINO...", vbExclamation
         End If
      Else
         MsgBox "Verifique fechas...", vbExclamation
      End If
      Exit Sub
   End If
End Sub

Private Sub CbodeHor_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboDeHor.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboDeHor_LostFocus()
   If IsEmpty(CboDeHor.Text) Then
   Else
      j = CboDeHor.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboDeHor.List(j), ">")
         nlen = Len(CboDeHor.List(j))
         TxdeHor = Trim(Right(CboDeHor.List(j), nlen - nposicion))
       End If
      Exit Sub
   End If
End Sub
Private Sub Cbodehor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboDeHor.Text) Then
      Else
         j = CboDeHor.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboDeHor.List(j), ">")
            nlen = Len(CboDeHor.List(j))
            TxdeHor = Right(CboDeHor.List(j), nlen - nposicion)
         End If
      End If
      Exit Sub
   End If
End Sub
Private Sub CboAhor_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboAHor.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboAHor_LostFocus()
   If IsEmpty(CboAHor.Text) Then
   Else
      j = CboAHor.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboAHor.List(j), ">")
         nlen = Len(CboAHor.List(j))
         TxAhor = Right(CboAHor.List(j), nlen - nposicion)
         If TxdeHor = TxAhor Then
            If MsgBox("<CHK> no cambio HORARIO...  continua ?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               CboAHor.Text = ""
               TxAhor = ""
               Exit Sub
             End If
         End If
         If TxAhor <> "O" Then
            '-Verifica horario Permitido en Unid. Organica
            'cpccostohor-> ccosto+horario
            Set rs = New Recordset
            sql = "SELECT * FROM CpCcostoHor WHERE ccosto='" + Txccosto + "' and horario='" & TxAhor & "'"
            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If rs.EOF Then
               MsgBox "<E> Horario no corresponde..", vbExclamation
               CboAHor.Text = ""
               TxAhor = ""
               rs.Close
               Exit Sub
            End If
            rs.Close
            
            If InStr("15_16_17_18_19_20_50_61_62", Trim(TxAhor)) > 0 Then      '18 y 24 Hrs
               'If InStr("RES_3RO", Trim(Txplh.Text)) = 0 And wDerEspec <> "S" Then
               '   MsgBox "<E> Horario no Autorizado,." + Txplh.Text, vbExclamation
               '   CboAHor.Text = ""
               '   TxAhor.Text = ""
               '   Exit Sub
               'End If
            End If
         End If
      End If
   End If
End Sub
Private Sub CboAhor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboAHor.Text) Then
      Else
         j = CboAHor.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboAHor.List(j), ">")
            nlen = Len(CboAHor.List(j))
            TxAhor = Right(CboAHor.List(j), nlen - nposicion)
            If TxdeHor = TxAhor Then
               If MsgBox("<CHK> no cambio HORARIO...  continua ?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                  CboAHor.Text = "":                TxAhor.Text = ""
                  Exit Sub
                End If
            End If
            If TxAhor <> "O" Then
               '-Verifica horario Permitido en Unid. Organica
               'cpccostohor-> ccosto+horario
               Set rs = New Recordset
               sql = "SELECT * FROM CpCcostoHor WHERE ccosto='" + Txccosto + "' and horario='" & TxAhor.Text & "'"
               rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If rs.EOF Then
                  MsgBox "<E> Horario no corresponde..", vbExclamation
                  CboAHor.Text = ""
                  TxAhor = ""
                  rs.Close
                  Exit Sub
               End If
               rs.Close
               
               If InStr("15_16_17_18_19_20_50_61_62", Trim(TxAhor)) > 0 Then   'horarios 18 y 24 hrs
                  'If InStr("RES_3RO", Trim(Txplh.Text)) = 0 And wDerEspec <> "S" Then
                  '   MsgBox "<E> Horario no Autorizado..", vbExclamation
                  '   CboAHor.Text = ""
                  '   TxAhor.Text = ""
                  '   Exit Sub
                  'End If
               End If
            End If
         End If
      End If
   End If
End Sub

Private Sub CargaCambios(nopc)
'***  CARGA  registros  ***   0:Todos   1:dia
    Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmpTabla"
    Set rs_d = New Recordset
    sq_l = "SELECT * FROM CpTmptabla"
    rs_d.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    nson = 0
    If nopc = 0 Then
       fecha1 = Format(primer, "yyyymmdd")
       fecha2 = Format(ultimo, "yyyymmdd")
       If Trim(Tx_tarjeta.Text) = "" Then
          sql = "SELECT * FROM CpChgTur" & genaño & " WHERE fecha>='" & fecha1 & "' AND fecha<='" & fecha2 & "' ORDER BY tarjeta,Fecha"
          '   sql = "SELECT    tb1.*, tb2.tarjeta as codigo, tb2.CodUPR " & _
          '   "FROM   CpChgTur" & genaño & " tb1 INNER JOIN " & _
          '   "CpGeneral tb2 ON tb2.Tarjeta = tb1.tarjeta AND tb1.FECHA >='" & Fecha1 & "' AND tb1.FECHA <='" & fecha2 & "' AND CHARINDEX(tb2.CodUPR,'" & cCadenaUpr & "') > 0" & _
          '   "ORDER BY tb2.apelpat, tb1.FECHA"
       Else
          sql = "SELECT * FROM CpChgTur" & genaño & " WHERE tarjeta='" & Trim(Tx_tarjeta.Text) & "' AND fecha>='" & fecha1 & "' AND fecha<='" & fecha2 & "' ORDER BY tarjeta,Fecha"
       End If
    Else
       fecha1 = Format(TxDia, "yyyymmdd")
       If Trim(Tx_tarjeta) = "" Then
          sql = "SELECT * FROM CpChgTur" & genaño & "  WHERE fecha<='" & fecha1 & "' AND nfecha>='" & fecha1 & "' ORDER BY tarjeta,Fecha"
          '   sql = "SELECT    tb1.*, tb2.tarjeta as codigo, tb2.CodUPR " & _
          '   "FROM   CpChgTur" & genaño & " tb1 INNER JOIN " & _
          '   "CpGeneral tb2 ON tb2.Tarjeta = tb1.tarjeta AND tb1.FECHA ='" & Fecha1 & "' AND CHARINDEX(tb2.CodUPR,'" & cCadenaUpr & "') > 0" & _
          '   "ORDER BY tb2.apelpat, tb1.FECHA"
       Else
          sql = "SELECT * FROM CpChgTur" & genaño & "  WHERE tarjeta='" & Trim(Tx_tarjeta) & "' AND fecha<='" & fecha1 & "' AND nfecha>='" & fecha1 & "' ORDER BY tarjeta,Fecha"
       End If
    End If
    
    Dim rt As Recordset
    Set rt = New Recordset
    rt.Open "SELECT * FROM cphorarionew  order by codigo", Cnx, adOpenStatic, adLockReadOnly
    '-Carga detalles del cambio
    Set r_d = New ADODB.Recordset
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not r_d.EOF
       Set rs = New Recordset
       sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & r_d!tarjeta & "'"
       rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       cTrab = ""
       TarjOk = "S"
       If rs.EOF Then
       Else
          If Tx_codupr <> "" Then
             If Tx_codupr = rs!CodUpr Then
             Else
                TarjOk = "N"
             End If
          End If
          cTrab = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
       End If
       rs.Close
       If TarjOk = "S" Then
          nson = nson + 1
          rs_d.AddNew
          rs_d!tarjeta = r_d!tarjeta
          rs_d!trabajador = cTrab
          rs_d!tipo = IIf(IsEmpty(r_d!tipo), "", r_d!tipo)
          If IsDate(r_d!fecha) Then rs_d!fechaIni = r_d!fecha
          If IsDate(r_d!nfecha) Then rs_d!fechaFin = r_d!nfecha
          rs_d!documento = IIf(IsEmpty(r_d!documento), "", r_d!documento)
          rs_d!dehor0 = IIf(IsEmpty(r_d!dehor0), "", r_d!dehor0)
          rs_d!Ahor0 = IIf(IsEmpty(r_d!Ahor0), "", r_d!Ahor0)
          rs_d!dehor = IIf(IsEmpty(r_d!dehor), "", r_d!dehor)
          rs_d!aHor = IIf(IsEmpty(r_d!aHor), "", r_d!aHor)
          rs_d!usuario = IIf(IsEmpty(r_d!usuario), "", r_d!usuario)
          'completa registro, graba turno de c/horario
          rt.MoveFirst
          Do While Not rt.EOF
             If Trim(r_d!dehor0) <> "" Then
                If rt!Codigo = r_d!dehor0 Then
                   rs_d!deTur0 = IIf(IsNull(rt!turno), "", rt!turno)
                End If
             End If
             If Trim(r_d!Ahor0) <> "" Then
                If rt!Codigo = r_d!Ahor0 Then
                   rs_d!aTur0 = IIf(IsNull(rt!turno), "", rt!turno)
                End If
             End If
             If Trim(r_d!dehor) <> "" Then
                If rt!Codigo = r_d!dehor Then
                   rs_d!Detur = IIf(IsNull(rt!turno), "", rt!turno)
                End If
             End If
             If Trim(r_d!aHor) <> "" Then
                If rt!Codigo = r_d!aHor Then
                   rs_d!Atur = IIf(IsNull(rt!turno), "", rt!turno)
                End If
             End If
             rt.MoveNext
          Loop
          rs_d.Update
       End If
       r_d.MoveNext
    Loop
    Cnx.CommitTrans
    rt.Close
    rs_d.Close
    r_d.Close
End Sub

Private Sub ChkExcepcion()
     cTurnoDAnt = "OK": cTurnoDSig = "OK"
     cturnvb = "":       cturnv = "":       Dia = 0
     FlgXGuar = False: FlgExcep = False: msgExcep = ""
     '-fechas y turnos finales
     z = Day(Txdefecha)
     d_1 = z - 1
     cturnvb = TxAtur0
     d_2 = 99
     If Trim(Txafecha) <> "" Then
        z = Day(Txafecha)
        d_2 = z - 1
        cturnv = TxAtur
     End If
       
     If InStr("3RO_RES", Txplh) = 0 Then
        If InStr("N_D_M_T_H_C_P_MT_MX", Left(cturnvb, 1)) > 0 And cturnvb <> "" And cturnvb <> "*" Then
           '-Valida Dia Anterior FechaB
           dfecha = DateValue(Txdefecha) - 1    'dia anterior ¿otro mes?
           ndia = Day(dfecha)
           z = ndia - 1
           cmes = Right("0" + Trim(Str(Month(dfecha))), 2)
           cano = Trim(Str(Year(dfecha)))
           If Txmes.Text <> cmes Then
              sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & cano & "' AND mes='" & cmes & "'"
              Set rs = New Recordset
              rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
              If Not rs.EOF Then
                 cTurnoDAnt = IIf(IsNull(rs.Fields(z + 4)), cTurnoDAnt, rs.Fields(z + 4))
              End If
              rs.Close
           Else
              cTurnoDAnt = IIf(IsNull(Textd(z)), cTurnoDAnt, Textd(z))
           End If
           If InStr("N_HCP", Left(cTurnoDAnt, 1)) > 0 And cTurnoDAnt <> "" And cTurnoDAnt <> "*" Then
              msgExcep = "DiaAnt:" + Format(dfecha, "dd/mm/yyyy") + " -> " + cTurnoDAnt + "-" + cturnvb + ","
              FlgExcep = True
                Exit Sub
           End If
        End If
        If Txtipo = "1" Then
           If InStr("N_D_M_T_HCP_MT_MX", Left(cturnv, 1)) > 0 And cturnv <> "" And cturnv <> "*" Then
              '-Valida Dia Anterior fechaNva
              dfecha = DateValue(Txafecha.Text) - 1
              ndia = Day(dfecha)
              z = ndia - 1
              If dfecha = DateValue(Txdefecha.Text) Then
                 cTurnoDAnt = TxAtur0
              Else
                 cmes = Right("0" + Trim(Str(Month(dfecha))), 2)
                 cano = Trim(Str(Year(dfecha)))
                 If Txmes.Text <> cmes Then
                    sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & cano & "' AND mes='" & cmes & "'"
                    Set rs = New Recordset
                    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                    If Not rs.EOF Then
                       cTurnoDAnt = IIf(IsNull(rs.Fields(z + 4)), cTurnoDAnt, rs.Fields(z + 4))
                    End If
                    rs.Close
                 Else
                    cTurnoDAnt = IIf(IsNull(Textd(z)), cTurnoDAnt, Textd(z))
                 End If
              End If
              If InStr("NN*N2N3N4_4_HCP_H*H1_C*C1_P*P1_NA_PA", cTurnoDAnt) > 0 And cTurnoDAnt <> "" And cTurnoDAnt <> "*" Then
                 msgExcep = "DiaAnt:" + Format(dfecha, "dd/mm/yyyy") + " -> " + cTurnoDAnt + "-" + cturnv + ","
                 FlgExcep = True
                   Exit Sub
              End If
           End If
        End If
         
        If Txtipo = "0" Then
           If InStr("NN*N2N3N4_4_HCP_H*H1_C*C1_P*P1_NA_PA", cturnvb) > 0 And cturnvb <> "" And cturnvb <> "*" Then
              '-Valida Dia Siguiente FechaB
              dfecha = DateValue(Txdefecha.Text) + 1    'ubica  dia siguiente
              ndia = Day(dfecha)
              z = ndia - 1
              cmes = Right("0" + Trim(Str(Month(dfecha))), 2)
              cano = Trim(Str(Year(dfecha)))
              If Txmes.Text = cmes Then
                 '-Next day
                 cTurnoDSig = IIf(IsNull(Textd(z)), cTurnoDSig, Textd(z))
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
              If InStr("N_D_M_T_HCP_MT_MX", Left(cTurnoDSig, 1)) > 0 And cTurnoDSig <> "" And cTurnoDSig <> "*" Then
                 msgExcep = msgExcep + "DiaSig:" + Format(dfecha, "dd/mm/yyyy") + " -> " + cturnvb + "-" + cTurnoDSig + ","
                 FlgExcep = True
                   Exit Sub
              End If
           End If
        Else
           If InStr("N_HCP", Left(cturnv, 1)) > 0 And cturnv <> "" And cturnv <> "*" Then
              '-Valida Dia Siguiente FechaNva
              dfecha = DateValue(Txafecha.Text) + 1    'ubica  dia siguiente
              ndia = Day(dfecha)
              z = ndia - 1
              cmes = Right("0" + Trim(Str(Month(dfecha))), 2)
              cano = Trim(Str(Year(dfecha)))
              If Txmes.Text = cmes Then
                 '-Next day
                 cTurnoDSig = IIf(IsNull(Textd(z)), cTurnoDSig, Textd(z))
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
              If InStr("N_D_M_T_HCP_MT_MX", Left(cTurnoDSig, 1)) > 0 And cTurnoDSig <> "" And cTurnoDSig <> "*" Then
                 msgExcep = msgExcep + "DiaSig:" + Format(dfecha, "dd/mm/yyyy") + " -> " + cturnvb + "-" + cturnv + ","
                 FlgExcep = True
                   Exit Sub
              End If
           End If
        End If
     End If
      
     dfecha = DateValue(Txdefecha.Text)     'dia base
     Dia = Day(dfecha)
     z = Dia - 1
     '-Valida NO EXCEDA guardias permitidas 150/240Hras
     If Txtipo = "0" Then
        'If (InStr("N_N*_H_C_P", Textd(z)) = 0 Or Textd(z) <> "*") And InStr("N_N*_H_C_P", TxAtur0) > 0 And TxAtur0 <> "" And TxAtur0 <> "*" Then
           '-Agrega NOCHE
        '   If Val(TxN.Text) + Val(TxD.Text) = 12 Then
        '      msgExcep = msgExcep + "Excede Guard.Permitidas,"
        '      FlgExcep = True
        '      FlgXGuar = True
        If Txcodupr <> "23" Then
           If Val(TxHoras) - Val(TxdeHrs0) + Val(TxAhrs0) > 150 Then
              msgExcep = msgExcep + "Excede Horas Permitidas,"
              FlgExcep = True
              FlgXGuar = True
              Exit Sub
           End If
        Else
           If Txcargo >= "300" And Txcargo < "400" Then
              If Val(TxHoras) - Val(TxdeHrs0) + Val(TxAhrs0) > 240 Then
                 msgExcep = msgExcep + "Excede Horas Permitidas,"
                 FlgExcep = True
                 FlgXGuar = True
                 Exit Sub
              End If
           End If
        End If
        'ElseIf (InStr("D1_D_D*", Textd(z)) = 0 Or Textd(z) <> "*") And InStr("D1_D_D*", TxAtur0) > 0 And TxAtur0 <> "" And TxAtur0 <> "*" Then
               '-Agrega DIA
        '       If Val(TxD.Text) + Val(TxN.Text) = 12 Then
        '          msgExcep = msgExcep + "Excede Guard.Permitidas,"
        '          FlgExcep = True
        '          FlgXGuar = True
        '       ElseIf Val(TxHoras.Text) - Val(TxdeHrs0.Text) + Val(TxAhrs0.Text) > 150 Then
        '          msgExcep = msgExcep + "Excede Horas Permitidas,"
        '          FlgXGuar = True
        '       End If
        'End If
     Else
        If Txcodupr <> "23" Then
           If Val(TxHoras) - Val(TxdeHrs0) + Val(TxAhrs0) - Val(TxDeHrs) + Val(TxAhrs) > 150 Then
              msgExcep = msgExcep + "Excede Horas Permitidas,"
              FlgExcep = True
              FlgXGuar = True
              Exit Sub
           End If
        Else
           If Txcargo >= "300" And Txcargo < "400" Then
              If Val(TxHoras) - Val(TxdeHrs0) + Val(TxAhrs0) - Val(TxDeHrs) + Val(TxAhrs) > 240 Then
                 msgExcep = msgExcep + "Excede Horas Permitidas,"
                 FlgExcep = True
                 FlgXGuar = True
                 Exit Sub
              End If
           End If
        End If
        'If InStr("N_N*_H_C_P", TxAtur) > 0 And TxAtur <> "" And TxAtur <> "*" Then
        '   If InStr("N_N*_H_C_P", TxDetur0) > 0 And TxDetur0 <> "" And TxDetur0 <> "*" Then
               '-Desplaza turno NOCHE
        '   Else
              '-Agrega NOCHE
        '      If Val(TxN.Text) + Val(TxD.Text) = 12 Then
        '         msgExcep = msgExcep + "Excede Guard.Permitidas,"
        '         FlgExcep = True
        '      End If
        '   End If
        'End If
        'If InStr("D_D1_D*", TxAtur) > 0 And TxAtur <> "" And TxAtur <> "*" Then
        '   If InStr("D_D1_D*", TxDetur0) > 0 And TxDetur0 <> "" And TxDetur0 <> "*" Then
              '-Desplaza turno DIA
        '   Else
              '-Agrega DIA
        '      If Val(TxD.Text) + Val(TxN.Text) = 12 Then
        '         msgExcep = msgExcep + "Excede Guard.Permitidas,"
        '         FlgExcep = True
        '      End If
        '   End If
        'End If
     End If
         
     If InStr("3RO_RES", Txplh) = 0 Then
        '-Verifica Excede Noches Permitidas
        If Txtipo = "0" Then
           If (InStr("NN*N2N3N4_HCP_H*H1_C*C1_P*P1_NA_PA", Textd(z)) = 0 Or Textd(z) <> "*") And InStr("NN*_HCP_H*_C*_P*", TxAtur0) > 0 And TxAtur0 <> "" And TxAtur0 <> "*" Then
              '-Agrega Noche
              If Val(TxN) = GenMaxTN Then
                 msgExcep = msgExcep + "Excede " + Trim(Str(GenMaxTN)) + " Noches,"
                 FlgExcep = True
                   Exit Sub
               End If
           End If
        Else
           '--Tipo 1 check   11/08 0->06    12/08 05->0     agrega NOCHE
           If InStr("NN*N2N3N4_HCP_H*H1_C*C1_P*P1_NA_PA", TxAtur) > 0 And TxAtur <> "" And TxAtur <> "*" Then
              '-Registra Noche en dia Nuevo
              If InStr("NN*N2N3N4_HCP_H*H1_C*C1_P*P1_NA_PA", TxDetur0) > 0 And TxDetur0 <> "" And TxDetur0 <> "*" Then
                 '-Desplaza turno Noche
              Else
                 If InStr("NN*N2N3N4_HCP_H*H1_C*C1|_P*P1_NA_PA", TxDeTur) > 0 And TxDeTur <> "" And TxDeTur <> "*" Then
                    '-Cambia turno Noche dia Nuevo n -> H/C/P viceversa
                 Else
                    '-Agrega Noche dia Nuevo
                    If Val(TxN) = GenMaxTN Then
                       msgExcep = msgExcep + "Excede " + Trim(Str(GenMaxTN)) + " Noches,"
                       FlgExcep = True
                         Exit Sub
                    End If
                 End If
              End If
           Else
              If InStr("NN*N2N3N4_HCP_H*H1_C*C1_P*P1_NA_PA", TxAtur0) > 0 And TxAtur0 <> "" And TxAtur0 <> "*" Then
                 If InStr("NN*N2N3N4_HCP_H*H1_C*C1_P*P1_NA_PA", TxDetur0) > 0 And TxDetur0 <> "" And TxDetur0 <> "*" Then
                    '-Desplaza turno Noche
                 Else
                    '-agrega Noche en dia Base
                    If Val(TxN) = GenMaxTN Then
                       msgExcep = msgExcep + "Excede " + Trim(Str(GenMaxTN)) + " Noches,"
                       FlgExcep = True
                         Exit Sub
                    End If
                 End If
              End If
           End If
        End If
     End If
      
     nposicion = 0
     cvalor = ""
     If InStr("3RO_RES", Txplh) = 0 Then
        '-Valida Max dias Libres      acuerdo 19.03.15 D.A. de 3 => 4   GENMAXDL
        'If InStr("NN*_DD1_D*_MM1_TT1_HCP_H*_C*_P*_MT_MX", TxAtur0) > 0 And TxAtur0 <> "*" And TxAtur0 <> "" Then
        
           '-Rastrea desde ultm dias MES Anterior
           '======================================
           dima = primer - GenMaxDL
           dfma = primer - 1
           cyear = Trim(Str(Year(dfma)))
           cmonth = Right("00" + Trim(Str(Month(dfma))), 2)
           sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & cyear & "' AND mes='" & cmonth & "'"
           Set rs_d = New Recordset
           rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
           If Not rs_d.EOF Then
            For i = Day(dima) To Day(dfma)
              z = i - 1
              T_xd = IIf(Trim(rs_d.Fields(z + 4)) = "", "X", rs_d.Fields(z + 4))
              If T_xd = "*" Or T_xd = "X" Then
                 nposicion = nposicion + 1
                 cvalor = cvalor + Trim(Str(i)) + "-"
              ElseIf (T_xd <> "*" And Trim(T_xd) <> "X") Then
                     nposicion = 0
                     cvalor = ""
              End If
              
              'MsgBox ("dia:" + Trim(Str(i)) + "->" + T_xd + "[" + Str(nposicion) + "]")
            Next i
           End If
           rs_d.Close
           If nposicion = GenMaxDL Then
              'MsgBox ("mes Anterior, BLOQUE de Dias " + cvalor)
              'MsgBox "BLOQUE Dias Libres " + cvalor + ", es Maximo permitido... ", vbinformation
           End If
           'cvalor = ""
           For i = Day(primer) To Day(ultimo)
              '-Rastrear bloques dias mes ACTUAL simulando cambios efectuados
              z = i - 1
              c_turno = ""
              If z = d_1 Then
                 c_turno = cturnvb
              ElseIf z = d_2 Then
                     c_turno = cturnv
              Else
                  c_turno = Textd(z)
              End If
              If c_turno = "*" Or Trim(c_turno) = "" Then
                 nposicion = nposicion + 1
                 cvalor = cvalor + Trim(Str(i)) + "-"
              ElseIf (c_turno <> "*" And Trim(c_turno) <> "") Then
                     nposicion = 0
                     cvalor = ""
              End If
              
              'MsgBox Trim(Str(i)) + ": BLOQUE Libre [" + cvalor + "]"
                            
              If nposicion > GenMaxDL Then
                 'If MsgBox("BLOQUE " + cvalor + " Excede Dias Libres... GRABA ", vbOKCancel + vbQuestion, "Confirme") = vbCancel Then
                 Beep 500, 350
                 'MsgBox "BLOQUE Dias Libres " + cvalor + ", Excede lo permitido... ", wvcritical
                 msgExcep = "Cadena " + cvalor + ", Excede dias libres permitido."
                 nposicion = 99
                 Exit For
              End If
              ''' MsgBox (" dias libres " + Str(nposicion) + " al: " + Str(i))
      
           Next i
        '//End If
     End If
     If nposicion > GenMaxDL Then    '4
        FlgExcep = True
        Exit Sub
     End If
     
End Sub

