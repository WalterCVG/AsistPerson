VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ProgEspecial 
   Caption         =   "Directorio - Especiales"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15075
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "ProgEspecial.frx":0000
   ScaleHeight     =   10515
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxCATEGORIA 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   190
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton CmdAudDirect 
      Height          =   555
      Left            =   12360
      Picture         =   "ProgEspecial.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   189
      ToolTipText     =   "Auditar cambios"
      Top             =   9480
      Width           =   645
   End
   Begin VB.TextBox Txcadhep 
      Enabled         =   0   'False
      Height          =   285
      Left            =   14040
      MaxLength       =   52
      TabIndex        =   181
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdLIMPIA 
      Caption         =   "Limpia"
      Height          =   615
      Left            =   14040
      Picture         =   "ProgEspecial.frx":80B30
      Style           =   1  'Graphical
      TabIndex        =   158
      ToolTipText     =   "Limpia Plantilla"
      Top             =   8520
      Width           =   735
   End
   Begin VB.TextBox Txbuscar1 
      Height          =   285
      Left            =   12120
      MaxLength       =   12
      TabIndex        =   10
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Txbuscar 
      Height          =   285
      Left            =   10200
      MaxLength       =   5
      TabIndex        =   9
      Top             =   4560
      Width           =   735
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TxTipoPla 
      Enabled         =   0   'False
      Height          =   285
      Left            =   11880
      MaxLength       =   3
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Tx_ccosto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   11520
      MaxLength       =   2
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Cbo_ccosto 
      Height          =   315
      Left            =   6840
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox Txdscccosto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox TxCarEst 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10920
      MaxLength       =   2
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Txccosto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9720
      MaxLength       =   2
      TabIndex        =   16
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxPlH 
      BorderStyle     =   0  'None
      ForeColor       =   &H00004080&
      Height          =   285
      Left            =   10440
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   18
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxCargo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   11280
      MaxLength       =   4
      TabIndex        =   20
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Tx_grupo 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   13320
      MaxLength       =   4
      TabIndex        =   22
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox TxcodUPR 
      Enabled         =   0   'False
      Height          =   285
      Left            =   10080
      MaxLength       =   2
      TabIndex        =   17
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxDscCargo 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   35
      TabIndex        =   15
      Top             =   5160
      Width           =   2775
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
      Left            =   4680
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   240
      TabIndex        =   157
      Top             =   5520
      Width           =   13575
      Begin VB.TextBox TxPA 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   185
         Top             =   3360
         Width           =   255
      End
      Begin VB.TextBox TxNA 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   184
         Top             =   3360
         Width           =   255
      End
      Begin VB.TextBox TxMT 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   182
         Top             =   3360
         Width           =   255
      End
      Begin VB.CommandButton CmdAplica 
         Caption         =   "Aplica"
         Height          =   375
         Left            =   7440
         TabIndex        =   151
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox TxnHra 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         TabIndex        =   150
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxHorario 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         MaxLength       =   4
         TabIndex        =   148
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox CboHorario 
         Height          =   315
         Left            =   1440
         TabIndex        =   149
         Top             =   3360
         Width           =   5295
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
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   156
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox TxD 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   152
         Top             =   3360
         Width           =   255
      End
      Begin VB.TextBox TxN 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   153
         Top             =   3360
         Width           =   255
      End
      Begin VB.TextBox TxM 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   154
         Top             =   3360
         Width           =   255
      End
      Begin VB.TextBox Txt 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   11880
         Locked          =   -1  'True
         TabIndex        =   155
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   30
         Left            =   9720
         TabIndex        =   54
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   29
         Left            =   8520
         TabIndex        =   53
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   28
         Left            =   7320
         TabIndex        =   52
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   27
         Left            =   6120
         TabIndex        =   51
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   26
         Left            =   4920
         TabIndex        =   50
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   25
         Left            =   3720
         TabIndex        =   164
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   24
         Left            =   2520
         TabIndex        =   48
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   0
         Left            =   120
         MaxLength       =   2
         TabIndex        =   55
         Text            =   "Text1"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   1
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   58
         Text            =   "Text2"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   2
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   61
         Text            =   "Text3"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   3
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   64
         Text            =   "Text4"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   4
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   5
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   70
         Text            =   "Text6"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   6
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   73
         Text            =   "Text7"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   7
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   76
         Text            =   "Text8"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   8
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   79
         Text            =   "Text9"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   9
         Left            =   10920
         MaxLength       =   2
         TabIndex        =   82
         Text            =   "Text10"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   10
         Left            =   12120
         MaxLength       =   2
         TabIndex        =   85
         Text            =   "Text11"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   11
         Left            =   120
         MaxLength       =   2
         TabIndex        =   88
         Text            =   "Text12"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   12
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   91
         Text            =   "Text13"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   13
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   94
         Text            =   "Text14"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   14
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   95
         Text            =   "Text15"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   15
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   98
         Text            =   "Text16"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   16
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   101
         Text            =   "Text17"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   17
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   104
         Text            =   "Text18"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   18
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   107
         Text            =   "Text19"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   19
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   110
         Text            =   "Text20"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   20
         Left            =   10920
         MaxLength       =   2
         TabIndex        =   113
         Text            =   "Text21"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   21
         Left            =   12120
         MaxLength       =   2
         TabIndex        =   116
         Text            =   "Text22"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   22
         Left            =   120
         MaxLength       =   2
         TabIndex        =   119
         Text            =   "Text23"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   23
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   122
         Text            =   "Text24"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   24
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   124
         Text            =   "Text25"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   25
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   127
         Text            =   "Text26"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   26
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   130
         Text            =   "Text27"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   27
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   133
         Text            =   "Text28"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   28
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   136
         Text            =   "Text29"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Textd 
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   29
         Left            =   8520
         MaxLength       =   2
         TabIndex        =   139
         Text            =   "Text30"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Textd 
         Height          =   375
         Index           =   30
         Left            =   9720
         MaxLength       =   2
         TabIndex        =   142
         Text            =   "Text31"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox TxTurno 
         Height          =   285
         Left            =   960
         MaxLength       =   2
         TabIndex        =   146
         Top             =   3000
         Width           =   375
      End
      Begin VB.ComboBox CboTurDir 
         Height          =   315
         Left            =   1440
         TabIndex        =   147
         Top             =   3000
         Width           =   5295
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   855
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
         Height          =   315
         Index           =   30
         Left            =   10080
         MaxLength       =   4
         TabIndex        =   143
         Text            =   "Text1"
         Top             =   2280
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
         Height          =   315
         Index           =   29
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   140
         Text            =   "Text1"
         Top             =   2280
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
         Height          =   315
         Index           =   28
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   137
         Text            =   "Text1"
         Top             =   2280
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
         Height          =   315
         Index           =   27
         Left            =   6480
         MaxLength       =   4
         TabIndex        =   134
         Text            =   "Text1"
         Top             =   2280
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
         Height          =   315
         Index           =   26
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   131
         Text            =   "Text1"
         Top             =   2280
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
         Height          =   315
         Index           =   25
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   128
         Text            =   "Text1"
         Top             =   2280
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
         Height          =   315
         Index           =   24
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   125
         Text            =   "Text1"
         Top             =   2280
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
         Height          =   315
         Index           =   23
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   160
         Text            =   "Text1"
         Top             =   2280
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
         Height          =   315
         Index           =   22
         Left            =   480
         MaxLength       =   4
         TabIndex        =   120
         Text            =   "Text1"
         Top             =   2280
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
         Height          =   285
         Index           =   21
         Left            =   12480
         MaxLength       =   4
         TabIndex        =   117
         Text            =   "Text1"
         Top             =   1440
         Width           =   375
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
         Height          =   315
         Index           =   20
         Left            =   11280
         MaxLength       =   4
         TabIndex        =   114
         Text            =   "Text1"
         Top             =   1440
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
         Height          =   315
         Index           =   19
         Left            =   10080
         MaxLength       =   4
         TabIndex        =   111
         Text            =   "Text1"
         Top             =   1440
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
         Height          =   315
         Index           =   18
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   108
         Text            =   "Text1"
         Top             =   1440
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
         Height          =   315
         Index           =   17
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   105
         Text            =   "Text1"
         Top             =   1440
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
         Height          =   315
         Index           =   16
         Left            =   6480
         MaxLength       =   4
         TabIndex        =   102
         Text            =   "Text1"
         Top             =   1440
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
         Height          =   315
         Index           =   15
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   99
         Text            =   "Text1"
         Top             =   1440
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
         Height          =   315
         Index           =   14
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   96
         Text            =   "Text1"
         Top             =   1440
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
         Height          =   315
         Index           =   13
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   159
         Text            =   "Text1"
         Top             =   1440
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
         Height          =   315
         Index           =   12
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   1440
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
         Height          =   315
         Index           =   11
         Left            =   480
         MaxLength       =   4
         TabIndex        =   89
         Text            =   "Text1"
         Top             =   1440
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
         Height          =   285
         Index           =   10
         Left            =   12480
         MaxLength       =   4
         TabIndex        =   86
         Text            =   "Text1"
         Top             =   600
         Width           =   375
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
         Height          =   315
         Index           =   9
         Left            =   11280
         MaxLength       =   4
         TabIndex        =   83
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
         Height          =   315
         Index           =   8
         Left            =   10080
         MaxLength       =   4
         TabIndex        =   80
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
         Height          =   315
         Index           =   7
         Left            =   8880
         MaxLength       =   4
         TabIndex        =   77
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
         Height          =   315
         Index           =   6
         Left            =   7680
         MaxLength       =   4
         TabIndex        =   74
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
         Height          =   315
         Index           =   5
         Left            =   6480
         MaxLength       =   4
         TabIndex        =   71
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
         Height          =   315
         Index           =   4
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   68
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
         Height          =   315
         Index           =   3
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   65
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
         Height          =   315
         Index           =   2
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   62
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
         Height          =   315
         Index           =   1
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   59
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
         Height          =   315
         Index           =   0
         Left            =   480
         MaxLength       =   4
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   600
         Width           =   355
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   4
         Left            =   4920
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   5
         Left            =   6120
         TabIndex        =   29
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   6
         Left            =   7320
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   7
         Left            =   8520
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   8
         Left            =   9720
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   9
         Left            =   10920
         TabIndex        =   33
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   10
         Left            =   12120
         TabIndex        =   34
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   12
         Left            =   1320
         TabIndex        =   36
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   13
         Left            =   2520
         TabIndex        =   37
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   14
         Left            =   3720
         TabIndex        =   38
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   15
         Left            =   4920
         TabIndex        =   39
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   16
         Left            =   6120
         TabIndex        =   40
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   17
         Left            =   7320
         TabIndex        =   41
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   18
         Left            =   8520
         TabIndex        =   42
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   19
         Left            =   9720
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   20
         Left            =   10920
         TabIndex        =   44
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   21
         Left            =   12120
         TabIndex        =   45
         Top             =   1080
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   22
         Left            =   120
         TabIndex        =   46
         Top             =   1920
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   375
         Index           =   23
         Left            =   1320
         TabIndex        =   47
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "PA"
         Height          =   255
         Left            =   10440
         TabIndex        =   187
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "NA"
         Height          =   255
         Left            =   10080
         TabIndex        =   186
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "MT"
         Height          =   255
         Left            =   9720
         TabIndex        =   183
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Horario"
         Height          =   255
         Left            =   120
         TabIndex        =   178
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "HORAS Prog."
         Height          =   255
         Left            =   12240
         TabIndex        =   177
         Top             =   3120
         Width           =   1095
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
         Height          =   255
         Left            =   8520
         TabIndex        =   176
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "D"
         Height          =   255
         Left            =   10800
         TabIndex        =   175
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "N"
         Height          =   255
         Left            =   11160
         TabIndex        =   174
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "T"
         Height          =   255
         Left            =   11880
         TabIndex        =   173
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "M"
         Height          =   255
         Left            =   11520
         TabIndex        =   172
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Turno Directorio"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   166
         Top             =   2880
         Width           =   855
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
         TabIndex        =   144
         Top             =   2280
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
         TabIndex        =   141
         Top             =   2280
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
         TabIndex        =   138
         Top             =   2280
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
         TabIndex        =   135
         Top             =   2280
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
         TabIndex        =   132
         Top             =   2280
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
         TabIndex        =   129
         Top             =   2280
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
         TabIndex        =   126
         Top             =   2280
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
         TabIndex        =   123
         Top             =   2280
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
         TabIndex        =   121
         Top             =   2280
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
         TabIndex        =   118
         Top             =   1440
         Width           =   615
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
         TabIndex        =   115
         Top             =   1440
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
         TabIndex        =   112
         Top             =   1440
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
         TabIndex        =   109
         Top             =   1440
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
         TabIndex        =   106
         Top             =   1440
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
         TabIndex        =   103
         Top             =   1440
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
         TabIndex        =   100
         Top             =   1440
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
         TabIndex        =   97
         Top             =   1440
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
         TabIndex        =   162
         Top             =   1440
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
         TabIndex        =   93
         Top             =   1440
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
         TabIndex        =   90
         Top             =   1440
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
         TabIndex        =   87
         Top             =   600
         Width           =   495
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
         TabIndex        =   84
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
         TabIndex        =   81
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
         TabIndex        =   78
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
         TabIndex        =   75
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
         TabIndex        =   72
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
         TabIndex        =   69
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
         TabIndex        =   66
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
         TabIndex        =   63
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
         TabIndex        =   60
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
         TabIndex        =   57
         Top             =   600
         Width           =   390
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   9480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton mnuFicImpSelec 
      Height          =   495
      Left            =   9960
      Picture         =   "ProgEspecial.frx":81062
      Style           =   1  'Graphical
      TabIndex        =   165
      ToolTipText     =   "Selec.Impresora"
      Top             =   9480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "ProgEspecial.frx":81840
      Style           =   1  'Graphical
      TabIndex        =   145
      Top             =   120
      Width           =   855
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
      Left            =   11040
      Picture         =   "ProgEspecial.frx":82482
      Style           =   1  'Graphical
      TabIndex        =   167
      ToolTipText     =   "Directorio"
      Top             =   9480
      Width           =   765
   End
   Begin VB.CommandButton CmdClear 
      Height          =   615
      Left            =   14160
      Picture         =   "ProgEspecial.frx":82D4C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Otra Consulta"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   615
      Left            =   12120
      Picture         =   "ProgEspecial.frx":83616
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ubicar Registros"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Txtarjeta 
      Height          =   285
      Left            =   1200
      MaxLength       =   5
      TabIndex        =   13
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   195
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox Txtrabajador 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5160
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   7320
      TabIndex        =   11
      Top             =   9360
      Width           =   2055
      Begin VB.CommandButton cmdGraba 
         Caption         =   "Graba"
         Height          =   615
         Left            =   1200
         Picture         =   "ProgEspecial.frx":837F1
         Style           =   1  'Graphical
         TabIndex        =   163
         ToolTipText     =   "Grabar"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdElimina 
         Caption         =   "Elimina"
         Height          =   615
         Left            =   120
         Picture         =   "ProgEspecial.frx":83963
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   14040
      Picture         =   "ProgEspecial.frx":83AD5
      Style           =   1  'Graphical
      TabIndex        =   170
      ToolTipText     =   "Cerrar"
      Top             =   9840
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   6165
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
      Caption         =   "Directorio de trabajadores"
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
   Begin VB.Label LblCATEGORIA 
      BackStyle       =   0  'Transparent
      Caption         =   "LblCATEGORIA"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1560
      TabIndex        =   191
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label LblDS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LblDS"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   11880
      TabIndex        =   188
      Top             =   10200
      Width           =   2055
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Apel. Paterno"
      Height          =   255
      Left            =   11040
      TabIndex        =   180
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarjeta"
      Height          =   255
      Left            =   9600
      TabIndex        =   179
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad Org."
      Height          =   255
      Left            =   5880
      TabIndex        =   171
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Hor.Grupo"
      Height          =   255
      Left            =   12480
      TabIndex        =   169
      Top             =   5160
      Width           =   735
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
      Left            =   6480
      TabIndex        =   168
      Top             =   4800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label_2 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   255
      Left            =   240
      TabIndex        =   49
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "ProgEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DD As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
 Dim rs_2 As New ADODB.Recordset
Dim Sw, f_sigue, F_add As Boolean
Dim nposicion, ndia, dfi, dff, z As Integer
Dim sql, cTarjeta, cNombre, cano, cmes, cCarest, cDscCargo, CCosto, cDscCosto, cTurno, cBloque, cOrden, cTurSINHor As String
Dim unafecha As Date
Dim fecCierre As Date
Dim xCheckBox As Control
Private Sub Command1_Click()
     c_grupo = "02": c_secuencia = "028"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub
Private Sub Form_Load()
  Me.Caption = "DIRECTORIO Mensual-Especiales - " & genaño & "   -   Sesión: " + wUsuNombre
  LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
  
  '-Variable de tipo control
    'Dim xCheckBox As Control
  '-Recorrer en un bucle todos los controles del formulario actual
  For Each xCheckBox In Me.Controls
      '-Comprobar con el operador TypeOf si es un control CheckBox
      If TypeOf xCheckBox Is CheckBox Then
         '-Cambia la propiedades value, enabled
         xCheckBox.Value = 0
         xCheckBox.Enabled = False
      End If
  Next
  For i = 1 To 31
    z = i - 1
    Textd(z).Text = ""
    Textd(z).ForeColor = &H0&
    Textd(z).Locked = True
    Textp(z).Text = "":   Textp(z).ForeColor = &H0&
    Lblhr(z).Caption = "":    Lblhr(z).ForeColor = &H0&
    Textp(z).Locked = True
  Next
  TxTurno = "": TxTurno.Locked = True
  CboTurDir.Text = "": CboTurDir.Enabled = False
  TxHorario = "": TxHorario.Locked = True
  CboHorario.Text = "": CboHorario.Enabled = False
  Set rs = New Recordset
  rs.Open "SELECT * FROM cptablas WHERE tabla='32' AND codigo<>'..' ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
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
  Set rs = New Recordset
  rs.Open "SELECT * FROM cptablas WHERE tabla='31' and codigo<>'..' and (bloqueo='N' or Bloqueo is NULL) ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To CboTurDir.ListCount - 1
         If CboTurDir.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        CboTurDir.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
     End If
     rs.MoveNext
  Wend
  rs.Close
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpHorarioNew where (bloqueo='N' or bloqueo is NULL) ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly    'CpHorarios
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
  Call creadd
  
  Txtarjeta.Locked = True
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
            TxNombMes.Text = Left(CboMes.List(j), nposicion - 1)
            unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)            'primer/ultimo
            Call Cargadias
            fecCierre = ultimo
            Set rs = New Recordset
            rs.Open "SELECT * FROM CpCronograma WHERE ano='" & genaño & "' and Mes='" & Txmes.Text & "'", Cnx, adOpenStatic, adLockReadOnly   'CpHorarios
            If rs.EOF Then
               MsgBox "No Existe CRONOGRAMA del MES... Verifique", vbExclamation
            Else
               If IsDate(fecCierre) Then fecCierre = rs!cierremes
               If fecCierre <= Date Then
                  Me.Caption = "DIRECTORIO Mensual-Especiales... CERRADO. - " & genaño & "   -   Sesión: " + wUsuNombre
               Else
                  Me.Caption = "DIRECTORIO Mensual-Especiales - " & genaño & "   -   Sesión: " + wUsuNombre
               End If
            End If
            rs.Close
         End If
      End If
      Exit Sub
   End If
End Sub
Private Sub Cargadias()
   For Each xCheckBox In Me.Controls
        '- Comprobar con el operador TypeOf si es un control CheckBox
        If TypeOf xCheckBox Is CheckBox Then
           '- Cambia la propiedad value, caption
           xCheckBox.Value = 0
           xCheckBox.Caption = ""
        End If
   Next
   For i = 1 To 31
      z = i - 1
      Textd(z).Text = "":       Textp(z).Text = ""
      Lblhr(z).Caption = ""
   Next
   unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
   Call ObtenerPrimerUltimoDia(unafecha)
   unafecha = primer
   dfi = Day(primer)
   dff = Day(ultimo)
   For i = dfi To dff
       z = i - 1
       Check1(z).Visible = True
       Check1(z).Enabled = True
       Check1(z).Caption = Left(WeekdayName(Weekday(primer + z)), 3) + " " & i
       Check1(z).ForeColor = &H80000012      'negro
       '''If Weekday(unafecha + i - 1) = 1 Then
       If UCase(Left(WeekdayName(Weekday(primer + z)), 3)) = "DOM" Then
          Check1(z).ForeColor = &HFF               'rojo
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
                Check1(z).ForeColor = &HFF
             End If
             rs.MoveNext
          Wend
          rs.Close
       End If
   Next
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
         Txdscccosto.Text = Left(Cbo_ccosto.List(j), nposicion - 1)
         Call cmdcarga_Click
      End If
      Exit Sub
   End If
End Sub

Private Sub cmdcarga_Click()
   If Txmes.Text = "" Then
       MsgBox "No Selecciono mes ", vbExclamation
       CboMes.SetFocus
      Exit Sub
   End If
   dfi = Day(primer)
   dff = Day(ultimo)
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
     .Fields.Append "MT", adChar, 2     '<===========
     .Fields.Append "NA", adChar, 2     '<===========
     .Fields.Append "PA", adChar, 2     '<===========
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
            '.NumberFormat = "##.#0"
            .Width = Me.TextWidth(String(6, "S"))
          Case Else
            .Width = Me.TextWidth(String(3, "S"))
        End Select
        '.NumberFormat = "##,##0.0"
        .Locked = True
     End With
  Next
  '-carga  DIRECTORIO "TURNO"
  sql = "SELECT  * FROM CpDirectorio  WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'  and ccosto='" & Tx_ccosto.Text & "' " & _
        "ORDER BY Tarjeta"
  Set rs = New Recordset
  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  If Not rs.EOF Then
     Do While Not rs.EOF
       cTarjeta = rs!tarjeta
       cNombre = "":   cCarest = "":  cCargo = "":  cDscCargo = "":       cGrpGen = "":      cplh = ""
       SQ1 = "SELECT * FROM Cpgeneral  WHERE  (baja is NULL) and tarjeta='" & cTarjeta & "'"
       Set rs_d = New Recordset
       rs_d.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       f_sigue = False
       If rs_d.EOF Then
       Else
          f_sigue = True
          cNombre = Left(Trim(rs_d!apelpat) + " " + Trim(rs_d!apelmat) + " " + Trim(rs_d!nombres), 25)
          cCarest = IIf(IsNull(rs_d!carEst), "", rs_d!carEst)
          cCargo = IIf(IsNull(rs_d!Cargo), "", rs_d!Cargo)
          cGrpGen = IIf(IsNull(rs_d!grpgen), "", rs_d!grpgen)
          cplh = IIf(IsNull(rs_d!PLH), "", rs_d!PLH)
       End If
       rs_d.Close
       If f_sigue Then
          Set rs_d = New Recordset
          rs_d.Open "SELECT * FROM Cpcargos WHERE codigo='" & cCargo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs_d.EOF Then
             cDscCargo = Left(rs_d!descripcion, 15)
          End If
          rs_d.Close
          DD.AddNew
          DD!tarjeta = cTarjeta
          DD!nombre = cNombre
          DD!Cargo = cDscCargo
          DD!d = IIf(IsNull(rs!d), "", rs!d)
          DD!n = IIf(IsNull(rs!n), "", rs!n)
          DD!m = IIf(IsNull(rs!m), "", rs!m)
          DD!T = IIf(IsNull(rs!T), "", rs!T)
          DD!Horas = IIf(IsNull(rs!Horas), 0, rs!Horas)
          For i = Day(primer) To Day(ultimo)
             z = i - 1
             '- directorio rs.fields(Z+4)  ==>    dd.fields(Z+8)  ==> CpTmpDCosto(Z+6)
             'cCampo = rs.Fields(z + 4).Name
             DD.Fields(z + 8) = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
          Next i
          DD!mT = IIf(IsNull(rs!mT), "", rs!mT)
          DD!na = IIf(IsNull(rs!na), "", rs!na)
          DD!pa = IIf(IsNull(rs!pa), "", rs!pa)
          DD.Update
       End If
       rs.MoveNext
     Loop
  End If
  rs.Close
  DataGrid1.Refresh
  Call CmdLimpia_Click    'Limpiacampos
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Txtarjeta.Text = DataGrid1.Columns(0)
   LblONOMASTICO.Caption = "Onomastico"
   LblONOMASTICO.Visible = False
   cClickDTG = "S"
   Call Txtarjeta_KeyPress(13)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   With DD      ' Adodc1.Recordset
      If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
         .Sort = .Fields(ColIndex).[Name] & " Desc"
      Else
         .Sort = .Fields(ColIndex).[Name] & " Asc"
      End If
   End With
   Call CmdLimpia_Click 'Limpiacampos
End Sub
Private Sub CmdLimpia_Click()
    Txbuscar = "": Txbuscar1 = ""
    Txtarjeta = "": Txtrabajador = "": Txfecnac = ""
    Txccosto = "": TxcodUPR = "": TxPlH = "":  TxCarEst = "": TxCargo = "": Txtipopla = "": TxDscCargo = ""
    TxD = "": TxN = "": TxM = "": Txt = "": TxNa = "": TxPa.Text = "": TxHoras = "":  TxMT = ""
    TxPlH.Visible = False
    CboTurDir.Text = "": TxTurno = ""
    TxCategoria = "":  LblCATEGORIA.Caption = ""
    CboTurDir.Enabled = False: TxTurno.Locked = True
    CboHorario.Text = "": TxHorario = "": TxnHra = ""
    CboHorario.Enabled = False: TxHorario.Locked = True
    For i = dfi To dff
       z = i - 1
       Check1(z).Value = 0
       Textd(z).Visible = True: Textp(z).Visible = True
       Textd(z).Text = "": Textd(z).ForeColor = &H80000012  '&HFF0000
       Textp(z).Text = "":  Textp(z).ForeColor = &H80000012 '&HFF0000
       Lblhr(z).Caption = ""
    Next
    DataGrid1.SetFocus
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
      If Tx_ccosto = "" Then
         MsgBox "<E> Indique Unid. Organica ..", vbExclamation
         Exit Sub
      End If
      Frame1.Caption = "Programa - TRABAJADOR"
      LblONOMASTICO.Caption = "Onomastico"
      LblONOMASTICO.Visible = False
      If Trim(Txtarjeta) = "" Then
         Load Ayuda1
         Ayuda1.Tag = 6.3
         Ayuda1.Show 1
      Else
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "Tarjeta, No Ubicada", vbExclamation
            Txtarjeta = "": Txtrabajador = ""
            Tx_grupo = "": Txccosto = "": TxcodUPR = ""
            TxCategoria = "": LblCATEGORIA.Caption = ""
         Else
            If rs!tipoPla = "ASI" Or rs!tipoPla = "ADM" Then
               Txtrabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
               Txccosto = IIf(IsNull(rs!CCosto), "", rs!CCosto)
               If rs!tipoPla = "ASI" Then
                  TxcodUPR = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)
               End If
               TxCarEst = IIf(IsNull(rs!carEst), "", rs!carEst)
               TxPlH = IIf(IsNull(rs!PLH), "", rs!PLH)
               TxCargo = IIf(IsNull(rs!Cargo), "", rs!Cargo)
               Tx_grupo = IIf(IsNull(rs!grpgen), "", rs!grpgen)
               If IsDate(rs!fecnac) Then Txfecnac = rs!fecnac
               Txtipopla = IIf(IsNull(rs!tipoPla), "", rs!tipoPla)
               TxCategoria = IIf(IsNull(rs!Categoria), "", rs!Categoria)
               LblCATEGORIA.Caption = IIf(InStr("IR", TxCategoria), IIf(TxCategoria = "I", "INTERNO", "RESIDENTE"), "")
               If TxPlH = "3RO" Then
                  TxPlH.Visible = True
               Else
                  TxPlH.Visible = False
               End If
               If rs!tipoPla = "ASI" Then
                  '-Cadena  a que pertenece Servidor
                  TxCadHep = IIf(IsNull(rs!CadHep), "", Trim(rs!CadHep))
               End If
               TxCategoria = IIf(IsNull(rs!Categoria), "", rs!Categoria)
               LblCATEGORIA.Caption = IIf(InStr("IR", TxCategoria), IIf(TxCategoria = "I", "INTERNO", "RESIDENTE"), "")
            Else
               MsgBox "No definio Tipo TRABAJADOR..", vbExclamation
               Txtarjeta = "": Txtrabajador = ""
               Tx_grupo = "": Txccosto = "": TxcodUPR = ""
            End If
         End If
         rs.Close
      End If
      If Txtarjeta <> "" Then
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "Tarjeta, No Ubicada", vbExclamation
            Txtarjeta = "": Txtrabajador = "": Txfecnac = ""
            TxCarEst = "": Txccosto = "": TxcodUPR = "": TxPlH = "": Txtipopla = "":   TxCargo = "":  TxDscCargo = ""
            LblONOMASTICO.Caption = "Onomastico"
            LblONOMASTICO.Visible = False
            TxCategoria = "":               LblCATEGORIA.Caption = ""
            rs.Close
            Exit Sub
         End If
         rs.Close
         If IsDate(Txfecnac) Then
            If Month(Txfecnac.Text) = Val(Txmes) Then
               LblONOMASTICO.Visible = True
               LblONOMASTICO.Caption = "Onomastico el " + Str(Day(Txfecnac))
            End If
         End If
         Set rs_d = New Recordset
         rs_d.Open "SELECT * FROM Cpcargos WHERE  codigo='" & TxCargo & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs_d.EOF Then
            TxDscCargo = Left(rs_d!descripcion, 15)
         End If
         rs_d.Close
         '-Caso Filtro Operador, pertenece a un wccosto
         f_sigue = True
         If Txccosto <> Tx_ccosto Then
            f_sigue = False
         End If
         If f_sigue Then
         Else
            MsgBox "Trabajador, No pertenece al unidad Organizacional..!", vbExclamation
            Txtarjeta = "": Txtrabajador = "": Txfecnac = ""
            TxCarEst = "": Txccosto = "": TxcodUPR = "": TxPlH = "": Txtipopla = "":  TxCargo = "":  TxDscCargo = ""
            LblONOMASTICO.Caption = "Onomastico"
            LblONOMASTICO.Visible = False
            TxCategoria = "":            LblCATEGORIA.Caption = ""
            Exit Sub
         End If
      End If
      Call cmdcargTRA_Click
   End If
End Sub
Private Sub cmdcargTRA_Click()
   TxHoras = "": TxD = "": TxN = "": TxM = "": Txt = "":    TxMT = "":    TxNa = "":    TxPa = ""
   TxTurno = "": TxTurno.Locked = False
   CboTurDir.Enabled = True
   TxHorario = "": TxHorario.Locked = False
   CboHorario.Enabled = True
   If Txtarjeta = "" Then Exit Sub
   For i = dfi To dff
       z = i - 1
       Check1(z).Value = 0: Check1(z).Enabled = True
       Textd(z).Visible = True: Textp(z).Visible = True
       Textd(z).Text = "": Textd(z).ForeColor = &H80000012
       Textp(z).Text = "":  Textp(z).ForeColor = &H80000012
       Lblhr(z).Caption = ""
   Next
   fecha1 = Format(primer, "yyyymmdd")
   fecha2 = Format(ultimo, "yyyymmdd")
   '-carga  DIRECTORIO
   sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   cDirectHoy = ""
   If Not rs.EOF Then
      For i = Day(primer) To Day(ultimo)
         z = i - 1
         Textd(z) = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
         
         cDirectHoy = cDirectHoy + IIf(IsNull(rs.Fields(z + 4)), "==", Left(Trim(rs.Fields(z + 4)) + "==", 2))
         '-D N P MTD*V ==N*...
         ' 1234567890123456
         
         If InStr("MT_NA_PA", Textd(z)) > 0 And Len(Trim(Textd(z))) = 2 Then
            Select Case Textd(z)
               Case "MT"
                 TxMT = Val(TxMT) + 1
               Case "NA"
                 TxNa = Val(TxNa) + 1
               Case "PA"
                 TxPa = Val(TxPa) + 1
            End Select
         Else
            Select Case Trim(Textd(z).Text)
              Case "A"
                 TxMT = Val(TxMT) + 1
              Case Else
                 If Left(Textd(z), 1) = "M" Then
                    TxM = Val(TxM) + 1
                 ElseIf Left(Textd(z), 1) = "T" Then
                        Txt = Val(Txt) + 1
                 ElseIf Left(Textd(z), 1) = "D" Then
                        TxD = Val(TxD) + 1
                 ElseIf Left(Textd(z), 1) = "N" Then
                        TxN = Val(TxN) + 1
                 ElseIf Left(Textd(z), 1) = "H" Then
                        TxM = Val(TxM) + 1
                        TxN = Val(TxN) + 1
                 ElseIf Left(Textd(z), 1) = "C" Then
                         Txt = Val(Txt) + 1
                         TxN = Val(TxN) + 1
                 ElseIf Left(Textd(z), 1) = "P" Then
                        TxD = Val(TxD) + 1
                        TxN = Val(TxN) + 1
                 End If
            End Select
         End If
      Next i
   End If
   rs.Close
   
   '-Carga HORARIO
   sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha >='" & fecha1 & "' AND fecha<='" & fecha2 & "' ORDER BY Fecha,tarjeta"
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
            sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & Trim(Textp(z)) & "'"
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
   Cbo_ccosto = "": Tx_ccosto = ""
   Txdscccosto = ""
   DD.MoveFirst
   Do While Not DD.EOF
      DD.Delete
      DD.MoveNext
   Loop
   Call CmdLimpia_Click
End Sub
Private Sub TxTurno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     TxTurno = UCase(TxTurno)
     If Trim(TxTurno) <> "" Then
        For i = 0 To CboTurDir.ListCount - 1
          npos = InStr(CboTurDir.List(i), ":")
          nlen = Len(CboTurDir.List(i))
          dato = Trim(Right(CboTurDir.List(i), nlen - npos))
          If dato = Trim(TxTurno) Then
             CboTurDir.ListIndex = i
             Exit For
          End If
        Next i
        Call CboTurDir_KeyPress(13)
        CboHorario.SetFocus     'TxHorario
     End If
  End If
End Sub
Private Sub CboTurDir_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboTurDir.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboTurDir_Click()
   CboTurDir_KeyPress (13)
End Sub
Private Sub CboTurDir_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If IsEmpty(CboTurDir.Text) Then
     Else
        j = CboTurDir.ListIndex
        If j <> -1 Then
           nposicion = InStr(CboTurDir.List(j), ":")
           nlen = Len(CboTurDir.List(j))
           TxTurno = Right(CboTurDir.List(j), nlen - nposicion)
        End If
        CboHorario.SetFocus     'TxHorario
        Exit Sub
     End If
   End If
End Sub
Private Sub TxHorario_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     TxHorario = UCase(TxHorario)
     If Trim(TxHorario) <> "" Then
       For i = 0 To CboHorario.ListCount - 1
         npos = InStr(CboHorario.List(i), "=>")
         nlen = Len(CboHorario.List(i))
         dato = Trim(Right(CboHorario.List(i), nlen - npos - 1))
         If dato = Trim(TxHorario) Then
            If Trim(Txtipopla) = "ADM" Then
            Else
               If Trim(TxTurno) <> "" And Trim(TxTurno) = "*" Then
               Else
               lerror = True
               If InStr("DMTNHCP", Left(TxTurno, 1)) > 0 And Trim(TxTurno) <> "" And Trim(TxTurno) <> "*" Then
               '-If InStr("DMTNHCP_D1D2D3D4M1M2M3T1T2_D*N*N2N3N4H*H1C*C1P*P1_MT_NA_PA", Trim(TxTurno)) > 0 And Trim(TxTurno) <> "" And Trim(TxTurno) <> "*" Then
                  sql = "SELECT * FROM CpCCostoHor WHERE ccosto='" & Txccosto & "' and turno='" & TxTurno & "' and Horario='" & TxHorario & "'"
                  Set rs = New Recordset
                  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  If Not rs.EOF Then
                     '-Horario Ok
                     lerror = False
                  End If
                  rs.Close
               End If
               If lerror Then
                  MsgBox "Selección NO aplica..", vbCritical, "[TURNO-HORARIO]"
                  TxHorario = ""
                  CboHorario.Text = ""
                  Exit Sub
               End If
               End If
            End If
            CboHorario.ListIndex = i
            Exit For
         End If
       Next i
       Call Cbohorario_KeyPress(13)
       CmdAplica.SetFocus
     End If
  End If
End Sub

Private Sub Cbohorario_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboHorario.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub Cbohorario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboHorario) Then
      Else
         j = CboHorario.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboHorario.List(j), "=>")
            nlen = Len(CboHorario.List(j))
            TxHorario = Right(CboHorario.List(j), nlen - nposicion - 1)
            'Valida uso Turno-Horario
            If Txtipopla = "ADM" Then
               MsgBox "ADM"
            Else
              lerror = True
              If Trim(TxTurno) <> "" And InStr("*_V_L_R", Trim(TxTurno)) > 0 Then
                  If TxHorario = "O" Then
                     lerror = False
                  End If
              Else
               
               If InStr("DMTNHCP", Left(TxTurno, 1)) > 0 And Trim(TxTurno) <> "" And Trim(TxTurno) <> "*" Then
               '-If InStr("DMTNHCP_D1D2D3D4M1M2M3T1T2_D*N*N2N3N4H*H1C*C1P*P1_MT_NA_PA", Trim(TxTurno)) > 0 And Trim(TxTurno) <> "" And Trim(TxTurno) <> "*" Then
                  sql = "SELECT * FROM CpCCostoHor WHERE ccosto='" & Txccosto & "' and turno='" + TxTurno + "' and Horario='" & TxHorario & "'"
                  Set rs = New Recordset
                  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  If Not rs.EOF Then
                     lerror = False
                  End If
                  rs.Close
               End If
               If lerror Then
                  MsgBox "Selección, No aplica al Servidor..", vbCritical, "[TURNO-HORARIO]"
                  TxHorario.Text = ""
                  CboHorario.Text = ""
                  Exit Sub
               End If
              End If
            End If
            '-Carga Hras de HORARIO
            sql = "SELECT * FROM CpHorarioNew WHERE codigo='" & TxHorario & "'"
            Set rs = New Recordset
            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs.EOF Then
               cvalor = IIf(IsNull(rs!Horas), "0000", rs!Horas)
               TxnHra = Left(cvalor, 2) + ":" + Right(cvalor, 2)
            End If
            rs.Close
         End If
      End If
      CmdAplica.SetFocus
      Exit Sub
   End If
End Sub
Private Sub Cbohorario_LostFocus()
   Call Cbohorario_KeyPress(13)
End Sub
Private Sub CmdAplica_Click()

   If IsEmpty(CboTurDir) Then
   Else
      
     ' j = CboTurDir.ListIndex
     ' If j <> -1 Then
     '    nposicion = InStr(CboTurDir.List(j), ":")
     '    nlen = Len(CboTurDir.List(j))
     '    Txturno = Right(CboTurDir.List(j), nlen - nposicion)
     ' End If
   End If
   
   '''--    MsgBox "turno:" + Txturno + ")"
   
   If IsEmpty(CboHorario.Text) Then
   Else
      j = CboHorario.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboHorario.List(j), "=>")
         nlen = Len(CboHorario.List(j))
         TxHorario = Right(CboHorario.List(j), nlen - nposicion - 1)
         '-Carga Hras de HORARIO
         sql = "SELECT * FROM CpHorarioNew WHERE codigo='" & TxHorario & "'"
         Set rs = New Recordset
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then
            cvalor = IIf(IsNull(rs!Horas), "0000", rs!Horas)
            TxnHra = Left(cvalor, 2) + ":" + Right(cvalor, 2)
         End If
         rs.Close
      End If
   End If
    '''--      MsgBox "horario:" + TxHorario + ")"
   
   If Trim(TxTurno) = "" Or Trim(TxHorario) = "" Then
   Else
      dfi = Day(primer)
      dff = Day(ultimo)
      
      TxHoras = "": TxD = "": TxN = "": TxM = "": Txt = "": TxMT = "": TxNa = "": TxPa = ""
      For i = dfi To dff
          z = i - 1
          If Check1(z).Value = 1 Then
             'Actualiza Horas Prog
             Textd(z).Text = Trim(TxTurno)
             Textd(z).ForeColor = &HFF0000
             Textp(z).Text = Trim(TxHorario)
             Textp(z).ForeColor = &HFF0000
             'cvalor = Str(Val(Left(Lblhr(z).Caption, 2)) + Val(Right(Lblhr(z).Caption, 2)) / 60)
             'cvalor1 = Str(Val(Left(TxnHra.Text, 2)) + Val(Right(TxnHra.Text, 2)) / 60)
             'TxHoras = Val(TxHoras) + (Val(cvalor1) - Val(cvalor))
             Lblhr(z).Caption = TxnHra
             Check1(z).Value = 0
             '-DataGrid1.SetFocus
             '-Exit Sub
          End If
          
          If Textd(z).Text <> "" Then
             '-RECALCULA
             Select Case Trim(Textd(z))
               Case "MT"
                 TxMT = Val(TxMT) + 1
               Case "NA"
                 TxNa = Val(TxNa) + 1
               Case "PA"
                 TxPa = Val(TxPa) + 1
               Case Else
                If Trim(Textd(z)) = "A" Then
                   'Administ -> Asisten
                   TxMT = Val(TxMT) + 1
                ElseIf Left(Textd(z), 1) = "M" Then
                       TxM = Val(TxM) + 1
                ElseIf Left(Textd(z), 1) = "T" Then
                       Txt = Val(Txt) + 1
                ElseIf Left(Textd(z), 1) = "D" Then
                       TxD = Val(TxD) + 1
                ElseIf Left(Textd(z), 1) = "N" Then
                      TxN = Val(TxN) + 1
                ElseIf Left(Textd(z), 1) = "H" Then
                       TxM = Val(TxM) + 1
                       TxN = Val(TxN) + 1
                ElseIf Left(Textd(z), 1) = "C" Then
                       Txt = Val(Txt) + 1
                       TxN = Val(TxN) + 1
                ElseIf Left(Textd(z), 1) = "P" Then
                       TxD = Val(TxD) + 1
                       TxN = Val(TxN) + 1
                End If
             End Select
             TxHoras = Val(TxHoras) + Val(Left(Trim(Lblhr(z).Caption), 2)) + (Val(Right(Trim(Lblhr(z).Caption), 2))) / 60
          End If
          
      Next
   End If
   
   TxTurno = "": CboTurDir.Text = ""
   TxHorario = "": CboHorario.Text = "": TxnHra = ""
End Sub

Private Sub Limpiacampos()    'OUT
   For i = 1 To dff
     z = i - 1
     Check1(z).Value = 0
     Textd(z).Text = ""
     Textd(z).ForeColor = &H0&
     Textd(z).Text = ""
     Textp(z).ForeColor = &H0&
     Lblhr(z).Caption = ""
   Next
   Txtarjeta = "":  Txtrabajador = "": Txtarjeta.BackColor = &H80000005
   TxTurno = "": CboTurDir = "": CboTurDir.Enabled = False
   'DataGrid1.SetFocus
End Sub

Private Sub cmdGraba_Click()
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If IsNull(Txmes) Then
      MsgBox "Defina PERIODO...  ", vbExclamation
      Exit Sub
   End If
   If fecCierre <= Date Then
      If TxPlH = "3RO" And Txtipopla = "ASI" Then
         '-31.03.2025 TZJ autoriza liberar programa 3RO/ASI
      Else
         MsgBox "Programa CERRADO...  ", vbExclamation
         Exit Sub
      End If
   End If
   If IsEmpty(Txtarjeta) Or IsEmpty(Txtrabajador) Then
      MsgBox "Defina trabajador...  ", vbCritical
      Exit Sub
   End If
   
   ''On Error GoTo final
   ''Cnx.BeginTrans
   
   'Actualiza DIRECTORIO "TURNO"
   Sw = False
   If Txtarjeta <> "" Then
       
      cTurSINHor = ""
      For i = dfi To dff
        z = i - 1
        If (InStr("*_L_V_R", Trim(Textd(z).Text)) = 0 Or Textd(z).Text <> "") And Trim(Textp(z).Text) = "" Then
           If Trim(Textp(z).Text) = "" Then
              cTurSINHor = cTurSINHor + (Right("0" + Trim(Str(i)), 2) + "_") 'Dias turno SIN Horario
           End If
        End If
      Next i
          
      If cTurSINHor <> "" Then
         Beep 500, 350
         MsgBox "Verifique dias " + cTurSINHor, vbCritical, "[DIAS: Turnos Sin HORARIO]"
         Exit Sub
      End If
   
      sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      If rs_d.EOF Then
         Sw = True
         rs_d.AddNew
         rs_d!tarjeta = Txtarjeta
         rs_d!ano = genaño
         rs_d!mes = Txmes
         rs_d!usuario = wNick              'wUsuario    Firma Usuario
         rs_d!fecha = Date
         rs_d!hora = Format(Time, "hh:mm:ss")
         rs_d!estacion = w_clienName
      Else
         rs_d!usuarioM = wNick      'Firma Usuario
         rs_d!fechaM = Date
      End If
      '-Imagen directorio mensual,generado hasta primer dia < Hoy
      cPlantilla0 = ""
      For i = Day(primer) To Day(ultimo)
          z = i - 1
          rs_d.Fields(z + 4) = IIf(IsNull(Textd(z)), "", Textd(z))
          If primer > Date Then
             '-Actualiza Plantilla0
             cPlantilla0 = cPlantilla0 + IIf(IsNull(Textd(z)), "==", Left(Trim(Textd(z)) + "==", 2))
             '-D=N=D=MTD*V===N*...
             ' 1234567890123456
          Else
             '-verifica en cDirectHoy cambio a Textd(z) --> AudDirect
             cvalor = Trim(Mid(cDirectHoy, (i - 1) * 2 + 1, 2))
             c_a = IIf(IsNull(Textd(z)), "==", Left(Trim(Textd(z)) + "==", 2))
             If cvalor <> c_a Then
                '--  MsgBox Trim(Str(i)) + ":de[" + cvalor + "]a[" + c_a + "]"
                   
                SqC = "INSERT INTO CPAudDirect(ano,mes,fecha,ccosto,codupr,tarjeta,deTurno,aTurno,Detalle,registro,estacion,fecregis,horregis) "
                SqC = SqC + " values('" + genaño + "','" + Txmes + "','" + Format(DateValue(primer + z), "yyyymmdd") + "','" + Txccosto + "','" + TxcodUPR + "','" + Txtarjeta + "','" & _
                cvalor + "','" + c_a + "','Directorio Especiales','" + wNick + "','" + w_clienName + "','" + Format(Date, "yyyymmdd") + "','" + Format(Time, "hh:mm:ss") + "')"
                Cnx.Execute SqC
             End If
          End If
      Next i
      rs_d!carEst = IIf(IsNull(TxCarEst), "", TxCarEst)
      rs_d!CCosto = IIf(IsNull(Txccosto), "", Txccosto)
      If Txtipopla = "ASI" Then
         rs_d!CodUpr = IIf(IsNull(TxcodUPR), "", TxcodUPR)
      End If
      rs_d!Cargo = IIf(IsNull(TxCargo), "", TxCargo)
      rs_d!d = IIf(IsNull(TxD), "", TxD)
      rs_d!n = IIf(IsNull(TxN), "", TxN)
      rs_d!m = IIf(IsNull(TxM), "", TxM)
      rs_d!T = IIf(IsNull(Txt), "", Txt)
        rs_d!mT = IIf(IsNull(TxMT), "", TxMT)
        rs_d!na = IIf(IsNull(TxNa), "", TxNa)
        rs_d!pa = IIf(IsNull(TxPa), "", TxPa)
      rs_d!Horas = IIf(IsNull(TxHoras), 0, Val(TxHoras))
      rs_d!Flag = IIf(Txtipopla = "ASI", "0", IIf(Txtipopla = "ADM", "1", " "))  'Asistencial/Administrativo
      
      If DateValue(primer) > Date Then
         '-Graba imagen hasta primer dia < Hoy
         rs_d!plantilla0 = cPlantilla0
      End If
      
      If Txtipopla = "ASI" Then
         '-SABER si una columna existe en una tabla
         SQ1 = "SELECT * FROM  INFORMATION_SCHEMA.COLUMNS AS c1  WHERE  c1.column_name = 'cadhep'  AND c1.table_name = 'cpdirectorio'"
         Set rs_2 = New Recordset
         rs_2.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If Not rs_2.EOF Then
            rs_d!CadHep = IIf(IsNull(TxCadHep), "", Trim(TxCadHep))
         End If
         rs_2.Close
      End If
      rs_d.Update
      rs_d.Close
      If Sw Then
         DD.AddNew
      End If
      DD!tarjeta = Txtarjeta
      DD!nombre = Left(Txtrabajador, 25)
      DD!Cargo = Left(TxDscCargo, 15)
      DD!d = Val(TxD)
      DD!m = Val(TxM)
      DD!T = Val(Txt)
      DD!n = Val(TxN)
      DD!Horas = Val(TxHoras)
      For i = dfi To dff
         z = i - 1
         DD.Fields(z + 8) = IIf(IsNull(Textd(z)), "", Textd(z))
      Next i
      DD!mT = Val(TxMT)
      DD!na = Val(TxNa)
      DD!pa = Val(TxPa)
      DD.Update
   End If
   '-Actualiza Programa de HORARIOS del grupo
   For i = dfi To dff
      z = i - 1
      f_echa = Format(primer + z, "yyyymmdd")
      
      If InStr("*_L_V_R", Trim(Textd(z).Text)) = 0 Then
         If Trim(Textp(z).Text) = 0 Then
            cTurSINHor = cTurSINHor + Right("0" + Trim(Str(i)), 2)   'Dias turno SIN Horario
         End If
      End If
      
      If Textp(z).Text <> "" Then
         Set rs = New Recordset
         sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & f_echa & "'"
         rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         c_a = "XXXX"
         If rs.EOF Then
            rs.AddNew
            rs!tarjeta = IIf(IsNull(Txtarjeta), "", Txtarjeta)
            rs!fecha = primer + z
         Else
            c_a = rs!Horario
         End If
         rs!Horario = Textp(z).Text
         If c_a <> Textp(z).Text Then
            rs!usuario = wNick   'wUsuario      Firma Usuario
            rs!fecRegistro = Date
            rs!HorRegistro = Format(Time, "hh:mm:ss")
            rs!estacion = w_clienName
            If c_a <> "XXXX" Then
               rs!cambio = "X"
            End If
         End If
         rs.Update
         rs.Close
      Else
         If Txtarjeta <> "" Then
            Cnx.Execute "DELETE FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "'AND fecha ='" & f_echa & "'"
         End If
      End If
   Next
   
   ''Cnx.CommitTrans
   DataGrid1.Refresh

   
  '-- If fecCierre < Date Or (TxPlH = "3RO" And Txtipopla = "ASI") Then
      '-TZJ 31.03.2025   autoriza liberar Programa 3RO Asistencial
      Call ProcProg3RO(Txtarjeta, genaño + Txmes)
  '-- End If
  
   
   Call CmdLimpia_Click   'Limpiacampos
   GoTo Final1
''final:
  '' MsgBox Err.Description & " Error: " & Err.Number & vbCr
  '' Cnx.RollbackTrans
Final1:
   Sw = False
End Sub
Private Sub CmdElimina_Click()
    If wrights = "CON" Then         'Or wEditor = "X"
       MsgBox "No Autorizado...  ", vbExclamation
       Exit Sub
    End If
    
    If IsNull(Txmes) Then
      MsgBox "Defina PERIODO...  ", vbExclamation
      Exit Sub
    End If
    If fecCierre <= Date Then
       MsgBox "Programa CERRADO...  ", vbExclamation
       Exit Sub
    End If
    
    If (Txmes <> "" And IsNull(Txtarjeta)) Then Exit Sub
      '-Valida retiro
      Set rt = New Recordset
      sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
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
    Set rs = New Recordset
    sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
       If wUsuario <> rs!usuario Then
          MsgBox "Error, No es Titular Autorizado..! ", vbExclamation
       End If
       rs.Close
       Exit Sub
    End If
    
    If IsNull(Txmes) Or IsNull(Txtarjeta) Then Exit Sub
    If MsgBox("Eliminar Directorio y Programa del mes ?      ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       fecha1 = Format(primer, "yyyymmdd")
       fecha2 = Format(ultimo, "yyyymmdd")
       '-Retira HORARIOS y DIRECTORIO
       Cnx.Execute "DELETE FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "'AND fecha >='" & fecha1 & "' AND fecha<='" & fecha2 & "'"
       Cnx.Execute "DELETE FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
       DD.Delete
    End If
    Call CmdLimpia_Click  'Limpiacampos
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

Private Sub CmdPrint_Click()
   If Tx_ccosto = "" Then
      MsgBox "Defina..  C. Costo ", vbExclamation
      Exit Sub
   End If
      CargaDirecHras
      nTipRpt = 11.3
      InfoCpRS.Left = 0: InfoCpRS.Top = 0
      InfoCpRS.Refresh
      InfoCpRS.Show
      Call CmdLimpia_Click   'Limpiacampos
   CmdPrint.Caption = ""
   Exit Sub
End Sub
Private Sub CmdAudDirect_Click()
   If Txmes <> "" And Tx_ccosto <> "" Then
      '-Lista Cambios del MES en UPR
      nTipRpt = 1011.3
      InfoCpRS.Left = 0: InfoCpRS.Top = 0
      InfoCpRS.Refresh
      InfoCpRS.Show
      Call Limpiacampos
      CmdAudDirect.Caption = ""
   End If
   Exit Sub
End Sub

Private Sub CargaDirecHras()
    If Trim(Tx_ccosto) = "" Then
        MsgBox "C. Costo.. ", vbExclamation
       Exit Sub
    End If
    On Error Resume Next
    Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmpDCosto"
    '-USO TEMPORAL
     Cnx.Execute ("SELECT * Into #Tabla  FROM CpTmpDCosto WHERE 1 = 2")
     Cnx.Execute "DELETE  FROM #Tabla"
     Cnx.CommitTrans
    Set r_d = New ADODB.Recordset
    nposicion = 0
    '-Directorio Mensual de CC
    sql = "SELECT  * FROM CpDirectorio WHERE ano='" & genaño & "' and mes='" & Txmes & "' and ccosto='" & Tx_ccosto & "' " & _
    "ORDER BY ccosto,Tarjeta "
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Set rs_d = New Recordset
    sq_l = "SELECT * FROM #Tabla"           'CpTmpDcosto
    rs_d.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    unafecha = "01/" + Trim(r_d!mes) + "/" + Trim(Str(genaño))
    Call ObtenerPrimerUltimoDia(unafecha)
    unafecha = primer
    Do While Not r_d.EOF
       CmdPrint.Caption = r_d!tarjeta
       Set rs = New Recordset
       rs.Open "SELECT * FROM cpgeneral WHERE  (baja is NULL) and tarjeta='" & r_d!tarjeta & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
       cNombre = ""
       ndia = 0
       If rs.EOF Then
          MsgBox "Servidor de baja " + r_d!tarjeta, vbExclamation, "[IMPORTANTE]"
          rs.Close
       Else     '<=======
       
          cNombre = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 30)
          If IsDate(rs!fecnac) Then
             If Month(rs!fecnac) = Val(Txmes) Then
                ndia = Day(rs!fecnac)
             End If
          End If
          rs.Close
       rs_d.AddNew
       rs_d!tarjeta = r_d!tarjeta
       rs_d!ano = r_d!ano
       rs_d!mes = r_d!mes
       cCarest = IIf(IsEmpty(r_d!carEst), "", r_d!carEst)
       ccCosto = IIf(IsEmpty(r_d!CCosto), "", r_d!CCosto)
       cCodUpr = IIf(IsEmpty(r_d!CodUpr), "", r_d!CodUpr)
       cCargo = IIf(IsEmpty(r_d!Cargo), "", r_d!Cargo)
       rs_d!Cargo = cCargo
       rs_d!CCosto = ccCosto
       rs_d!CodUpr = cCodUpr
       
       '-Carga Plantilla 0 del servidor
       cPlantilla0 = IIf(IsNull(r_d!plantilla0), "", Trim(r_d!plantilla0))
       For i = Day(primer) To Day(ultimo)
          z = i - 1
          If ndia = i Then   'Onomastico
             rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "º", "º" + r_d.Fields(z + 4))
          Else
             '-valor turno Actual
             c_a = IIf(IsNull(r_d.Fields(z + 4)), "==", r_d.Fields(z + 4))
             c_a = Left(Trim(c_a) + "==", 2)
          
             FechaB = Format(primer + z, "yyyymmdd")
             Set rs_d1 = New Recordset
             Sq_0 = "SELECT * FROM CpchgTurM" & genaño & " WHERE tarjeta='" & r_d!tarjeta & "' and Fecha='" & FechaB & "'"      'muestra TODO CAMBIO and CamHOR<>'O'"
             rs_d1.Open Sq_0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If rs_d1.EOF Then
                If cPlantilla0 <> "" Then
                   '-verifica cambio en plantilla0
                   cvalor = Trim(Mid(cPlantilla0, (i - 1) * 2 + 1, 2))
                   If cvalor <> c_a Then
                      '-Cambio TurnoDia en Ditectorio
                      rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "¹", "¹" + r_d.Fields(z + 4))
                   Else
                      '-No cambios
                      rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
                   End If
                Else
                   rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
                End If
             Else
                '-realizo Cambio x Papeleta
                rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "²", "²" + r_d.Fields(z + 4))
             End If
             rs_d1.Close
          End If
          rs_d.Fields(z + 6 + 31) = Left(WeekdayName(Weekday(primer + z)), 1)   'LU/MA/MI/JU/VI/SA/DO
       Next i
       rs_d!nombre = cNombre
       If cCargo <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM cpcargos WHERE codigo='" & cCargo & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
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
       rs_d!mT = IIf(IsEmpty(r_d!mT), "", r_d!mT)     '<========
       rs_d!na = IIf(IsEmpty(r_d!na), "", r_d!na)     '<========
       rs_d!pa = IIf(IsEmpty(r_d!pa), "", r_d!pa)     '<========
       rs_d!Horas = IIf(IsEmpty(r_d!Horas), 0, r_d!Horas)
       rs_d!Dia = ndia
       rs_d.Update
       
       End If    '======>
       r_d.MoveNext
    Loop
    'Cnx.CommitTrans
    rs_d.Close
    r_d.Close
    Exit Sub
End Sub
Private Sub creadd()
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
     .Fields.Append "MT", adChar, 2     '<========
     .Fields.Append "NA", adChar, 2     '<========
     .Fields.Append "PA", adChar, 2     '<========
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
            '.NumberFormat = "##.#0"
            .Width = Me.TextWidth(String(6, "S"))
          Case Else
            .Width = Me.TextWidth(String(3, "S"))
        End Select
        '.NumberFormat = "##,##0.0"
        .Locked = True
     End With
  Next
End Sub

Private Sub cmdsalir_Click()
   Set r_d = Nothing
   Set rs_d = Nothing
   Set rs = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If UCase(Forms(cont).Name) <> "CPMENU" And UCase(Forms(cont).Name) <> "PROGESPECIAL" Then
           Unload Forms(cont)
        End If
      End If
   Next
   
   Me.Hide
   Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set r_d = Nothing
   Set rs_d = Nothing
   Set rs = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "ProgEspecial" Then
           Unload Forms(cont)
        End If
      End If
   Next
   
   Me.Hide
   Unload Me
End Sub

