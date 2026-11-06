VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RecMarcasRS 
   Caption         =   "Record Trabajador"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15075
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "RecMarcasRS.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   14160
      Picture         =   "RecMarcasRS.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   216
      ToolTipText     =   "Cerrar"
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton CmdPerTarjeta 
      Caption         =   "Periodo del trabajador"
      Height          =   495
      Left            =   13080
      TabIndex        =   215
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox TxVerAcum 
      Height          =   285
      Left            =   12120
      MaxLength       =   1
      TabIndex        =   214
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox TxPLH 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   212
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxVerDia 
      Height          =   285
      Left            =   10680
      MaxLength       =   2
      TabIndex        =   211
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxIniLac 
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
      Left            =   9840
      TabIndex        =   209
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TxFinLac 
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
      Left            =   10800
      TabIndex        =   208
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TxHorasFMP 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      TabIndex        =   206
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox TxmTarda 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5520
      TabIndex        =   202
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox Txultdia 
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
      Left            =   1200
      TabIndex        =   201
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Txpridia 
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
      Left            =   1200
      TabIndex        =   200
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxCodUpr 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8520
      TabIndex        =   199
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxDni 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   198
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Txactividad 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      MaxLength       =   4
      TabIndex        =   197
      Top             =   240
      Visible         =   0   'False
      Width           =   615
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
      Left            =   10080
      TabIndex        =   196
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxDscCarEst 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   194
      Top             =   1800
      Width           =   4935
   End
   Begin VB.TextBox TxDscCcosto 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   192
      Top             =   1440
      Width           =   4935
   End
   Begin VB.TextBox TxDscActividad 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   190
      Top             =   1080
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "RecMarcasRS.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   188
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TxHorasFI 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   187
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox TxHorasCom 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   185
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox TxCarEst 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   183
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxCCosto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8160
      MaxLength       =   2
      TabIndex        =   182
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Txcargo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      MaxLength       =   4
      TabIndex        =   181
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxHorasMP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   180
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox TxHorasFJ 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   114
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox TxmDscto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   112
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox TxHorasT 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   110
      Top             =   7800
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   9840
      TabIndex        =   106
      Top             =   6600
      Width           =   2295
      Begin VB.CommandButton mnuFicImpSelec 
         Height          =   675
         Left            =   120
         Picture         =   "RecMarcasRS.frx":80FF2
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Selec.Impresora"
         Top             =   240
         Width           =   765
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
         Height          =   675
         Left            =   1320
         Picture         =   "RecMarcasRS.frx":817D0
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Imprimir"
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox TxHorasP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   13320
      Picture         =   "RecMarcasRS.frx":8209A
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Otra Consulta"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox TxGrpGen 
      Height          =   285
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   14775
      Begin VB.TextBox Txfecha 
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
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   178
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   30
         Left            =   10800
         TabIndex        =   177
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   29
         Left            =   9480
         TabIndex        =   176
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   28
         Left            =   8160
         TabIndex        =   175
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   27
         Left            =   6720
         TabIndex        =   174
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   26
         Left            =   5400
         TabIndex        =   173
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   25
         Left            =   4080
         TabIndex        =   172
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   24
         Left            =   2760
         TabIndex        =   171
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   23
         Left            =   1440
         TabIndex        =   170
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   22
         Left            =   120
         TabIndex        =   169
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   21
         Left            =   13440
         TabIndex        =   168
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   20
         Left            =   12120
         TabIndex        =   167
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   19
         Left            =   10800
         TabIndex        =   166
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   18
         Left            =   9480
         TabIndex        =   165
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   17
         Left            =   8160
         TabIndex        =   164
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   16
         Left            =   6840
         TabIndex        =   163
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   15
         Left            =   5520
         TabIndex        =   162
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   14
         Left            =   4200
         TabIndex        =   161
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   13
         Left            =   2880
         TabIndex        =   160
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   12
         Left            =   1560
         TabIndex        =   159
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   158
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   10
         Left            =   13440
         TabIndex        =   157
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   9
         Left            =   12120
         TabIndex        =   156
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   8
         Left            =   10800
         TabIndex        =   155
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   7
         Left            =   9480
         TabIndex        =   154
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   6
         Left            =   8160
         TabIndex        =   153
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   5
         Left            =   6840
         TabIndex        =   152
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   4
         Left            =   5520
         TabIndex        =   151
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   3
         Left            =   4200
         TabIndex        =   150
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   2880
         TabIndex        =   149
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   148
         Top             =   240
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   147
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Cmdetalle 
         Caption         =   "Detalle Dia"
         Height          =   375
         Left            =   13080
         TabIndex        =   146
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   11520
         TabIndex        =   105
         Text            =   "Txt1"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   10200
         TabIndex        =   104
         Text            =   "Txt1"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   8880
         TabIndex        =   103
         Text            =   "Txt1"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   7440
         TabIndex        =   102
         Text            =   "Txt1"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   6120
         TabIndex        =   101
         Text            =   "Txt1"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   4800
         TabIndex        =   100
         Text            =   "Txt1"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   3480
         TabIndex        =   99
         Text            =   "Txt1"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   2160
         TabIndex        =   98
         Text            =   "Txt1"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   840
         TabIndex        =   97
         Text            =   "Txt1"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   14160
         TabIndex        =   96
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   12840
         TabIndex        =   95
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   11520
         TabIndex        =   94
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   10200
         TabIndex        =   93
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   8880
         TabIndex        =   92
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   7560
         TabIndex        =   91
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   6240
         TabIndex        =   90
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   4920
         TabIndex        =   89
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   3600
         TabIndex        =   88
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   2280
         TabIndex        =   87
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   840
         TabIndex        =   86
         Text            =   "Txt1"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   14160
         TabIndex        =   85
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   12840
         TabIndex        =   84
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   8
         Left            =   11520
         TabIndex        =   83
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   10200
         TabIndex        =   82
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   8880
         TabIndex        =   81
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7560
         TabIndex        =   80
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   6240
         TabIndex        =   79
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4920
         TabIndex        =   78
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3600
         TabIndex        =   77
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   76
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   75
         Text            =   "Txt1"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   11160
         MaxLength       =   4
         TabIndex        =   37
         Text            =   "Text31"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   29
         Left            =   9840
         MaxLength       =   4
         TabIndex        =   36
         Text            =   "Text30"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   8520
         MaxLength       =   4
         TabIndex        =   35
         Text            =   "Text29"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   7080
         MaxLength       =   4
         TabIndex        =   34
         Text            =   "Text28"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   33
         Text            =   "Text27"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   4440
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "Text26"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "Text25"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   30
         Text            =   "Text24"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   480
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "Text23"
         Top             =   2520
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   13800
         MaxLength       =   4
         TabIndex        =   28
         Text            =   "Text22"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   12480
         MaxLength       =   4
         TabIndex        =   27
         Text            =   "Text21"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   11160
         MaxLength       =   4
         TabIndex        =   26
         Text            =   "Text20"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   9840
         MaxLength       =   4
         TabIndex        =   25
         Text            =   "Text19"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   8520
         MaxLength       =   4
         TabIndex        =   24
         Text            =   "Text18"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   23
         Text            =   "Text17"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   22
         Text            =   "Text16"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   4560
         MaxLength       =   4
         TabIndex        =   21
         Text            =   "Text15"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   20
         Text            =   "Text14"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   19
         Text            =   "Text13"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   480
         MaxLength       =   4
         TabIndex        =   18
         Text            =   "Text12"
         Top             =   1560
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   13800
         MaxLength       =   4
         TabIndex        =   17
         Text            =   "Text11"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   12480
         MaxLength       =   4
         TabIndex        =   16
         Text            =   "Text10"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   11160
         MaxLength       =   4
         TabIndex        =   15
         Text            =   "Text9"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   9840
         MaxLength       =   4
         TabIndex        =   14
         Text            =   "Text8"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   8520
         MaxLength       =   4
         TabIndex        =   13
         Text            =   "Text7"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "Text6"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5880
         MaxLength       =   4
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4560
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   600
         Width           =   395
      End
      Begin VB.TextBox Textd 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   600
         MaxLength       =   4
         TabIndex        =   7
         Text            =   "Txd1"
         Top             =   600
         Width           =   395
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
         Left            =   8760
         TabIndex        =   195
         Top             =   3360
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   30
         Left            =   10800
         TabIndex        =   145
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   29
         Left            =   9480
         TabIndex        =   144
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   8160
         TabIndex        =   143
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   6720
         TabIndex        =   142
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   5400
         TabIndex        =   141
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   4080
         TabIndex        =   140
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   2760
         TabIndex        =   139
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   1440
         TabIndex        =   138
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   120
         TabIndex        =   137
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   13440
         TabIndex        =   136
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   12120
         TabIndex        =   135
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   10800
         TabIndex        =   134
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   9480
         TabIndex        =   133
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   8160
         TabIndex        =   132
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   6840
         TabIndex        =   131
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   5520
         TabIndex        =   130
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   4200
         TabIndex        =   129
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   2880
         TabIndex        =   128
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   1560
         TabIndex        =   127
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   126
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   13440
         TabIndex        =   125
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   12120
         TabIndex        =   124
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   10800
         TabIndex        =   123
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   9480
         TabIndex        =   122
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   8160
         TabIndex        =   121
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   6840
         TabIndex        =   120
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   119
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4200
         TabIndex        =   118
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   117
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   116
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbltur 
         Caption         =   "lbltur"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   115
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   30
         Left            =   11040
         TabIndex        =   74
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   29
         Left            =   9720
         TabIndex        =   73
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   28
         Left            =   8400
         TabIndex        =   72
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   27
         Left            =   6960
         TabIndex        =   71
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   26
         Left            =   5640
         TabIndex        =   70
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   25
         Left            =   4320
         TabIndex        =   69
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   24
         Left            =   3000
         TabIndex        =   68
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   23
         Left            =   1680
         TabIndex        =   67
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   22
         Left            =   360
         TabIndex        =   66
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   11
         Left            =   360
         TabIndex        =   65
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   12
         Left            =   1800
         TabIndex        =   64
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   13
         Left            =   3120
         TabIndex        =   63
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   14
         Left            =   4440
         TabIndex        =   62
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   15
         Left            =   5760
         TabIndex        =   61
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   16
         Left            =   7080
         TabIndex        =   60
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   17
         Left            =   8400
         TabIndex        =   59
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   18
         Left            =   9720
         TabIndex        =   58
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   19
         Left            =   11040
         TabIndex        =   57
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   20
         Left            =   12360
         TabIndex        =   56
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   21
         Left            =   13680
         TabIndex        =   55
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   10
         Left            =   13680
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   9
         Left            =   12360
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   8
         Left            =   11040
         TabIndex        =   52
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   7
         Left            =   9720
         TabIndex        =   51
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   6
         Left            =   8400
         TabIndex        =   50
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   5
         Left            =   7080
         TabIndex        =   49
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   48
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   47
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   2
         Left            =   3120
         TabIndex        =   46
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Lbldia 
         Alignment       =   2  'Center
         Caption         =   "Lbldia"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Txtrabajador 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   4695
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2880
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   195
      Left            =   4560
      MaxLength       =   2
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Txtarjeta 
      Height          =   285
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Label LblVerAcum 
      BackStyle       =   0  'Transparent
      Caption         =   "Acumula"
      Height          =   255
      Left            =   11400
      TabIndex        =   213
      Top             =   240
      Width           =   975
   End
   Begin VB.Label LblverDia 
      BackStyle       =   0  'Transparent
      Caption         =   "Visualizar   Dia"
      Height          =   255
      Left            =   9600
      TabIndex        =   210
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Dni"
      Height          =   255
      Left            =   3840
      TabIndex        =   207
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "T/H Falta (MP)"
      Height          =   255
      Left            =   1800
      TabIndex        =   205
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Tardanza"
      Height          =   255
      Left            =   5520
      TabIndex        =   204
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Dscto"
      Height          =   255
      Left            =   6360
      TabIndex        =   203
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      Height          =   255
      Left            =   2040
      TabIndex        =   193
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "U.Organica"
      Height          =   255
      Left            =   2040
      TabIndex        =   191
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Actividad"
      Height          =   255
      Left            =   2040
      TabIndex        =   189
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "T/H. Falta Injust."
      Height          =   255
      Left            =   1800
      TabIndex        =   186
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Mot. Particular  Hras"
      Height          =   255
      Left            =   4680
      TabIndex        =   184
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Comision"
      Height          =   255
      Left            =   7440
      TabIndex        =   179
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "T/H. Falta Justf. "
      Height          =   255
      Left            =   1800
      TabIndex        =   113
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Minutos "
      Height          =   255
      Left            =   4680
      TabIndex        =   111
      Top             =   6600
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Horas Trabajadas"
      Height          =   255
      Left            =   4680
      TabIndex        =   109
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "G. General"
      Height          =   255
      Left            =   8280
      TabIndex        =   43
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Horas Programadas"
      Height          =   255
      Left            =   1800
      TabIndex        =   42
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "RESUMEN"
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label_2 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "RecMarcasRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim flgDir, Sw, F_falta As Boolean
Dim nSonLacta, nLacta, nposicion, dfi, dff, z, nsondias, nTtarda, nTDscto, ndscto, nTotros, nValorH, nValorM As Integer
Dim SQL0, sql, Sq0, cgrupo, c_a, c_turno, cvalor, p_ing, p_ing1, p_sal, p_sal1, p_Sale, p_horas, m_ing, m_ing1, m_sal, m_sal1 As String
Dim cAuPrg, cDiahora, cTipo, choraini, chorafin, cDescuenta, cMantGuardia As String
Dim csexo, cAplica, cDActividad, cDccosto, cDcarest, cDscPLH As String
Dim nHrFalt, nDFJ, nDFI, nDMP, nHrProg, nDescuento, nDesCta, nHrTrab As Double
Dim unafecha As Date
Dim nVerDia As Integer
Private Sub Command1_Click()
     c_grupo = "03": c_secuencia = "032"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub

Private Sub Form_Load()
 TxVerDia = "": TxVerAcum = ""
 flgDir = False
 
  If wUsuario <> "AMASTER" Then
     LblverDia.Visible = False: LblVerAcum.Visible = False
     TxVerDia.Visible = False:      TxVerAcum.Visible = False
  End If
  RecMarcasRS.Caption = "Record Trabajador - " & genaño
  For i = 1 To 31
    z = i - 1
    Option1(z).Value = False
    Option1(z).Enabled = False
    Option1(z).Visible = False
    Textd(z).Text = ""
    Textd(z).ForeColor = &H0&
    Textd(z).Locked = True
    Text1(z).Text = ""
    Text1(z).ForeColor = &H0&
    Text1(z).Locked = True
  Next
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
  Txmes = ""
  '-inicia en mes actual
  CboMes.ListIndex = Month(Date) - 1
  Cbomes_KeyPress (13)
  
End Sub
Private Sub Cbomes_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboMes.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbomes_LostFocus()
   If IsEmpty(CboMes.Text) Then
   Else
      TxVerDia = "": TxVerAcum = ""
      j = CboMes.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboMes.List(j), ":")
         nlen = Len(CboMes.List(j))
         Txmes = Right(CboMes.List(j), nlen - nposicion)
         unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
         Call ObtenerPrimerUltimoDia(unafecha)
         Txpridia = primer
         TxUltDia = ultimo
         If Day(TxUltDia) < 31 Then
            For i = Day(txultimo) + 1 To 31
               z = i - 1
               Lbldia(z).Visible = False
            Next i
         End If
         Call Cargadias       'primer/ultimo
      End If
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
            Txmes = Right(CboMes.List(j), nlen - nposicion)
            unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)
            Txpridia = primer
            TxUltDia = ultimo
            Call Cargadias       'primer/ultimo
         End If
      End If
      Exit Sub
   End If
End Sub

Private Sub Cargadias()
  
      '-Verifica datos de posicion del servidor en la fecha de consulta
      flgDir = False
      Set rs = New Recordset
      rs.Open "SELECT * FROM CpCronograma where ano ='" & genaño & "' ORDER BY mes", Cnx, adOpenStatic, adLockReadOnly
      rs.MoveLast
      If Not rs.EOF Then
         fechaB = rs!fecha
         If Left(Format(fechaB, "YYYYMMDD"), 6) > Left(Format(primer, "YYYYMMDD"), 6) Or Val(Left(Format(primer, "YYYYMMDD"), 4)) <> Year(Date) Then
            '-ccosto,codupr,cargo,Carest,actividad del directorio
            '--MsgBox "uso datos historicos " + Left(Format(fechab, "yyyymmdd"), 6) + "," + Left(Format(primer, "yyyymmdd"), 6)
            flgDir = True
         End If
      End If
      rs.Close
      
   unafecha = primer
   dfi = Day(primer)
   dff = Day(ultimo)
   For i = dfi To dff
       z = i - 1
       Lbldia(z).Visible = True
       Lbldia(z).Caption = Left(WeekdayName(Weekday(primer + z)), 9) + " " & i
       Lbldia(z).ForeColor = &H80000012            'negro
       '''If Weekday(unafecha + i - 1) = 1 Then
       If UCase(Left(WeekdayName(Weekday(primer + z)), 3)) = "DOM" Then
          '- UCase(Left(WeekdayName(Weekday(unafecha + i - 1)), 3)) = "DOM"
          Lbldia(z).ForeColor = &HFF               'rojo
       Else
          Set rs = New Recordset
          rs.Open "SELECT * FROM Feriados ORDER BY codigo", Cnx0, adOpenStatic, adLockReadOnly
          rs.MoveFirst
          While Not rs.EOF
             If Len(Trim(rs!codigo)) = 6 Then
                c_a = Right(rs!codigo, 2)
             ElseIf Len(Trim(rs!codigo)) = 4 Then
                c_a = Right(Year(unafecha), 2)
             End If
             If Val(Left(rs!codigo, 2)) = Month(unafecha) And Val(Mid(rs!codigo, 3, 2)) = i And Right(Year(unafecha), 2) = c_a Then
                Lbldia(z).ForeColor = &HFF
             End If
             rs.MoveNext
          Wend
          rs.Close
       End If
   Next
End Sub
Private Sub CmdClear_Click()
  '-TxPridia = TxPridia
  '-TxUltDia = TxUltDia
  Txtarjeta = "": Txtrabajador = "": Txfecnac = "": TxDni = "": TxGrpGen = ""
  TxPlh = "": Txdscactividad = "": Txdscccosto = "": TxDscCarEst = ""
  TxhorasP = "": TxHorast = "": TxhorasFj = "": TxhorasFI = ""
  TxmTarda = "": TxmDscto = "": Txhorasfmp = "": Txhorasmp = "": TxHorasCom = ""
  LblONOMASTICO.Caption = "": LblONOMASTICO.Visible = False
  For i = 1 To 31
    z = i - 1
    Option1(z).Value = False
    Option1(z).Enabled = False
    Option1(z).Visible = False
    Lbltur(z).Caption = ""
    Textd(z).Text = ""
    Textd(z).ForeColor = &H0&
    Textd(z).Locked = True
    Text1(z).Text = ""
    Text1(z).ForeColor = &H0&
    Text1(z).Locked = True
  Next
  Cmdetalle.Enabled = False
  Txtarjeta.SetFocus
End Sub

Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      LblONOMASTICO.Caption = ""
      LblONOMASTICO.Visible = False
      If Trim(Txtarjeta) = "" Or (Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta)) Then
         If Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta) Then
            Ayuda1.TxUbiAP = Trim(Txtarjeta)
         End If
         Load Ayuda1
         Ayuda1.Tag = 3.1
         Ayuda1.Show 1
      End If
      If Trim(Txtarjeta) <> "" Then
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "Tarjeta, No Ubicada", vbExclamation
            Txtarjeta = ""
            rs.Close
            Exit Sub
         Else
            Txtrabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 45)
            csexo = IIf(IsNull(rs!sexo), "", rs!sexo)
            TxGrpGen = IIf(IsNull(rs!grpgen), "", Trim(rs!grpgen))
            If Not flgDir Then
               TxPlh = IIf(IsNull(rs!PLH), "", Trim(rs!PLH))
               Txactividad = IIf(IsNull(rs!actividad), "", Trim(rs!actividad))
               TxCargo = IIf(IsNull(rs!cargo), "", Trim(rs!cargo))
               TxCarEst = IIf(IsNull(rs!carEst), "", Trim(rs!carEst))
               Txccosto = IIf(IsNull(rs!CCosto), "", Trim(rs!CCosto))      'en Directorio..periodo CERRADO!
               TxCodUpr = IIf(IsNull(rs!CodUpr), "XX", Trim(rs!CodUpr))    'en Directorio ..periodo CERRADO!
            Else
               '-datos en DIRECTORIO  -*
               Set rt = New Recordset
               sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Trim(Txtarjeta) & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
               rt.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not rt.EOF Then
                  TxPlh = IIf(IsNull(rt!PLH), "", Trim(rt!PLH))
                  Txactividad = IIf(IsNull(rt!actividad), "", Trim(rt!actividad))
                  TxCargo = IIf(IsNull(rt!cargo), "", Trim(rt!cargo))
                  TxCarEst = IIf(IsNull(rt!carEst), "", Trim(rt!carEst))
                  Txccosto = IIf(IsNull(rt!CCosto), "", Trim(rt!CCosto))      'en Directorio..periodo CERRADO!
                  TxCodUpr = IIf(IsNull(rt!CodUpr), "XX", Trim(rt!CodUpr))    'en Directorio ..periodo CERRADO!
                End If
                rt.Close
            End If
            
            If IsDate(rs!fecnac) Then Txfecnac = rs!fecnac
            TxDni = IIf(IsNull(rs!Documento), "", Trim(rs!Documento))
         End If
         rs.Close
      End If
     
      If Wccosto <> "" Then
         '-evalua servidor en CCosto
         If TxDni = wUsuario Then
            '''MsgBox "0"
            ceditor = "d"
         Else
            If wDirector = "S" And Wccosto = Txccosto Then
               '-Verifica si es Director Dpto/Ofic
               '''MsgBox "1"
               ceditor = "d"
            Else
               '-Valida si servidor pertenece a UPR
               '''MsgBox "2"
               ceditor = "X"
               If weditor <> "X" And Trim(TxCodUpr) <> "" Then
                  '-verifica editor ejecuta en  UPR
                  npos = InStr(wCadenaUpr, TxCodUpr)   ' 99S_99N_
                  If npos > 0 Then
                     'MID("Ejemplo",3,2) devuelve: "em"
                     ceditor = Right(Mid(wCadenaUpr, npos, 3), 1)
                  End If
                  '''MsgBox "3"
               Else
                  '-verifica si es editor de UPR   MsgBox "3 " + TxCodUpr
                  Set rs = New Recordset
                  sq2 = "SELECT * FROM CpCcostoupr WHERE codupr='" + TxCodUpr + "'"
                  rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  If Not rs.EOF Then
                     '''MsgBox "4"
                     '-Valida si servidor Edita UPR
                     If rs!Dni1 = wUsuario Or rs!Dni2 = wUsuario Or rs!dni3 = wUsuario Then
                        ceditor = "d"
                        '''MsgBox "5"
                     End If
                  End If
                   '''MsgBox "6"
               End If
            End If
            '''MsgBox "7"
        
         End If
      End If
        '''MsgBox "8"
      If ceditor = "X" Then
         Beep 500, 350
         Beep 500, 350
         MsgBox "Tarjeta, No Autorizada ", vbCritical
         Call CmdClear_Click
         Exit Sub
      End If
      
      If IsDate(Txfecnac) Then
         If Month(Txfecnac) = Val(Txmes) Then
            LblONOMASTICO.Visible = True
            LblONOMASTICO.Caption = "Onomastico el " + Str(Day(Txfecnac))
         End If
      End If
      '-Establece Afecto X LACTANCIA -- aplica DIARIO
      '=============================
      cAplica = "N": nSonLacta = 0   '-caso parto multiple max aplica 2Hras
      TxIniLac = "": TxFinLac = ""
      If csexo = "F" Then
         '-Ubica lactancia al periodo
         Set rt = New Recordset
         sql = "SELECT * FROM CpLactancia WHERE DNI ='" & TxDni & "' order by fecnac"
         rt.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         Do While Not rt.EOF
            dftope = DateAdd("YYYY", 1, rt!fecnac)   'Tope Lactancia
            'MsgBox Format(dftope, "dd/mm/yyyy")
            If (rt!fecnac >= DateValue(Txpridia) And DateValue(Txpridia) <= dftope) Or (DateValue(TxUltDia) >= rt!fecnac And DateValue(TxUltDia) <= dftope) Then
               '-Carga Parametros de Ajuste horario por Lactancia
               TxIniLac = rt!fecnac
               TxFinLac = rt!FecFIN
               cAplica = Trim(rt!Aplica)
            End If
            rt.MoveNext
         Loop
         rt.Close
      End If
      If Trim(cAplica) <> "N" Then
         Txtrabajador = Trim(Txtrabajador) + IIf(nSonLacta = 1, " <LAC>", " <LAC2>")
      End If
      Call CargaRegistro
      Exit Sub
   End If
End Sub
Private Sub CargaRegistro()
   If Txmes = "" Then Exit Sub
   If Txtrabajador = "" Then Exit Sub
   TxhorasP = "": TxHorasCom = "": TxhorasFj = "": TxhorasFI = ""
   TxmTarda = "": TxmDscto = "": Txhorasfmp = "": Txhorasmp = "": TxHorast = ""
   cDscPLH = "": cDActividad = "": Txdscactividad = "": cDcarest = "": TxDscCarEst = ""
   cDccosto = "": Txdscccosto = ""
   Set rs = New Recordset
   rs.Open "SELECT * FROM CpTablasAUX WHERE  grupo='REG' and codigo='" & TxPlh & "'", Cnx, adOpenStatic, adLockReadOnly
   If Not rs.EOF Then
      cDscPLH = Trim(rs!descripcion)
   End If
   rs.Close
   
   Set rs = New Recordset
   rs.Open "SELECT * FROM CpActividad WHERE  codigo='" & Txactividad & "'", Cnx, adOpenStatic, adLockReadOnly
   If Not rs.EOF Then
      cDActividad = Trim(rs!descripcion)
      Txdscactividad = Trim(rs!descripcion)
   End If
   rs.Close
   Set rs = New Recordset
   rs.Open "SELECT * FROM Cpcargos WHERE  codigo='" & TxCargo & "'", Cnx, adOpenStatic, adLockReadOnly
   If Not rs.EOF Then
      cDcarest = Trim(rs!descripcion)
      TxDscCarEst = Trim(rs!descripcion)
   End If
   rs.Close
   Set rs = New Recordset
   rs.Open "SELECT * FROM Cpccosto WHERE codigo='" & Txccosto & "'", Cnx, adOpenStatic, adLockReadOnly
   If Not rs.EOF Then
      cDccosto = Trim(rs!descripcion)
      Txdscccosto = Trim(rs!descripcion)
   End If
   rs.Close
    
  '-creo TEMPORAL
  Cnx.BeginTrans
  On Error GoTo borra
  Cnx.Execute ("SELECT * Into #tabla1  FROM CpTmRecord WHERE 1 = 2")
    GoTo sigue
borra:
     Cnx.Execute "DELETE  FROM #tabla1"
sigue:
  Cnx.CommitTrans
  
  sql = "SELECT * FROM #tabla1"
  Set rt = New Recordset
  rt.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   For i = 1 To 31
      z = i - 1
      Lbltur(z).Caption = ""
      Textd(z).Text = ""
      Text1(z).ForeColor = &H80000012      'negro
      Text1(z).Text = ""
   Next
   dfi = Day(primer)
   dff = Day(ultimo)
   '-Carga turnos programados DIRECTORIO
   sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Trim(Txtarjeta) & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      For i = dfi To dff
         z = i - 1
         Option1(z).Value = False
         Option1(z).Visible = True
         Option1(z).Enabled = True
         Lbltur(z).Caption = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
      Next i
   End If
   rs.Close
   nTtarda = 0: nTDscto = 0:  nTotros = 0
   cdiaant = "N"      'marcacion caso Cierra turno NOCHE (2dias  19:30-07:30)
   p_Sale = ""
   
   For i = dfi To dff
       z = i - 1
       fechaB = Format(primer + z, "yyyymmdd")
       dfentra = primer + z
       dfsale = primer + z
       c_turno = "": p0_ing = "": p0_sal = "": p0_ing1 = "": p0_sal1 = "" 'HORARIO ORIGINAL
       p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": c2dias = "N"
       m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
       nHrProg = 0: nValorH = 0: nValorM = 0: nmintarde = 0: ndscto = 0: F_falta = False: nmotpersonal = 0: nDsctoMP = 0: notros = 0: nComision = 0
       cvalor = "": cAuPrg = "": cDiahora = "": cTipo = "": choraini = "": chorafin = "": cDescuenta = "": cMantGuardia = "": c_a = ""
       
       '-Tolerancia x LACTANCIA (I)ngreso/(S)alida -- APLICA si fecha es AFECTA
       '=======================
       nLacta = 0
       cAplica = "N": nSonLacta = 0   '-caso parto multiple max aplica 2Hras
       TxIniLac = "": TxFinLac = ""
       If csexo = "F" Then
          '-Establece Afecto X LACTANCIA dia es AFECTO ¿?
          '=============================
          '-Ubica lactancia del periodo
          Set rs = New Recordset
          sql = "SELECT * FROM CpLactancia WHERE DNI ='" & TxDni & "' order by fecnac"
          rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          Do While Not rs.EOF
             dftope = DateAdd("YYYY", 1, rs!fecnac)   'Tope Lactancia
             'MsgBox Format(dftope, "dd/mm/yyyy")
             If fechaB >= rs!fecnac And fechaB <= Format(dftope, "yyyymmdd") Then
                nSonLacta = nSonLacta + 1
                '-Carga Parametros de Ajuste horario del DIA por Lactancia
                TxIniLac = rs!fecnac
                TxFinLac = rs!FecFIN
                cAplica = Trim(rs!Aplica)
            End If
            rs.MoveNext
         Loop
         rs.Close
         
         If cAplica <> "N" Then
            '-parto multiple aplica Max 2Hras
            nLacta = IIf(nSonLacta = 1, 1, 2)
            
            If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
               MsgBox "aplica=>[" + cAplica + "] " + Trim(Str(nSonLacta))
            End If
            
         End If
       End If
       
       '-HORARIO del dia -**
       sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Trim(Txtarjeta) & "' AND fecha ='" & fechaB & "'"
       Set rs = New Recordset
       rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       If rs.EOF Then
          If Trim(TxGrpGen) <> "" Then
             rs.Close
             sql = "SELECT * FROM CpHor" & genaño & " WHERE grupo='" & TxGrpGen & "' AND fecha ='" & fechaB & "'"
             Set rs = New Recordset
             rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          End If
       End If
       If Not rs.EOF Then
          cvalor = IIf(IsNull(rs!Horario), "", Trim(rs!Horario))
          Sq0 = "SELECT * FROM CpHorarioNew  WHERE Codigo='" & cvalor & "'"
          Set rs_d = New Recordset
          rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If rs_d.EOF Then
             rs_d.Close
             Sq0 = "SELECT * FROM CpHorarios  WHERE Codigo='" & cvalor & "'"
             Set rs_d = New Recordset
             rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          End If
          If Not rs_d.EOF Then
             p_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
             c_turno = IIf(IsNull(rs_d!turno), "", Trim(rs_d!turno))
             If rs_d!Horas < "1800" Then                               '========>18-25 hras
                p_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
                p_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
             End If
             p_sal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
             
             '-horario ORIGINAL
             '=================
             p0_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
             '-If rs_d!Horas <> "2400" And rs_d!Horas <> "1800" Then                               '=======>
                p0_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                p0_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
             '-End If
             p0_sal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
             
             If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                MsgBox cvalor + "=>" + p0_ing + ":" + p0_sal + "/" + p0_ing1 + "¨:" + p0_sal1
             End If
             
             '-Ajuste horario x LACTANCIA
             '===========================
             If cAplica <> "N" And nLacta >= 1 Then
             
                If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                   MsgBox cAplica + " [" + Trim(Str(nLacta)) + "]"
                End If
                
                If cAplica = "I" Then   'al Ingreso
                   If Val(p_ing) > 0 And Val(p_ing1) = 0 Then
                      phora = Right("00" + Trim(Str(Val(Left(p_ing, 2)) + nLacta)), 2)
                      pmints = Right(p_ing, 2)
                      p_ing = phora + pmints
                   End If
                ElseIf cAplica = "S" Then   'a Salida
                       If Val(p_sal) > 0 And Val(p_sal1) = 0 Then
                          phora = Right("00" + Trim(Str(Val(Left(p_sal, 2)) - nLacta)), 2)
                          pmints = Right(p_sal, 2)
                          p_sal = phora + pmints
                       ElseIf Val(p_sal1) > 0 Then
                              phora = Right("00" + Trim(Str(Val(Left(p_sal1, 2)) - nLacta)), 2)
                              pmints = Right(p_sal1, 2)
                              p_sal1 = phora + pmints
                       End If
                End If
             End If
             
             If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                MsgBox "[" + cAplica + "] " + p_ing + ":" + p_sal + "/" + p_ing1 + "¨:" + p_sal1
             End If
             
             p_horas = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
             nValorH = Val(Left(p_horas, 2))
             nValorM = Val(Right(p_horas, 2))
             nHrProg = (nValorH + (nValorM / 60))
             TxhorasP = Val(TxhorasP) + nHrProg
             
               If Trim(UCase(TxVerAcum)) = "S" And p_horas <> "0000" Then
                  MsgBox "Dia:" + Trim(Str(i)) + "[" + c_turno + "-" + cvalor + "] " + "H.Prog:" + p_horas + " van->" + Str(nHrProg) + " Acumula:" + TxhorasP
               End If
             
             If Trim(c_turno) <> "" And InStr("NHCP", c_turno) > 0 Then
                c2dias = "S"        '1900-0700  0700-1900/1900-0800
             End If
          End If
          rs_d.Close
       End If
       rs.Close
       Textd(z) = Val(p_horas) / 100      'Hras Programadas
       
         If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
            MsgBox " 2dias:[" + c2dias + "] " + Trim(Str(Day(primer + z))) + " Lact:[" + cAplica + "] " + Trim(Str(nSonLacta)) + " Horario->" + p0_ing + "-" + p0_sal + " / " + p_ing + "-" + p_sal
         End If
              
       '-AUTORIZACIONES: Motivos Personales, Comisiones Serv
       cDiahora = "X"
       cMantGuardia = "N"
       sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta='" & Trim(Txtarjeta) & "' and FechaIni<='" + fechaB + "' and FechaFin>='" + fechaB + "'"
       Set rs = New Recordset
       rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       Do While Not rs.EOF
          If Not rs.EOF Then
             cAuPrg = Trim(rs!codigo)
             cDiahora = rs!diahora        'D/H
             cTipo = IIf(IsNull(rs!tipo), "", rs!tipo)            'I/S/O   caso horas
             choraini = IIf(IsNull(rs!horaini), "", rs!horaini)
             chorafin = IIf(IsNull(rs!horafin), "", rs!horafin)
          End If
          If cAuPrg <> "" Then
             sql = "SELECT * FROM CpOcurrencias WHERE codigo='" & cAuPrg & "'"
             Set rs_d = New Recordset
             rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                cDescuenta = IIf(rs_d!DESCUENTO = "", "N", Trim(rs_d!DESCUENTO))
                cMantGuardia = IIf(IsNull(rs_d!MantGuardia) Or Trim(rs_d!MantGuardia) = "", "N", rs_d!MantGuardia)
             End If
             rs_d.Close
             '-Aplica Autorización Entrada(I)/(S)alida
             If cTipo = "I" Then
                m_ing = choraini
             ElseIf cTipo = "S" Then
                    m_sal = chorafin
             End If
          End If
          rs.MoveNext
       Loop
       rs.Close
       
         '-Muestra  AUTORIZACIONES
         If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
            If cAuPrg <> "" Then
               MsgBox "Autoriza [" + cAuPrg + "]" + cTipo + ":" + Trim(Str(Day(primer + z))) + " -> " + p_ing + "-" + p_sal + ": " + m_ing + " , " + m_sal
            End If
         End If
         
       If Val(p_horas) > 0 Then
          '-Marcas de Turno Programado
          Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
          '-Marcas del dia, solo si Justifica "D"
          p_Sale = ""
          If cDiahora <> "D" Then
             If Val(p_sal1) > 0 Or Val(p_sal) < Val(p_ing) Then
                '-Ubica marcas 2dias Actual-siguiente caso: turno t/n, n/d
                fecha1 = Format(primer + z + 1, "yyyymmdd")
                Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha >='" & fechaB & "' AND fecha <='" & fecha1 & "' ORDER BY fecha,hora"
                cdiaant = "S"
                p_Sale = p_sal
                If Val(p_sal1) > 0 Then
                   p_Sale = p_sal1         'Corrige fin turno
                End If
             End If
          End If
       
          '-ASIGNA marcas basado en Horario   I - S  /I1 - S1
          Set rs_d = New Recordset
          rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If Not rs_d.EOF Then
             d_fentra = rs_d!fecha          '========> 1era marca del servidor para turno programado
             Do While Not rs_d.EOF
                '-caso 1300-1900/1900-0700
                If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                   MsgBox Format(rs_d!fecha, "dd/mm/yyyy") + "-" + rs_d!hora
                End If
                cInicio = p_ing
                cFin = rs_d!hora
                dfsale = rs_d!fecha
                If Val(m_ing) = 0 Then
                   'If (nTotMints(cInicio, cFin) >= -60 And nTotMints(cInicio, cFin) <= 60) And Day(dfsale) = i Then
                   If Day(dfsale) = i Then
                          m_ing = rs_d!hora
                   Else
                      m_sal = rs_d!hora
                   End If
                   If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                      MsgBox "1Ing " + Str(i) + ": " + Str(Day(rs_d!fecha)) + ": " + m_ing + "-" + m_sal
                   End If
                ElseIf Val(m_ing) > 0 And (Val(rs_d!hora) - Val(m_ing) < 60) And c2dias = "S" And i = Day(dfsale) Then
                       'Doble Ingreso¿2dias?         ============>
                       m_ing1 = rs_d!hora
                       If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                         MsgBox "2Ing1-" + Str(i) + ": " + Str(Day(dfsale)) + ": " + m_ing1
                       End If
                ElseIf i = Day(dfsale) And Val(m_ing) > 0 And (Val(rs_d!hora) - Val(m_ing) < 60) And ((Val(rs_d!hora) < Val(p_sal) And c2dias = "N")) Then
                       'Doble Ingreso¿?
                       m_ing1 = rs_d!hora
                       If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                          MsgBox "3Ing1-" + Str(i) + ": " + Str(Day(dfsale)) + ": " + m_ing1
                       End If
                       
                ElseIf i = Day(dfsale) And Val(p_ing1) > 0 And Val(m_ing1) = 0 And (Val(rs_d!hora) - Val(p_ing1) < 10) Then
                       '-ing Turno partido
                       m_ing1 = rs_d!hora
                       If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                          MsgBox "4Ing1-" + Str(i) + ": " + Str(Day(dfsale)) + ": " + m_ing1
                       End If
                       
                ElseIf i = Day(dfsale) And Val(p_ing1) = 0 And Val(m_ing1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then
                       m_ing1 = rs_d!hora
                       If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                          MsgBox "5ing-" + Str(i) + ": " + Str(Day(dfsale)) + ": " + (m_ing1)
                       End If
                       
                ElseIf ((Val(m_sal) = 0 And (Val(rs_d!hora) - Val(p_sal) >= 0)) Or (Val(m_sal) = 0 And Val(p_sal) > Val(rs_d!hora))) Then
                       '07001900-19000700   0701 0738x 0740    1352 0712
                       m_sal = rs_d!hora
                       If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                          MsgBox "6Sal-" + Str(i) + ": " + Str(Day(dfsale)) + ": " + m_sal
                       End If
                 
                ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) > 0 And Val(m_sal1) = 0 And (Val(rs_d!hora) - Val(p0_sal1) >= 0) Then       '<=60
                       '-sal Turno partido  1300-1900/1900-0700
                       m_sal1 = rs_d!hora
                       If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                          MsgBox "7Sal1-" + Str(i) + ": " + Str(Day(dfsale)) + ": " + (m_sal1)
                       End If
                        
                '-Check carga marcaciones sin valor para turno
                ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then     'And Val(m_ing1) > 0
                       m_sal1 = rs_d!hora
                       If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                         MsgBox "8sal1-" + Str(i) + ": " + Str(Day(dfsale)) + ": " + (m_sal1)
                       End If
                End If
                rs_d.MoveNext
             Loop
          End If
          rs_d.Close
          
          '-Muestra  marcas realizadas  Txtarjetas
             If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                MsgBox "Marcas [2dias->" + c2dias + "] " + Trim(Str(Day(dfentra))) + "/" + Trim(Str(Day(dfsale))) + " ->  prog:marca =" + p_ing + ":" + m_ing + " , " + p_sal + ":" + m_sal + " / " + p_ing1 + ":" + m_ing1 + " , " + p_sal1 + ":" + m_sal1
             End If
                    
          wwind = ""    '-- Parametro EVALUACION FALTA
          '-Evalua MARCAS Ingreso
          If cDiahora = "D" And cDescuenta = "S" Then
             F_falta = True     'FALTA ¿JUSTIFICADA?
             wwind = "1"
          ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And Val(p_ing1) > 0 And p_sal <> p_ing1 Then  '======>
                 'Caso Horario:2000-0800   marca 0805-2010   O marca 0749-1404 1930-0820 ó  0803-2015/-0824
                 '-07001900-19000700 marca 0701 -07180719 & 1252,1900- ,07
                 ' 13001900-19000700       12521905- ,0715
                 F_falta = True
                 wwind = "&"
          ElseIf (Val(m_ing) + Val(m_ing1) = 0) Then
             '-NO marco Ing,   valida vs autoriza
             If cDiahora = "H" And cTipo = "I" Then
                cInicio = p_ing
                cFin = choraini
                If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                   c_a = c_a + "e"
                   F_falta = True
                   cDescuenta = "S"
                   wwind = "2"
                End If
             ElseIf cDiahora = "D" Then     'valida autoriza Dia Comp/ pierde
                If cDescuenta = "S" Then
                   c_a = c_a + "D"
                   F_falta = True
                   wwind = "3"
                End If
             Else
                c_a = c_a + "i"
                F_falta = True
                wwind = "4"
             End If
          ElseIf (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
             'Marco tarde, check Autoriza H al ingreso
             If cDiahora = "H" And cTipo = "I" Then
                cInicio = p_ing
                cFin = choraini
                If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                   F_falta = True
                   wwind = "5"
                End If
             End If
          ElseIf Val(m_ing) > 0 Then
              '--- INSERTADO Excepciones ----09.08.17
              If Val(m_ing1) = 0 Then
                 '-No hay doble marca
                 '  PI   PS            MI   MS     MI1  MS1
                 '- 2000-0800      :   0740-1401   FALTA
                 '                 :   2130-0835   FALTA Ingreso fuera de Tolerancia
                 '-                :   1750-0820   ok
                 '--               :   0741-2044             marcas fuera Turno (antes/despues)-> FALTA
                 '--               :   0812-1847        0900
                 '---  0700-1900 1900-0700  24Hras  :marcas   0715- / -0748
                 '---  1300-1900 1900-0700  18hrs   :marca    1252-   -0705
                 If (m_sal > m_ing And m_sal < p_ing) And p_horas <> 2400 Then  '¿2400?
                    '-trabajo Antes de Turno
                    F_falta = True
                    wwind = "6"
                 Else
                   ' If i = 31 Then
                   '    MsgBox ("[" + c2dias + "=> " + m_ing + "-" + m_sal + "]")
                   ' End If
                   If c2dias = "S" And ((m_sal > p_sal And d_fentra = dfsale) Or (m_sal < p_sal And d_fentra < dfsale)) Then
                       '07001900-19000700   0716, -07580759
                       '-No completo el turno
                       F_falta = True
                       wwind = "7"
                    End If
                 End If
                 '  If i = 31 Then
                 '     MsgBox ("falta [" + wwind + "]")
                 '  End If
                                 
              Else
                 '-Marca DOBLE
                 '  PI   PS            MI   MS     MI1  MS1
                 '- 2000-0800      :   0740-1401   1930-0830 (a)
                 '-                :   0743-1352   1400-1848 X
                 '-                :   1950-0820   1200-1400
                 '-                :   1941-1942   0801
                 '- 1930-0730      :   1859-1900   0804-0805 x ¿?
                 '- 0730-1930      :   0745-0746   2001
                 '-                :   0741-0742   1935-1936
                 If m_sal > m_ing And m_sal < p_ing Then
                    '-1er pareja Antes de TURNO
                    If CalcMTard(p_ing, m_ing) > 30 Or (m_sal1 <> "" And m_ing1 < p_sal) Or (m_sal1 > m_ing1 And m_sal1 < p_sal) Then
                       '-     1930, 1859                    0805            0804     0730       0805      0804      0805     0730
                       '      2000, 1942                    ""                                  ""        0801       ""      0800
                       '      1930, 0740                    0830            1930     0800       0830      1930      0830     0830
                       '-2do Marca fuera de tolerancia o Turno
                       F_falta = True
                       wwind = "$"
                    End If
                 End If
              End If
              
          End If
          
           '-Muestra  marcas Falta ¿?
           If F_falta Then
              If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                 MsgBox "Evaluo ING :" + Trim(Str(Day(primer + z))) + " " + cDiahora + " ¿ING. Falta? [" + wwind + "] " + IIf(F_falta, "S", "N") + " -> Marcas : " + m_ing + " - " + m_sal + " / " + m_ing1 + " - " + m_sal1
              End If
           End If
           
          '-Evalua MARCAS Salida
          If F_ALTA Then
          ElseIf (Val(m_sal) + Val(m_sal1) = 0) Then
             '-NO marco  Sal   valida vs autoriza
             If cDiahora = "H" And cTipo = "S" Then
                If (Val(p_sal) - Val(chorafin) > 0) Or (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                   c_a = c_a + "f"
                   F_falta = True
                   cDescuenta = "S"
                   wwind = "A"
                End If
             ElseIf cDiahora = "D" Then     'valida autoriza Dia Comp/ pierde
                If cDescuenta = "S" Then
                   c_a = c_a + "2"
                   F_falta = True
                   wwind = "B"
                End If
             Else
                c_a = c_a + "s"
                F_falta = True
                wwind = "C"
             End If
          ElseIf Val(p_sal1) > 0 And p0_sal <> p0_ing1 Then    'p_horas<>"1800 / 2400 / 2500"
                If Val(m_sal1) = 0 Then
                   F_falta = True
                   wwind = "@"
                Else
                   If Val(p_sal1) > Val(m_sal1) Then
                      F_falta = True
                      wwind = "#"
                    End If
                End If
               
          ElseIf Val(m_sal) > 0 And Val(p0_sal) > Val(m_sal) Then       'p_sal
             'Marco antes de, Autoriza H
             If cDiahora = "H" And cTipo = "S" Then    '====> And cTipo = "S"  ** TO CHECK **
                cInicio = p_sal
                cFin = chorafin
                If Val(p_sal) - Val(chorafin) > 0 Then
                   F_falta = True
                   wwind = "D"
                End If
             ElseIf cDiahora = "X" Then
                '----------cInicio = p_sal
                If p0_sal > m_sal And m_ing1 <> "" And m_sal1 = "" Then
                   'Correcion caso   11/May/15
                   '04769(20/05):0758-0800-2001
                   '01951(12/05):0759-0800-1501
                   cFin = m_ing1
                Else
                   cFin = IIf(m_sal1 <> "", m_sal1, m_sal)
                End If
                If Val(p_sal) - Val(cFin) > 0 Then
                   F_falta = True
                   c_a = c_a + "x"
                   wwind = "E"
                   
                ElseIf Val(p_sal1) - Val(cFin) > 0 Then
                       F_falta = True
                       c_a = c_a + "y"
                       wwind = "H"
                End If
                
             End If
          Else
               If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                  MsgBox Trim(Str(Day(primer + z))) + " " + cDiahora + " SAL. Falt[" + wwind + "]" + IIf(F_falta, "S", "N") + " ->  prog:marca:  " + p_ing + ":" + m_ing + " , " + p0_sal + ":" + m_sal + " / " + p0_ing1 + ":" + m_ing1 + " , " + p_sal1 + ":" + m_sal1
               End If
          End If
              
           '-Muestra  marcas Falta ¿?
           If F_falta Then
              If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                MsgBox "EVALUO SAL: " + Trim(Str(Day(primer + z))) + " " + cDiahora + " ¿SAL. Falta?  [" + wwind + "]" + IIf(F_falta, "S", "N") + " ->  Marcas:  " + m_ing + " , " + m_sal + " / " + m_ing1 + " , " + m_sal1
              End If
           End If
              
          If F_falta Then
          Else
             If cDiahora = "H" And cTipo = "I" Then     '...1
             ElseIf cDiahora <> "D" Then      '<===No Autorizaciones D  ...1
                If (Val(m_ing) + Val(m_ing1) > 0) Then
                   '-Hay marcas, aplica tolerancia, calcula tardanza
                   If cDiahora = "H" And cTipo = "I" Then
                      cInicio = p_ing
                      cFin = choraini
                      nmintarde = CalcMints(cInicio, cFin)
                   Else
                      nmintarde = CalcMTard(p_ing, m_ing)
                   End If
                   If nmintarde < 6 Then
                      nmintarde = 0
                   End If
                   If nmintarde >= 6 And nmintarde < 31 Then
                      If nmintarde >= 6 And nmintarde <= 10 Then
                         ndscto = 10
                      End If
                      If nmintarde >= 11 And nmintarde <= 20 Then
                         ndscto = 20
                      End If
                      If nmintarde >= 21 And nmintarde < 31 Then
                         ndscto = 30
                      End If
                      'acumula tardanza  /acumula tm30
                      nTtarda = nTtarda + nmintarde
                      nTDscto = nTDscto + ndscto
                   Else
                      If nmintarde > 30 Then
                         c_a = c_a + "4"
                         F_falta = True
                         wwind = "T"
                      End If
                      nmintarde = 0
                   End If
                End If
             End If              '...1
          End If
           '---LISTA EVALUA marcas----
           If falta Then
              If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                MsgBox "Marcas tardanza : " + Trim(Str(Day(primer + z))) + " falta:[" + wwind + "] ->  " + m_ing + " , " + m_sal + " / " + m_ing1 + " , " + m_sal1
              End If
           End If
           
      Else
          '-MARCAS en Dia  NO PROGRAMADO,  fecha  hora =>  hhmm
          '- Retira marca final de turno Dia Anterior, luego debe MOSTRAR
          Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
          Set rs_d = New Recordset
          rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If Not rs_d.EOF Then
             Do While Not rs_d.EOF
                If Val(TxVerDia) > 0 And i = Val(TxVerDia) Then
                   MsgBox "marca " + Str(i) + ": " + Str(Day(dfsale)) + "=> " + rs_d!hora + " " + p_Sale
                End If
             
                If (c2dias = "N" And Val(p_Sale) = 0) Or (Val(p_Sale) > 0 And nTotMints(p_Sale, rs_d!hora) > 120) Then
                   cInicio = m_ing
                   cFin = rs_d!hora
                   If Trim(m_ing) = "" Or (Val(m_ing) > 0 And (Val(rs_d!hora) - Val(m_ing) < 5)) Then
                      m_ing = rs_d!hora
                   ElseIf Trim(m_ing) <> "" And CalcMints(cInicio, cFin) > 60 Then
                          m_sal = rs_d!hora
                   ElseIf Trim(m_sal) <> "" Or (Val(rs_d!hora) - Val(m_sal) < 5) Then
                          m_ing1 = rs_d!hora
                   End If
                End If
                rs_d.MoveNext
             Loop
          End If
          p_Sale = ""    'apaga efecto turno anterior N
          rs_d.Close
       End If
       Text1(z).ForeColor = &H80000012      'negro
       If F_falta Then              '...A
          Text1(z).ForeColor = &HFF             'rojo
          Text1(z) = "F" + cDescuenta + c_a
          nvalor = nValorH + (nValorM / 60)                    'hrs PROG
          If cAuPrg <> "" And cAuPrg <> "FI" And cDiahora = "D" And cDescuenta = "S" Then
             If cAuPrg = "MP" Then
                Txhorasfmp = Val(Txhorasfmp) + nvalor             'acumula hrs Faltas MP
                nDMP = nDMP + 1
                If InStr("D_N", Left(Lbltur(z).Caption, 1)) > 0 Then
                   nDMP = nDMP + 1
                End If
             Else
                TxhorasFj = Val(TxhorasFj) + nvalor             'acumula FJ x DIA
                nDFJ = nDFJ + 1
                If InStr("D_N", Left(Lbltur(z).Caption, 1)) > 0 Then
                   nDFJ = nDFJ + 1
                End If
             End If
          Else
             If cAuPrg = "" Or cAuPrg = "FI" Or cDiahora = "H" Then
                TxhorasFI = Val(TxhorasFI) + nvalor               'acumula
                nDFI = nDFI + 1
                If InStr("D_N", Left(Lbltur(z).Caption, 1)) > 0 Then
                   nDFI = nDFI + 1
                End If
                'Corrige Descuento
                cDescuenta = "S"
             End If
          End If
       Else
          If p_ing <> "" Or cdiaant = "S" Then
             Text1(z) = ndscto
             TxmTarda = Val(TxmTarda) + nmintarde              'acumulado Tardanzas
             TxmDscto = Val(TxmDscto) + ndscto                 'acumulado Trm130
             If cDiahora = "H" Then
                If cDescuenta = "S" Then
                   If cAuPrg = "MP" Then
                      cInicio = choraini
                      cFin = chorafin
                      nmotpersonal = CalcMints(cInicio, cFin)
                      If nmotpersonal < 0 And c2dias = "S" Then
                         nmotpersonal = nmotpersonal + 24
                      End If
                      If nmotpersonal Mod 60 > 30 Then
                         nDsctoMP = Int(nmotpersonal / 60) + 1
                      Else
                         nDsctoMP = Int(nmotpersonal / 60)
                      End If
                      Txhorasmp = Val(Txhorasmp) + nDsctoMP        'acumula Hras MP
                   Else
                      notros = Val(Left(chorafin, 2)) - Val(Left(choraini, 2)) + (Val(Right(chorafin, 2)) - Val(Right(choraini, 2))) / 60
                      If notros < 0 And c2dias = "S" Then
                         notros = notros + 24
                      End If
                      nTotros = nTotros + notros
                   End If
                Else
                   nComision = Val(Left(chorafin, 2)) - Val(Left(choraini, 2)) + (Val(Right(chorafin, 2)) - Val(Right(choraini, 2))) / 60
                   If nComision < 0 And c2dias = "S" Then
                      nComision = nComision + 24
                   End If
                   TxHorasCom = Val(TxHorasCom) + nComision        'acumula
                End If
             End If
             MSG = ""
             nvalor = nValorH + ((nValorM - ndscto) / 60)       'Hrs PROG - Dscto(tardanza)
             VerHrsMint (nvalor - nDsctoMP - notros)
             Textd(z) = MSG      'Hras trabajadas
          End If
       End If           '...A
       rt.AddNew
       '- nlacta>=1   => fecha en periodo de LACTANCIA
       rt!TARJETA = Txtarjeta
       rt!PLH = TxPlh
        rt!DscPLH = cDscPLH
        W_RegPlh = cDscPLH
       rt!dscactividad = Left(cDActividad, 30)
       rt!dscccosto = Left(cDccosto, 20)
       rt!dsccarest = Left(cDcarest, 30)
       rt!nombdia = Left(WeekdayName(Weekday(primer + z)), 3)
       rt!fecha = primer + z
       rt!turno = Lbltur(z).Caption
       rt!Horario = cvalor
       '-- Turno Original
       rt!ping = "" + p0_ing
       rt!psal = "" + p0_sal
       rt!ping1 = "" + p0_ing1
       rt!psal1 = "" + p0_sal1
       rt!hrasprog = nHrProg
       rt!HrProg = IIf(nHrProg = 0, "", cHorsMint(nHrProg))
       rt!ming = "" + m_ing
       rt!msal = "" + m_sal
       rt!ming1 = "" + m_ing1
       rt!msal1 = "" + m_sal1
       rt!marca2 = IIf(Val(m_ing1) > 0, m_sal, "")
       rt!marca3 = IIf(Val(m_ing1) = 0, m_sal, m_ing1)
       rt!marca4 = IIf(Val(m_sal1) > 0, m_sal1, "")
       rt!tardanza = 0 + IIf(F_falta, 0, nmintarde)
       rt!DESCUENTO = 0 + IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", ndscto, 0))
       nDescuento = IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", ndscto, 0))
       rt!autorizac = "" + cAuPrg
       rt!descuenta = "" + cDescuenta
       rt!diahora = "" + cDiahora
       rt!tipo = "" + cTipo
       rt!horaini = "" + choraini
       rt!horafin = "" + chorafin
       rt!HrIni = IIf(choraini = "", "", Left(choraini, 2) + ":" + Right(choraini, 2))
       rt!HrFin = IIf(chorafin = "", "", Left(chorafin, 2) + ":" + Right(chorafin, 2))
       rt!falta = IIf(F_falta, "S", "N")
       rt!Comision = nComision
       rt!motpersonal = nmotpersonal
       rt!DescuentoMP = 0 + IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", nDsctoMP, 0))
       rt!DatoMp = IIf(cAuPrg <> "" And cAuPrg <> "FI", IIf(nmotpersonal = 0, IIf(nComision > 0, cHorsMint(nComision), ""), cHorsMint(nmotpersonal)), "")
       
       If cDiahora = "D" And (cAuPrg = "COMP" Or Left(cAuPrg, 2) = "VA") And InStr("DNPCH", Left(Lbltur(z).Caption, 1)) > 0 Then
          '--04.10.2023 solo reconoce 6hras trabajo PIERDE GUARDIA
          nHrTrab = 6
          rt!HrasTrab = nHrTrab
          TxhorasFj = Val(TxhorasFj) + (nHrProg - nHrTrab)
       Else
          rt!HrasTrab = IIf(F_falta, IIf(cDescuenta = "S", 0, nvalor), IIf(p_ing <> "" Or cdiaant = "S", nvalor - nDsctoMP - notros, 0))
          nHrTrab = IIf(F_falta, IIf(cDescuenta = "S", 0, nvalor), IIf(p_ing <> "" Or cdiaant = "S", nvalor - nDsctoMP - notros, 0))
       End If
       rt!HrTrab = IIf(nHrTrab > 0, cHorsMint(nHrTrab), "")
       rt!hrasfalt = IIf(F_falta, IIf(cDescuenta = "S", nHrProg, 0), 0)
       nHrFalt = IIf(F_falta, IIf(cDescuenta = "S", nHrProg, 0), 0)
       c_Falta = "Falta"
       If nHrTrab = 0 And cDiahora = "D" Then
           '-autorizado "D" faltar
          c_Falta = ""
       End If
       rt!Descta = IIf(nHrFalt > 0, IIf(primer + z < Date, c_Falta, "* P *"), IIf(nDescuento > 0, CStr(Round(nDescuento, 2)), ""))
       rt!Valido = IIf(p_ing <> "" Or cdiaant = "S", "S", "")
       '-Campo GRUPO: usado para indicar en periodo de LACTANCIA
       rt!grupo = IIf(nLacta = 0, "", IIf(nLacta = 1, "Lact", "Lac2"))
       rt.Update
       
            If Trim(UCase(TxVerAcum)) = "S" And nHrProg > 0 Then
               MsgBox "graba " + Format(primer + z, "dd/mm/yyyy") + " HProg:" + Format(nHrProg, "##.0#")
            End If
       
       If c2dias = "N" And cdiaant = "S" Then
          '-No turno 2dias y marca corresponde turno Dia Anterior
          cdiaant = "N"
       End If
   Next i
   rt.Close
   
   nvalor = 0
   If nTtarda > 60 Then        ' *** PARA RESOLUCIONES ***
      nvalor = nTDscto / 60     'aplica Descuento por Tardanzas en Hras
      If nTDscto Mod 60 > 30 Then
         nvalor = Int(nvalor) + 1
      Else
         nvalor = Int(nvalor)
      End If
   Else
      TxmDscto = ""
   End If
   TxHorast = Val(TxhorasP) - Val(TxhorasFI) - Val(TxhorasFj) - Val(Txhorasfmp) - Val(Txhorasmp) - nTotros - nvalor
   MSG = ""
            If Trim(UCase(TxVerAcum)) = "S" Then
               MsgBox "inicio total=>" + TxhorasP
            End If
   
   VerHrsMint (Val(TxhorasP))
   TxhorasP.Text = MSG
            If Trim(UCase(TxVerAcum)) = "S" Then
               MsgBox "verHrsMint=(" + MSG + ")"
            End If
            
    MSG = ""
    ConvertDecimalenHora (Val(TxhorasP))
            If Trim(UCase(TxVerAcum)) = "S" Then
               MsgBox TxhorasP + "=>convert=(" + MSG + ")"
            End If
   
   MSG = ""
   VerHrsMint (Val(Txhorasfmp))
   Txhorasfmp = MSG
   MSG = ""
   VerHrsMint (Val(TxhorasFj))
   TxhorasFj = MSG
   MSG = ""
   VerHrsMint (Val(TxhorasFI))
   TxhorasFI = MSG
   MSG = ""
   VerHrsMint (Val(Txhorasmp))
   Txhorasmp = MSG
   MSG = ""
   VerHrsMint (Val(TxHorast))
   TxHorast = MSG
   MSG = ""
   VerHrsMint (Val(TxHorasCom))
   TxHorasCom = MSG
   
      '  Beep 500, 350
      '  MsgBox "hras dscto x tardanza :" + Trim(Str(nvalor))
   
   Cmdetalle.Enabled = True
   Cmdetalle.SetFocus
End Sub
Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
      Case 1
   End Select
   Txfecha.Text = primer + Index
   Cmdetalle_Click
End Sub
Private Sub Cmdetalle_Click()
   If Val(TxHorast.Text) + Val(TxhorasFj.Text) + Val(TxhorasFI.Text) = 0 Then Exit Sub
   Load RecMarDet
   RecMarDet.Left = 3090: RecMarDet.Top = 4500
   RecMarDet.Refresh
   RecMarDet.Show
End Sub
Private Sub CmdPerTarjeta_Click()
     If Trim(Txtrabajador) = "" Then
        MsgBox "No indica trabajador..", vbExclamation, "[ ERROR ]"
        Exit Sub
     End If
        '-Cortado/sabana
        Load RecPerTARJETA
        RecPerTARJETA.Left = 0: RecPerTARJETA.Top = 0
        RecPerTARJETA.Refresh
        RecPerTARJETA.Show
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
   If Txtarjeta = "" Then
      MsgBox "No definio servidor..", vbExclamation
      Exit Sub
   End If
   If Val(TxhorasP) + Val(TxHorast) = 0 Then
      MsgBox "Trabajador, No Tiene Horas Programadas", vbExclamation
   End If
   nTipRpt = 5
   InfoCpRS.Left = 0: InfoCpRS.Top = 0
   InfoCpRS.Refresh
   InfoCpRS.Show
   Exit Sub
End Sub
Private Sub cmdsalir_Click()
   Set rs_d = Nothing
   Set rs = Nothing
   Set rt = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        '- Forms Abiertos
        If UCase(Forms(cont).Name) <> "CPMENU" And UCase(Forms(cont).Name) <> "RECMARCASRS" Then
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
        '- Form Abiertos
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "RecMarcasRS" Then
           Unload Forms(cont)
        End If
      End If
   Next
    '-Kill TEMPORAL
    'Cnx.Execute "DROP table #tabla1"
   Me.Hide
   Unload Me
End Sub

