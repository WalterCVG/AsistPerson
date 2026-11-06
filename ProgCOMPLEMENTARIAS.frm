VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ProgComplementarios 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Programa HORAS COMPLEMENTARIAS"
   ClientHeight    =   10515
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15075
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "ProgCOMPLEMENTARIAS.frx":0000
   ScaleHeight     =   10515
   ScaleMode       =   0  'User
   ScaleWidth      =   15075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Tx_grupo 
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   199
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Txactividad 
      Enabled         =   0   'False
      Height          =   195
      Left            =   3120
      MaxLength       =   4
      TabIndex        =   197
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxDirector 
      Enabled         =   0   'False
      Height          =   195
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   196
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxLstJefServ 
      Enabled         =   0   'False
      Height          =   285
      Left            =   8520
      TabIndex        =   195
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxUprCCosto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7920
      MaxLength       =   2
      TabIndex        =   194
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxUPRjefaturas 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7440
      MaxLength       =   1
      TabIndex        =   193
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton CmdAudDirect 
      Height          =   555
      Left            =   11280
      Picture         =   "ProgCOMPLEMENTARIAS.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   192
      ToolTipText     =   "Auditar cambios"
      Top             =   9120
      Width           =   645
   End
   Begin VB.TextBox TxCadHep 
      Enabled         =   0   'False
      Height          =   195
      Left            =   3480
      MaxLength       =   52
      TabIndex        =   183
      Top             =   5040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Cmdchk 
      Caption         =   "Check"
      Height          =   495
      Left            =   9240
      Picture         =   "ProgCOMPLEMENTARIAS.frx":80EA8
      Style           =   1  'Graphical
      TabIndex        =   182
      ToolTipText     =   "Verifica Registro"
      Top             =   5280
      Width           =   615
   End
   Begin VB.CommandButton CmdLIMPIA 
      Caption         =   "Limpia"
      Height          =   615
      Left            =   14160
      Picture         =   "ProgCOMPLEMENTARIAS.frx":811B2
      Style           =   1  'Graphical
      TabIndex        =   181
      ToolTipText     =   "Limpia Plantilla"
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   615
      Left            =   9480
      Picture         =   "ProgCOMPLEMENTARIAS.frx":816E4
      Style           =   1  'Graphical
      TabIndex        =   180
      ToolTipText     =   "Eliminar"
      Top             =   9120
      Width           =   735
   End
   Begin VB.TextBox TxRegPlh 
      Enabled         =   0   'False
      Height          =   225
      Left            =   13560
      MaxLength       =   3
      TabIndex        =   177
      Text            =   "   "
      Top             =   9120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboRegPlh 
      Height          =   315
      Left            =   12000
      TabIndex        =   176
      Top             =   9360
      Width           =   2175
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
      Height          =   195
      Left            =   240
      TabIndex        =   174
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Tx_ccosto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   173
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame4 
      Caption         =   "Seguridad"
      Height          =   735
      Left            =   11400
      TabIndex        =   171
      Top             =   120
      Width           =   2655
      Begin VB.TextBox Txclave 
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   10
         PasswordChar    =   "&"
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Clave :"
         Height          =   255
         Left            =   240
         TabIndex        =   172
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox TxcodUPR 
      Enabled         =   0   'False
      Height          =   195
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxdscUpr 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5520
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Txbuscar 
      Height          =   285
      Left            =   10080
      MaxLength       =   5
      TabIndex        =   9
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Txbuscar1 
      Height          =   285
      Left            =   12000
      MaxLength       =   12
      TabIndex        =   10
      Top             =   3600
      Width           =   1695
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
      Left            =   14280
      Picture         =   "ProgCOMPLEMENTARIAS.frx":81856
      Style           =   1  'Graphical
      TabIndex        =   167
      ToolTipText     =   "Directorio"
      Top             =   9120
      Width           =   645
   End
   Begin VB.ComboBox Cbo_Upr 
      Height          =   315
      Left            =   5520
      TabIndex        =   4
      Top             =   240
      Width           =   5775
   End
   Begin VB.TextBox Tx_codupr 
      Enabled         =   0   'False
      Height          =   195
      Left            =   10560
      MaxLength       =   2
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Caption         =   "TURNOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1215
      Left            =   240
      TabIndex        =   165
      Top             =   8640
      Width           =   1815
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "D: Día          M: Mañana    T: Tarde        N: Noche "
         ForeColor       =   &H00004080&
         Height          =   855
         Left            =   240
         TabIndex        =   168
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox TxCargo 
      Enabled         =   0   'False
      Height          =   195
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   20
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxPlH 
      Enabled         =   0   'False
      Height          =   195
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Txccosto 
      Enabled         =   0   'False
      Height          =   195
      Left            =   840
      MaxLength       =   2
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Txtrabajador 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   35
      TabIndex        =   12
      Top             =   5400
      Width           =   3615
   End
   Begin VB.TextBox TxTurno 
      Enabled         =   0   'False
      Height          =   195
      Left            =   7920
      MaxLength       =   2
      TabIndex        =   160
      Top             =   8880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboTurDir 
      Height          =   315
      Left            =   4320
      TabIndex        =   161
      Top             =   9120
      Width           =   4215
   End
   Begin VB.TextBox TxDscCargo 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      MaxLength       =   35
      TabIndex        =   14
      Top             =   5400
      Width           =   2895
   End
   Begin VB.TextBox TxCarEst 
      Enabled         =   0   'False
      Height          =   195
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   14280
      Picture         =   "ProgCOMPLEMENTARIAS.frx":82120
      Style           =   1  'Graphical
      TabIndex        =   158
      ToolTipText     =   "Cerrar"
      Top             =   9840
      Width           =   615
   End
   Begin VB.TextBox TxHorario 
      Enabled         =   0   'False
      Height          =   195
      Left            =   7800
      MaxLength       =   4
      TabIndex        =   162
      Top             =   9840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox CboHorario 
      Height          =   315
      Left            =   4320
      TabIndex        =   164
      Top             =   9480
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   21
      Top             =   5760
      Width           =   14535
      Begin VB.TextBox TxPa 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   188
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox TxNa 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   187
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox TxMT 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   184
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox Txt 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   144
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox TxM 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   142
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox TxN 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   140
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox TxD 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   138
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
         Left            =   10920
         MaxLength       =   4
         TabIndex        =   134
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
         Left            =   9600
         MaxLength       =   4
         TabIndex        =   130
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
         Left            =   8400
         MaxLength       =   4
         TabIndex        =   126
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
         Left            =   7080
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
         Index           =   26
         Left            =   5760
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
         Index           =   25
         Left            =   4560
         MaxLength       =   4
         TabIndex        =   156
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
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   152
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   148
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
         Height          =   315
         Index           =   22
         Left            =   720
         MaxLength       =   4
         TabIndex        =   112
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
         Left            =   13440
         MaxLength       =   4
         TabIndex        =   108
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
         Left            =   12240
         MaxLength       =   4
         TabIndex        =   104
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
         Left            =   10920
         MaxLength       =   4
         TabIndex        =   100
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
         Left            =   9600
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
         Index           =   17
         Left            =   8400
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
         Index           =   16
         Left            =   7080
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
         Index           =   15
         Left            =   5760
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
         Index           =   14
         Left            =   4560
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
         Index           =   13
         Left            =   3240
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
         Index           =   12
         Left            =   1920
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
         Height          =   315
         Index           =   11
         Left            =   720
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
         Index           =   10
         Left            =   13440
         MaxLength       =   4
         TabIndex        =   64
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
         Left            =   12240
         MaxLength       =   4
         TabIndex        =   60
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
         Left            =   10920
         MaxLength       =   4
         TabIndex        =   56
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
         Left            =   9600
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
         Index           =   6
         Left            =   8400
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
         Index           =   5
         Left            =   7080
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
         Index           =   4
         Left            =   5760
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
         Index           =   3
         Left            =   4560
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
         Index           =   2
         Left            =   3240
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
         Index           =   1
         Left            =   1920
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
         Height          =   315
         Index           =   0
         Left            =   720
         MaxLength       =   4
         TabIndex        =   24
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
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   146
         Top             =   2400
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   30
         Left            =   10560
         TabIndex        =   132
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   29
         Left            =   9240
         TabIndex        =   128
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   28
         Left            =   8040
         TabIndex        =   124
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   27
         Left            =   6720
         TabIndex        =   120
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   26
         Left            =   5400
         TabIndex        =   116
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   25
         Left            =   4200
         TabIndex        =   154
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   24
         Left            =   2880
         TabIndex        =   150
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   23
         Left            =   1560
         TabIndex        =   114
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   22
         Left            =   360
         TabIndex        =   110
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   21
         Left            =   13080
         TabIndex        =   106
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   20
         Left            =   11880
         TabIndex        =   102
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   19
         Left            =   10560
         TabIndex        =   98
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   18
         Left            =   9240
         TabIndex        =   94
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   17
         Left            =   8040
         TabIndex        =   90
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   16
         Left            =   6720
         TabIndex        =   86
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   15
         Left            =   5400
         TabIndex        =   82
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   14
         Left            =   4200
         TabIndex        =   78
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   13
         Left            =   2880
         TabIndex        =   74
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   70
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   66
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   10
         Left            =   13080
         TabIndex        =   62
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   9
         Left            =   11880
         TabIndex        =   58
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   8
         Left            =   10560
         TabIndex        =   54
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   7
         Left            =   9240
         TabIndex        =   50
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   8040
         TabIndex        =   46
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   5
         Left            =   6720
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   38
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   34
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   26
         Top             =   360
         Width           =   855
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
         Height          =   285
         Index           =   0
         Left            =   360
         MaxLength       =   2
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "Text2"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   2
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "Text3"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   3
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   35
         Text            =   "Text4"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   4
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   5
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   43
         Text            =   "Text6"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   6
         Left            =   8040
         MaxLength       =   2
         TabIndex        =   47
         Text            =   "Text7"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   7
         Left            =   9240
         MaxLength       =   2
         TabIndex        =   51
         Text            =   "Text8"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   8
         Left            =   10560
         MaxLength       =   2
         TabIndex        =   55
         Text            =   "Text9"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   9
         Left            =   11880
         MaxLength       =   2
         TabIndex        =   59
         Text            =   "Text10"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   10
         Left            =   13080
         MaxLength       =   2
         TabIndex        =   63
         Text            =   "Text11"
         Top             =   600
         Width           =   375
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
         Height          =   285
         Index           =   11
         Left            =   360
         MaxLength       =   2
         TabIndex        =   67
         Text            =   "Text12"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   12
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   71
         Text            =   "Text13"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   13
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   75
         Text            =   "Text14"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   14
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   79
         Text            =   "Text15"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   15
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   83
         Text            =   "Text16"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   16
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   87
         Text            =   "Text17"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   17
         Left            =   8040
         MaxLength       =   2
         TabIndex        =   91
         Text            =   "Text18"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   18
         Left            =   9240
         MaxLength       =   2
         TabIndex        =   95
         Text            =   "Text19"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   19
         Left            =   10560
         MaxLength       =   2
         TabIndex        =   99
         Text            =   "Text20"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   20
         Left            =   11880
         MaxLength       =   2
         TabIndex        =   103
         Text            =   "Text21"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   21
         Left            =   13080
         MaxLength       =   2
         TabIndex        =   107
         Text            =   "Text22"
         Top             =   1200
         Width           =   375
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
         Height          =   285
         Index           =   22
         Left            =   360
         MaxLength       =   2
         TabIndex        =   111
         Text            =   "Text23"
         Top             =   1800
         Width           =   375
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
         Height          =   285
         Index           =   23
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   147
         Text            =   "Text24"
         Top             =   1800
         Width           =   375
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
         Height          =   285
         Index           =   24
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   151
         Text            =   "Text25"
         Top             =   1800
         Width           =   375
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
         Height          =   285
         Index           =   25
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   155
         Text            =   "Text26"
         Top             =   1800
         Width           =   375
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
         Height          =   285
         Index           =   26
         Left            =   5400
         MaxLength       =   2
         TabIndex        =   117
         Text            =   "Text27"
         Top             =   1800
         Width           =   375
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
         Height          =   285
         Index           =   27
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   121
         Text            =   "Text28"
         Top             =   1800
         Width           =   375
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
         Height          =   285
         Index           =   28
         Left            =   8040
         MaxLength       =   2
         TabIndex        =   125
         Text            =   "Text29"
         Top             =   1800
         Width           =   375
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
         Height          =   285
         Index           =   29
         Left            =   9240
         MaxLength       =   2
         TabIndex        =   129
         Text            =   "Text30"
         Top             =   1800
         Width           =   375
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
         Height          =   285
         Index           =   30
         Left            =   10560
         MaxLength       =   2
         TabIndex        =   133
         Text            =   "Text31"
         Top             =   1800
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opt1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "PA"
         Height          =   255
         Left            =   10800
         TabIndex        =   190
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "NA"
         Height          =   255
         Left            =   10440
         TabIndex        =   189
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "MT"
         Height          =   255
         Left            =   10080
         TabIndex        =   185
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "M"
         Height          =   255
         Left            =   12000
         TabIndex        =   141
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "T"
         Height          =   255
         Left            =   12360
         TabIndex        =   143
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "N"
         Height          =   255
         Left            =   11640
         TabIndex        =   139
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "D"
         Height          =   255
         Left            =   11280
         TabIndex        =   137
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
         Left            =   8760
         TabIndex        =   136
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
         Left            =   11280
         TabIndex        =   135
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
         Left            =   9960
         TabIndex        =   131
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
         Left            =   8760
         TabIndex        =   127
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
         Left            =   7440
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
         Index           =   26
         Left            =   6120
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
         Index           =   25
         Left            =   4920
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
         Index           =   24
         Left            =   3600
         TabIndex        =   153
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
         Left            =   2280
         TabIndex        =   149
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
         Left            =   1080
         TabIndex        =   113
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
         Left            =   13800
         TabIndex        =   109
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
         Left            =   12600
         TabIndex        =   105
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
         Left            =   11280
         TabIndex        =   101
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
         Left            =   9960
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
         Index           =   17
         Left            =   8760
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
         Index           =   16
         Left            =   7440
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
         Index           =   15
         Left            =   6120
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
         Index           =   14
         Left            =   4920
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
         Index           =   13
         Left            =   3600
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
         Index           =   12
         Left            =   2280
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
         Index           =   11
         Left            =   1080
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
         Index           =   10
         Left            =   13800
         TabIndex        =   65
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
         Left            =   12600
         TabIndex        =   61
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
         Left            =   11280
         TabIndex        =   57
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
         Left            =   9960
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
         Index           =   6
         Left            =   8760
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
         Index           =   5
         Left            =   7440
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
         Index           =   4
         Left            =   6120
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
         Index           =   3
         Left            =   4920
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
         Index           =   2
         Left            =   3600
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
         Index           =   1
         Left            =   2280
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
         Index           =   0
         Left            =   1080
         TabIndex        =   25
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label3 
         Caption         =   "HORAS Programadas"
         Height          =   255
         Left            =   12840
         TabIndex        =   145
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   615
      Left            =   14160
      Picture         =   "ProgCOMPLEMENTARIAS.frx":8226A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Ubicar Registros"
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Txtarjeta 
      Height          =   285
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   11
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4080
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
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2415
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   4260
      _Version        =   393216
      ForeColor       =   0
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
      Caption         =   "ROL DE HORAS COMPLEMENTARIAS"
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
   Begin VB.CommandButton cmdGraba 
      Caption         =   "Graba"
      Height          =   615
      Left            =   8640
      Picture         =   "ProgCOMPLEMENTARIAS.frx":82445
      Style           =   1  'Graphical
      TabIndex        =   179
      ToolTipText     =   "Grabar"
      Top             =   9120
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   975
      Left            =   120
      TabIndex        =   198
      Top             =   3960
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   1720
      _Version        =   393216
      ForeColor       =   64
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
      Caption         =   "DIRECTORIO DEL SERVIDOR"
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
      BackStyle       =   0  'Transparent
      Caption         =   "LblDS"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   11880
      TabIndex        =   191
      Top             =   10080
      Width           =   2055
   End
   Begin VB.Label Lblpuntero 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntero :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   240
      TabIndex        =   186
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Regimen Laboral                             "
      Height          =   255
      Left            =   12240
      TabIndex        =   178
      Top             =   9120
      Width           =   1335
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
      Height          =   375
      Left            =   10680
      TabIndex        =   175
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarjeta"
      Height          =   255
      Left            =   9480
      TabIndex        =   170
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Apel. Paterno"
      Height          =   255
      Left            =   10920
      TabIndex        =   169
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo U.P.R."
      Height          =   375
      Left            =   4920
      TabIndex        =   166
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Turno "
      Height          =   255
      Left            =   3720
      TabIndex        =   163
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
      Left            =   4680
      TabIndex        =   159
      Top             =   8880
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Horario"
      Height          =   255
      Left            =   3720
      TabIndex        =   157
      Top             =   9480
      Width           =   615
   End
   Begin VB.Label Label_3 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVIDOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   5400
      Width           =   1215
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
Attribute VB_Name = "ProgComplementarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
 Dim rs_d1 As New ADODB.Recordset
 Dim rs_2 As New ADODB.Recordset
Dim rD As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim D0 As New ADODB.Recordset
Dim flgDir, Sw, f_sigue As Boolean
Dim nposicion, nJefServ, cCriterio, dfi, dff, z, ndia, npuntDia, npuntero, RowSelected As Integer
Dim sql, SqC, lstJefServ, cvalor, ceditor, c_a, ccampo, cTarjeta As String
Dim cCadTurDIR, cCadTurHRC, cTurnoDAnt, cTurnoDSig, cTurnoAnt, cTurnoDIrSig, cTurnoDirAnt As String
Dim unafecha As Date
Dim onofecha, fecCierre As Date
Dim cvalMt, cvalD, cvalN, cvalM, cvalT, cvalH, cvalC, cvalP As String
Private Sub Form_Load()
  wForma = "PROGCOMPLEMENTARIOS"
  flgDir = False
  Me.Caption = "Programa HORAS COMPLEMENTARIAS - " & genaño & "   -   Sesión: " + wUsuNombre
  LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
          cvalMt = "xxoo"
          cvalD = "xxoo"
          cvalN = "ooxx"
          cvalM = "xooo"
          cvalT = "oxoo"
          cvalH = "xoxx"
          cvalC = "oxxx"
          cvalP = "xxxx"
     

  '-HEP/CAS  no Aplica Restriciones
  Frame4.Enabled = False
  
  DataGrid1.Enabled = False
  For i = 1 To 31
    z = i - 1
    Option1(z).Value = False
    Option1(z).Enabled = False
    Option1(z).Visible = False
    Textd(z).Text = "":   Textd(z).ForeColor = &H0&
    Textp(z).Text = "":   Textp(z).ForeColor = &H0&
    Textd(z).Locked = True
    Textp(z).Locked = True
    Textd(z).Visible = False
    Textp(z).Visible = False
    Lblhr(z).Caption = "":    Lblhr(z).ForeColor = &H0&
  Next
  CboHorario.Enabled = False
  '-Regimen Laboral
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpTablasAux WHERE grupo='REG' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboRegPlh.ListCount - 1
           If CboRegPlh.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboRegPlh.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  '-Meses
  Set rs = New Recordset
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
  
  Txmes = genmes
  For i = 0 To CboMes.ListCount - 1
     npos = InStr(CboMes.List(i), ":")
     nlen = Len(CboMes.List(i))
     dato = Trim(Right(CboMes.List(i), nlen - npos))
     If dato = Trim(Txmes) Then
        CboMes.ListIndex = i
        Exit For
     End If
  Next i
  Call Cbomes_KeyPress(13)
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpUndPrgRol  ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To Cbo_Upr.ListCount - 1
         If Cbo_Upr.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("Codupr") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        Cbo_Upr.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("Codupr")
     End If
     rs.MoveNext
  Wend
  rs.Close
    
  CboTurDir.Enabled = False
  '-Turnos Habile"S" para programar Asistencial  6/12 hras
  Set rs = New Recordset
  rs.Open "SELECT * FROM cptablas WHERE tabla='31' and valor='S'  and (bloqueo='N' or Bloqueo is NULL) and codigo <>'..' ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
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
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpHorarioNew where (bloqueo='N' or Bloqueo is NULL) ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly    'CpHorarios
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
  
  If WcodUPR <> "" Then
     Tx_codupr = Wccosto
     For i = 0 To Cbo_Upr.ListCount - 1
        npos = InStr(Cbo_Upr.List(i), ":")
        nlen = Len(Cbo_Upr.List(i))
        dato = Trim(Right(Cbo_Upr.List(i), nlen - npos))
        If dato = Trim(WcodUPR) Then
           Cbo_Upr.ListIndex = i
           Exit For
        End If
     Next i
  End If
  
  Set D0 = New Recordset
  With D0
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
     .Fields.Append "MT", adChar, 2   '<========incluye MX
     .Fields.Append "NA", adChar, 2   '<========
     .Fields.Append "PA", adChar, 2   '<========
     .Open
  End With
  Set DataGrid1.DataSource = D0
  
  Set DD = New Recordset
  With DD
     .Fields.Append "Tarjeta", adChar, 5
     .Fields.Append "Horas", adDouble
     For i = dfi To dff
        ccampo = "D" + Trim(Str(i))
        .Fields.Append ccampo, adChar, 2
     Next i
     .Fields.Append "D", adChar, 2
     .Fields.Append "M", adChar, 2
     .Fields.Append "T", adChar, 2
     .Fields.Append "N", adChar, 2
     .Fields.Append "MT", adChar, 2   '<========
     .Fields.Append "NA", adChar, 2   '<========
     .Fields.Append "PA", adChar, 2   '<========
     .Open
  End With
  Set DataGrid2.DataSource = DD
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
            Txmes = Right(CboMes.List(j), nlen - nposicion)
            TxNombMes = Trim(Left(CboMes.List(j), nposicion - 1))
            unafecha = "01/" + Txmes + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)
            Call Cargadias       'primer/ultimo
            fecCierre = ultimo
            
            Set rs = New Recordset
            rs.Open "SELECT * FROM CpCronograma WHERE ano='" & genaño & "' and Mes='" & Txmes.Text & "'", Cnx, adOpenStatic, adLockReadOnly   'CpHorarios
            If rs.EOF Then
               MsgBox "No Existe Cronograma del MES... Verifique", vbExclamation, "[ " + genaño + " ]"
            Else
               If IsDate(rs!fecha) Then fecCierre = rs!fecha
               If fecCierre <= Date Then
                  Beep 500, 350
                  Me.Caption = "Programa de ASISTENCIALES... CERRADO.  - " & genaño & "   -   Sesión: " + wUsuNombre
               Else
                  Me.Caption = "Programa de ASISTENCIALES - " & genaño & "   -   Sesión: " + wUsuNombre
               End If
            End If
            rs.Close
         End If
      End If
      Exit Sub
   End If
End Sub

Private Sub Cbo_UPR_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Cbo_Upr.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbo_UPR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If IsEmpty(Cbo_Upr.Text) Then
     Else
        Tx_ccosto = "": TxUprCCosto = ""
        Txtarjeta = "": Txtrabajador = "": TxCarEst = "": TxDscCargo = ""
        TxcodUPR = "": Txclave = "":        TxCadHep = ""
        DataGrid1.Enabled = True
        j = Cbo_Upr.ListIndex
        If j <> -1 Then
           nposicion = InStr(Cbo_Upr.List(j), ":")
           nlen = Len(Cbo_Upr.List(j))
           Tx_codupr = Right(Cbo_Upr.List(j), nlen - nposicion)
           TxdscUpr = Left(Cbo_Upr.List(j), nposicion - 1)
           ceditor = "X"
           If Wccosto <> "" Then
              '-Verifica  derechos sobre UPR
              If InStr(wCadenaUpr, TxcodUPR) = 0 Then
                 MsgBox "Ud. no esta Autorizado...", vbExclamation, "Programaciòn UPR"
                 Exit Sub
              End If
              Tx_ccosto = Wccosto      'rs!cCosto
              '---wPwdOk = "NO.VA."
              ceditor = "X"
              Set rs = New Recordset
              rs.Open "SELECT * FROM CpCcostoUPR WHERE ccosto='" & Wccosto & "' and Codupr='" & Tx_codupr & "'", Cnx, adOpenStatic, adLockReadOnly
                 '- MsgBox "SELECT * FROM CpCcostoUPR WHERE ccosto='" & Wccosto & "' and Codupr='" & Tx_codupr & "'"
              If rs.EOF Then
              Else
                 If rs!Dni1 = wUsuario Then
                    ceditor = "S"
                 ElseIf rs!Dni2 = wUsuario Then
                    ceditor = "S"
                 ElseIf rs!dni3 = wUsuario Then
                    ceditor = "S"
                 End If
              End If
              rs.Close
              
              Frame4.Enabled = True
              Txclave.SetFocus
               '-  MsgBox (Wccosto + "[" + wCadenaUpr + "]" + ceditor + "[" + TxUprCCosto + "]")
              Exit Sub
           End If
           
           cHRcompl = "N"
           Set rs = New Recordset
           sql = "SELECT * FROM CpUndPrgRol where Codupr='" & Tx_codupr & "'"
           rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
           If Not rs.EOF Then
              cHRcompl = rs!hrcomplementaria
           End If
           rs.Close
           
           If cHRcompl <> "S" Then
              MsgBox "Error, Servicio NO aplica horas complemantarias..", vbCritical, "[ IMPORTANTE ]"
              rs.Close
              Exit Sub
           End If
           
           Set rs = New Recordset
           rs.Open "SELECT * FROM CpCcostoUPR WHERE Codupr='" & Tx_codupr & "'", Cnx, adOpenStatic, adLockReadOnly
           If rs.EOF Then
              MsgBox "Error, No existe..", vbExclamation, "Grupo UPR"
              rs.Close
              Exit Sub
           End If
           Tx_ccosto = rs!CCosto
           rs.Close
           Call cmdcarga_Click
        End If
        Exit Sub
     End If
   End If
End Sub


Private Sub Txclave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Txclave = UCase(Txclave)
      ' MsgBox "[" + Txclave + "=" + wPwdOk + "]"
      
      If Trim(Txclave) = Trim(wPwdOk) Then
         Call cmdcarga_Click
         Exit Sub
      Else
         MsgBox "Operador, No Identificado...", vbExclamation
         ' MsgBox "Operador, No Identificado...´(" + Txclave + "-" + UCase(wPwdOk) + ")", vbExclamation
          
         Txclave = ""
         Tx_codupr = "":         Cbo_Upr.Text = "":     Cbo_Upr.SetFocus
      End If
   End If
End Sub

Private Sub cmdcarga_Click()
   If Tx_codupr = "" Then Exit Sub
   Set D0 = New Recordset
   With D0
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
        'cCampo = rs.Fields(i).Name
        'If Left(cCampo, 2) = "C1" Then
        ' If cCampo = "C1047" Then
        '   ndato = IIf(IsNull(rs.Fields(i)), 0, rs.Fields(i))
        .Fields.Append ccampo, adChar, 2
     Next i
     .Fields.Append "MT", adChar, 2    '<=======
     .Fields.Append "NA", adChar, 2    '<=======
     .Fields.Append "PA", adChar, 2    '<=======
     .Open
   End With
   Set DataGrid1.DataSource = D0
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
   
   fecha1 = Format(primer, "yyyymmdd")
   fecha2 = Format(ultimo, "yyyymmdd")
   npuntero = 0
   '-DIRECTORIO Asistencial
   wLstHC = ""
   Set rs = New Recordset
   '-miembros de UPR c/derecho HorasComp
   n_son = 0
   sql = "SELECT * FROM Cpgeneral  WHERE  (baja is NULL) and codUpr='" & Tx_codupr & "'"
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   n_son = rs.RecordCount
   nvan = 0
   If n_son > 0 Then
      rs.MoveFirst
      Do While Not rs.EOF
         cTarjeta = rs!tarjeta
         nvan = nvan + 1
         If wLstHC = "" Then
               If n_son = nvan Then
                  wLstHC = " and tarjeta='" + cTarjeta + "' "
               Else
                  '1er servidor en lista
                  wLstHC = " and ( tarjeta IN ('" + cTarjeta + "','"
               End If
         Else
               If nvan < n_son Then
                  '-encadena item
                  wLstHC = wLstHC + cTarjeta + "','"
               Else
                  '-cierra ultimo item
                  wLstHC = wLstHC + cTarjeta + "'))"
               End If
         End If
         rs.MoveNext
      Loop
   End If
   rs.Close
   
   sql = "SELECT * FROM CpDirectorio  WHERE ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "' and codupr='" & Tx_codupr & "'  AND flag='2' " + wLstHC + " ORDER BY Tarjeta"
   '--   MsgBox sql
    
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   Do While Not rs.EOF
      '-  MsgBox "llego"
        
      cTarjeta = rs!tarjeta
      lok = True
      '-   MsgBox "paso " + cTarjeta + IIf(lok, "ok", "xx")
         
      If lok Then      '=============>
      cNombre = "":   cCarest = "":  cCargo = "":  cDscCargo = "":       cGrpGen = "":      cplh = "": ccadhep = ""
      
      If flgDir Then
         '-datos en Directorio
         cCarest = IIf(IsNull(rs!carEst), "", rs!carEst)
         cCargo = IIf(IsNull(rs!Cargo), "", rs!Cargo)
         cplh = IIf(IsNull(rs!PLH), "", rs!PLH)
      End If
      
      SQ1 = "SELECT * FROM Cpgeneral  WHERE  (baja is NULL) and tarjeta='" & cTarjeta & "'"
      Set rs_d = New Recordset
      rs_d.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      f_sigue = False
      If Not rs_d.EOF Then
         If rs_d!CodUpr = Tx_codupr Or (rs_d!CCosto = TxUprCCosto And nJefServ > 0) Then
            f_sigue = True
            cNombre = Left(Trim(rs_d!apelpat) + " " + Trim(rs_d!apelmat) + " " + Trim(rs_d!nombres), 25)
            cGrpGen = IIf(IsNull(rs_d!grpgen), "", rs_d!grpgen)
            ccadhep = IIf(IsNull(rs_d!CadHep), "", Trim(rs_d!CadHep))
            If Not flgDir Then
               cCarest = IIf(IsNull(rs_d!carEst), "", rs_d!carEst)
               cCargo = IIf(IsNull(rs_d!Cargo), "", rs_d!Cargo)
               cplh = IIf(IsNull(rs_d!PLH), "", rs_d!PLH)
            End If
         End If
      End If
      rs_d.Close
      If f_sigue Then
         npuntero = npuntero + 1
         Set rs_d = New Recordset
         rs_d.Open "SELECT * FROM Cpcargos WHERE codigo='" & cCargo & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs_d.EOF Then
            cDscCargo = Left(rs_d!descripcion, 15)
         End If
         rs_d.Close
         D0.AddNew
         D0!tarjeta = cTarjeta
         D0!nombre = cNombre
         D0!Cargo = cDscCargo
         D0!d = IIf(IsNull(rs!d), "", rs!d)
         D0!n = IIf(IsNull(rs!n), "", rs!n)
         D0!m = IIf(IsNull(rs!m), "", rs!m)
         D0!T = IIf(IsNull(rs!T), "", rs!T)
         D0!Horas = IIf(IsNull(rs!Horas), 0, rs!Horas)
         For i = Day(primer) To Day(ultimo)
            z = i - 1
            '- directorio rs.fields(Z+4)  ==>    dd.fields(Z+8)  ==> CpTmpDCosto(Z+6)
            '   Textd(z) = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
            'cCampo = rs.Fields(z + 4).Name
            D0.Fields(z + 8) = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
         Next i
         D0!mT = IIf(IsNull(rs!mT), "", rs!mT)   '<======
         D0!na = IIf(IsNull(rs!na), "", rs!na)   '<======
         D0!pa = IIf(IsNull(rs!pa), "", rs!pa)   '<======
         D0.Update
      End If
      End If    '<=============
      rs.MoveNext
   Loop
   rs.Close
   '-MsgBox "paso ¿?"
   
    npuntero = D0.RecordCount
    RowSelected = DataGrid1.Row + 1
    Lblpuntero = "Puntero: " + Trim(Str(RowSelected)) + " de " + Trim(Str(npuntero))
   
   Txclave = ""
   Frame4.Enabled = False
   DataGrid1.Refresh
End Sub
Private Sub DataGrid1_Click()
   If D0.RecordCount = 0 Then Exit Sub       'bloquea seleccion en dbgrid vacio
   Txtarjeta = DataGrid1.Columns(0)
   LblONOMASTICO.Caption = "Onomastico"
   LblONOMASTICO.Visible = False
   Call Txtarjeta_KeyPress(13)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   With DC
      If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
         .Sort = .Fields(ColIndex).[Name] & " Desc"
      Else
         .Sort = .Fields(ColIndex).[Name] & " Asc"
      End If
   End With
   Call Limpiacampos
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    RowSelected = DataGrid1.Row + 1
    npuntero = D0.RecordCount
    Lblpuntero = "Puntero: " + Trim(Str(RowSelected)) + " de " + Trim(Str(npuntero))
End Sub

Private Sub TxBUscar_KeyPress(KeyAscii As Integer)
  If nposicion = 0 Then Exit Sub
  If KeyAscii = 13 Then
     Dim cBuscar As String
     cpuntero = DataGrid1.Bookmark
     If Txbuscar = "" Then Exit Sub
     cBuscar = "tarjeta like '" & Txbuscar & "*'"
     D0.MoveFirst
     D0.Find (cBuscar)
     If D0.EOF Then
        Txbuscar.BackColor = &HFF&
     Else
        Txbuscar.BackColor = &H80000005
        DataGrid1.Refresh
        Txbuscar = ""
        DataGrid1_Click
     End If
  End If
End Sub
Private Sub Txbuscar1_KeyPress(KeyAscii As Integer)
   If nposicion = 0 Then Exit Sub
   If KeyAscii = 13 Then
     Dim cBuscar As String
     cpuntero = DataGrid1.Bookmark
     If Txbuscar1 = "" Then Exit Sub
     cBuscar = "Nombre like '" & Txbuscar1 & "*'"
     D0.MoveFirst
     D0.Find (cBuscar)
     If D0.EOF Then
        Txbuscar1.BackColor = &HFF&
     Else
        Txbuscar1.BackColor = &H80000005
        DataGrid1.Refresh
        Txbuscar1 = ""
        DataGrid1_Click
     End If
  End If
End Sub

Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Frame1.Caption = "HORAS COMPLEMENTARIAS del Servidor"
      For i = 1 To 31      '-dfi To dff
         z = i - 1
         Option1(z).Value = False
         Textd(z).Text = "": Textd(z).ForeColor = &H80000012
         Textp(z).Text = "":  Textp(z).ForeColor = &H80000012
         Textd(z).Visible = False: Textp(z).Visible = False
         Lblhr(z).Caption = ""
      Next
      
      LblONOMASTICO.Caption = "Onomastico"
      LblONOMASTICO.Visible = False
      If Trim(Txtarjeta) = "" Then
         Load Ayuda0HC
         Ayuda0HC.Tag = 6.21
         Ayuda0HC.Left = 6525: Ayuda0HC.Top = 0
         Ayuda0HC.Show 1
      Else
      
         If flgDir Then
            '-ubica en historial en directorio
            Set rs = New Recordset
            sq2 = "SELECT * FROM CpDirectorio WHERE ano='" + genaño + "' and mes='" + Txmes + "' and tarjeta = '" & Txtarjeta & "'  and codupr='" + Tx_codupr + "' and FLAG='2'"
            rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If rs.EOF Then
               MsgBox "No ubicado, complementarias UPR del servidor.", vbExclamation
               Txtarjeta = "": Txtrabajador = ""
               rs.Close
               Exit Sub
            End If
            rs.Close
            
          End If
      End If
      
      Set rs = New Recordset
      sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
      rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "No hay trabajadores para horas complementarias..", vbExclamation, "[IMPORTANTE]"
         Txtarjeta = "": Txtrabajador = ""
      Else
            If rs!tipoPla = "ASI" Then
               Txtrabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
               Txccosto = IIf(IsNull(rs!CCosto), "", rs!CCosto)
               TxcodUPR = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)
               TxDirector = IIf(IsNull(rs!Director), "N", rs!Director)
               TxCarEst = IIf(IsNull(rs!carEst), "", rs!carEst)
               TxPlH = IIf(IsNull(rs!PLH), "", rs!PLH)
               TxCargo = IIf(IsNull(rs!Cargo), "", rs!Cargo)
               Txactividad = IIf(IsNull(rs!actividad), "", rs!actividad)
               '-Cadena  a que pertenece Servidor
               TxCadHep = IIf(IsNull(rs!CadHep), "", Trim(rs!CadHep))
               If IsDate(rs!fecnac) Then Txfecnac = rs!fecnac
               TxDirector = IIf(IsNull(rs!Director), "N", Trim(rs!Director))
            Else
               MsgBox "No definido ASISTENCIAL..", vbExclamation, "[SERVIDOR]"
               Txtarjeta = "": Txtrabajador = ""
            End If
      End If
      rs.Close
      If Trim(Txtarjeta) <> "" Then
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and  tarjeta = '" & Trim(Txtarjeta) & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "No ubicado /sin derecho a H.C.", vbExclamation, "[SERVIDOR]"
            Txtarjeta = "": Txtrabajador = "": Txfecnac = ""
            TxCarEst = "": Txccosto = "": TxDirector = "": TxcodUPR = "": TxPlH = "": TxCargo = "": TxDscCargo = "": Txactividad = "": TxCadHep = ""
            LblONOMASTICO.Caption = "Onomastico"
            LblONOMASTICO.Visible = False
            rs.Close
            Exit Sub
         End If
         rs.Close
         If IsDate(Txfecnac) Then
            If Month(Txfecnac) = Val(Txmes) Then
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
         If Trim(Wccosto) <> "" Then
            If Txccosto <> Wccosto Then
               f_sigue = False
            End If
         End If
         If f_sigue Then
            If IsDate(Txfecnac) Then
               If Month(Txfecnac) = Val(Txmes) Then
                  '-Valida y Genera autorización ONO en Dia, jefe UPR debe programar 1/2 turno
                  Set rs_d = New Recordset
                  onofecha = DateValue(Right("0" + Trim(Day(Txfecnac)), 2) + "/" + Trim(Txmes) + "/" + Trim(Str(genaño)))
                  fechaB = Format(onofecha, "yyyymmdd")
                  rs_d.Open "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta='" & Txtarjeta & "' AND DiaHora='D' AND fechaini<='" & fechaB & "'AND fechaFin>='" & fechaB & "' ", Cnx, adOpenStatic, adLockReadOnly
                  If rs_d.EOF Then
                     rs_d.Close
                     rs_d.Open "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta='" & Txtarjeta & "' AND codigo='ONO' AND DiaHora='D' AND fechaini='" & fechaB & "'", Cnx, adOpenStatic, adLockOptimistic
                     If rs_d.EOF Then
                        rs_d.AddNew
                        rs_d!tarjeta = Txtarjeta
                        rs_d!Codigo = "ONO"
                        rs_d!diahora = "D"
                        rs_d!fechaIni = onofecha
                        rs_d!fechaFin = onofecha
                        rs_d!documento = "Autorizacion Automatica"
                        rs_d!usuario = wUsuario      'Firma Usuario
                        rs_d!fecha = Date
                        rs_d.Update
                     End If
                  End If
                  rs_d.Close
               End If
            End If
         Else
            MsgBox "No pertenece a Unidad Organizacional..! " + Wccosto, vbExclamation, "[SERVIDOR]"
            Txtarjeta = "": Txtrabajador = "": Txfecnac = ""
            TxCarEst = "": Txccosto = "": TxcodUPR = "": TxPlH = "":  TxCargo = "":  TxDscCargo = "": Txactividad = "": TxCadHep = ""
            LblONOMASTICO.Caption = "Onomastico"
            LblONOMASTICO.Visible = False
            Exit Sub
         End If
        Call cmdcargaDIR
        Call cmdcargHRC_Click
       '- MsgBox "dir[" + cCadTurDIR + "]"
       '- MsgBox "hrC[" + cCadTurHRC + "]"
        
      End If
   End If
End Sub

Private Sub cmdcargaDIR()
   If Tx_codupr = "" Then Exit Sub
   dfi = Day(primer)
   dff = Day(ultimo)
   Set DD = New Recordset
   '-Inicia cadena DIRECTORIO
   cCadTurDIR = ""
   With DD
     .Fields.Append "Tarjeta", adChar, 5
     .Fields.Append "Horas", adDouble
     For i = dfi To dff
        ccampo = "D" + Trim(Str(i))
        'cCampo = rs.Fields(i).Name
        'If Left(cCampo, 2) = "C1" Then
        ' If cCampo = "C1047" Then
        '   ndato = IIf(IsNull(rs.Fields(i)), 0, rs.Fields(i))
        .Fields.Append ccampo, adChar, 2
        cCadTurDIR = cCadTurDIR + "oooo_"
     Next i
     .Fields.Append "D", adChar, 2
     .Fields.Append "M", adChar, 2
     .Fields.Append "T", adChar, 2
     .Fields.Append "N", adChar, 2
     .Fields.Append "MT", adChar, 2    '<=======
     .Fields.Append "NA", adChar, 2    '<=======
     .Fields.Append "PA", adChar, 2    '<=======
     .Open
   End With
   Set DataGrid2.DataSource = DD
   For i = 0 To DataGrid2.Columns.Count - 1
     With DataGrid2.Columns(i)
        Select Case i
         Case 0
            .Caption = "Tarjeta"
            .Width = Me.TextWidth(String(7, "S"))
            .Visible = False
         Case 1
            '.Caption =
            .Width = Me.TextWidth(String(5, "S"))
         Case i < 12
            .Width = Me.TextWidth(String(3, "S"))
         Case Else
             .Width = Me.TextWidth(String(4, "S"))
        End Select
        '.NumberFormat = "##,##0.0"
        .Locked = True
     End With
   Next
   
   '-DIRECTORIO Asistencial
   sql = "SELECT * FROM CpDirectorio  WHERE ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "' and codupr='" & Tx_codupr & "'  AND flag='0'  and Tarjeta='" + Trim(Txtarjeta) + "'"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      '-  MsgBox "llego"
         DD.AddNew
         DD!tarjeta = Txtarjeta
         DD!Horas = IIf(IsNull(rs!Horas), 0, rs!Horas)
         nlen = Len(cCadTurDIR)
         For i = Day(primer) To Day(ultimo)
            z = i - 1
            '- directorio rs.fields(Z+4)  ==>    dd.fields(Z+8)  ==> CpTmpDCosto(Z+6)
            '   Textd(z) = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
            'cCampo = rs.Fields(z + 4).Name
            DD.Fields(z + 2) = IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4))
            '         1         2         3         4
            '1234567890123456789012345678901234567890
            '0000_0000_0000_0000_0000_     =>   0000_0000_MT00_0000_0000_
            '1    2    3    4    5         left(cad,5): [5]   +xxoo+   right(cad,25-(5+4) :[16]
            'dia i:2             (i-1)*5     nlen-((i-1)*5 +4)
            cvalor = "oooo"
            If DD.Fields(z + 2) = "MT" Or DD.Fields(z + 2) = "MX" Then
               cvalor = cvalMt
            ElseIf Left(DD.Fields(z + 2), 1) = "D" Then
                   cvalor = cvalD
            ElseIf Left(DD.Fields(z + 2), 1) = "N" Then
                   cvalor = cvalN
            ElseIf Left(DD.Fields(z + 2), 1) = "M" Then
                   cvalor = cvalM
            ElseIf Left(DD.Fields(z + 2), 1) = "T" Then
                   cvalor = cvalT
            ElseIf Left(DD.Fields(z + 2), 1) = "H" Then
                   cvalor = cvalH
            ElseIf Left(DD.Fields(z + 2), 1) = "C" Then
                   cvalor = cvalC
            ElseIf Left(DD.Fields(z + 2), 1) = "P" Then
                   cvalor = cvalT
            End If
            
            '-cCadTurDIR = Left(cCadTurDIR, (i - 1) * 5) + (IIf(IsNull(rs.Fields(z + 4)), "oo", Left(Trim(rs.Fields(z + 4)) + "oo", 2))) + Right(cCadTurDIR, nlen - (i - 1) * 5 - 4)
            cCadTurDIR = Left(cCadTurDIR, (i - 1) * 5) + cvalor + Right(cCadTurDIR, nlen - (i - 1) * 5 - 4)
         Next i
         DD!d = IIf(IsNull(rs!d), "", rs!d)
         DD!n = IIf(IsNull(rs!n), "", rs!n)
         DD!m = IIf(IsNull(rs!m), "", rs!m)
         DD!T = IIf(IsNull(rs!T), "", rs!T)
         DD!mT = IIf(IsNull(rs!mT), "", rs!mT)   '<======
         DD!na = IIf(IsNull(rs!na), "", rs!na)   '<======
         DD!pa = IIf(IsNull(rs!pa), "", rs!pa)   '<======
         DD.Update
   End If
   rs.Close
   '-MsgBox "paso ¿?"
   DataGrid2.Refresh
End Sub

Private Sub cmdcargHRC_Click()
   TxHoras = "": TxD = "": TxN = "": TxM = "": Txt = "": TxMT = "": TxNa = "": TxPA = ""
   If Txtarjeta = "" Then Exit Sub
   '-HRCOMPLENT. Actual   check AudDirect a partir del primer<date
   cDirectHoy = ""
   cCadTurHRC = ""
   For i = dfi To dff
       z = i - 1
       Option1(z).Value = False
       Textd(z).Visible = True: Textp(z).Visible = True
       Textd(z).Text = "": Textd(z).ForeColor = &H80000012
       Textp(z).Text = "":  Textp(z).ForeColor = &H80000012
       Lblhr(z).Caption = ""
       cCadTurHRC = cCadTurHRC + "oooo_"
   Next
   fecha1 = Format(primer, "yyyymmdd")
   fecha2 = Format(ultimo, "yyyymmdd")
   '-carga  COMPLEMENTARIAS del MES
   sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Trim(Txtarjeta) & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'  and FLAG='2'"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   nlen = Len(cCadTurHRC)
   If Not rs.EOF Then
      For i = Day(primer) To Day(ultimo)
         z = i - 1
          Textd(z) = IIf(IsNull(rs.Fields(z + 4)), "", Trim(rs.Fields(z + 4)))
         cDirectHoy = cDirectHoy + IIf(IsNull(Textd(z)), "==", Left(Textd(z) + "==", 2))
         
            cvalor = "oooo"
            If Textd(z) = "MT" Or Textd(z) = "MX" Then
               cvalor = cvalMt
            ElseIf Left(Textd(z), 1) = "D" Then
                   cvalor = cvalD
            ElseIf Left(Textd(z), 1) = "N" Then
                   cvalor = cvalN
            ElseIf Left(Textd(z), 1) = "M" Then
                   cvalor = cvalM
            ElseIf Left(Textd(z), 1) = "T" Then
                   cvalor = cvalT
            ElseIf Left(Textd(z), 1) = "H" Then
                   cvalor = cvalH
            ElseIf Left(Textd(z), 1) = "C" Then
                   cvalor = cvalC
            ElseIf Left(Textd(z), 1) = "P" Then
                   cvalor = cvalT
            End If
         
         '         1
         '1234567890123456
         '0000_0000_0000_0000_0000_     =>   0000_0000_xxoo_00_00_
         '1  2  3  4  5       left(cad,6)+MT+right(15-(6+2)
         'dia i:3             (i-1)*3        nlen- ((i-1)*3 +2)
         
         
         '-cCadTurHRC = Left(cCadTurHRC, (i - 1) * 3) + (IIf(IsNull(rs.Fields(z + 4)), "oo", Left(Trim(rs.Fields(z + 4)) + "oo", 2))) + Right(cCadTurDIR, nlen - (i - 1) * 3 - 2)
         cCadTurHRC = Left(cCadTurHRC, (i - 1) * 5) + cvalor + Right(cCadTurDIR, nlen - (i - 1) * 5 - 4)
         
         If InStr("MT_MX_NA_PA", Textd(z)) > 0 And Len(Textd(z)) = 2 Then
            Select Case Textd(z)
               Case "MT", "MX"
                 TxMT = Val(TxMT) + 1
               Case "NA"
                 TxNa = Val(TxNa) + 1
               Case "PA"
                 TxPA = Val(TxPA) + 1
             End Select
         Else
            Select Case Trim(Textd(z).Text)
               Case "A"
                 'Administ -> Asistencial
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
   nerror = 0
   sql = "SELECT * FROM CpHorCOMP WHERE tarjeta='" & Trim(Txtarjeta) & "' AND fecha >='" & fecha1 & "' AND fecha<='" & fecha2 & "' ORDER BY Fecha,tarjeta"
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
            'CpHorarios
            rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs_d.EOF Then
               cvalor = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
               TxHoras = Val(TxHoras) + Val(Left(cvalor, 2)) + (Val(Right(cvalor, 2)) / 60)
               If Val(cvalor) > 0 Then
                  Lblhr(z).Caption = Left(cvalor, 2) + ":" + Right(cvalor, 2)
                  If Trim(Textd(z)) = "" Then
                     nerror = nerror + 1
                  End If
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
   If nerror > 0 Then
      MsgBox "Existe " + Trim(Str(nerror)) + " Dia(s) Inconsistente(s).. ", vbExclamation
      Exit Sub
   End If
End Sub

Private Sub CmdLimpia_Click()
    Txtarjeta = "": Txtrabajador = "": Txfecnac = ""
    TxCarEst = "": Txccosto = "": TxcodUPR = "": TxPlH = "": TxCargo = "": TxDscCargo = "": Txactividad = "": TxCadHep = ""
    TxHoras = "": TxD = "": TxN = "": TxM = "": Txt = "": TxMT = "": TxNa = "": TxPA = ""
    CboTurDir.Text = "": TxTurno = ""
    CboHorario.Text = "": TxHorario = ""
    For i = 1 To 31        'dfi To dff
       z = i - 1
       Option1(z).Value = False
       Textd(z).Visible = True: Textp(z).Visible = True
       Textd(z).Text = "": Textd(z).ForeColor = &H80000012  '&HFF0000
       Textp(z).Text = "":  Textp(z).ForeColor = &H80000012 '&HFF0000
       Lblhr(z).Caption = ""
    Next
    CboRegPlh.Text = "": TxRegPlh = ""
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
      Case 1
   End Select
   CboTurDir.Enabled = True
   TxTurno = ""
   cTurnoAnt = ""
   CboHorario.Enabled = True
   TxHorario = ""
   dfi = Day(primer)
   dff = Day(ultimo)
   npuntDia = 0     'posc. Dia seleccionado
   For i = dfi To dff
     z = i - 1
     If Option1(z).Value = True Then
        '--- Dia seleccionado ---
        If InStr("L_V_R", Trim(Textd(z).Text)) > 0 And Trim(Textd(z).Text) <> "" Then
           '-No Editable
           Textd(z).ForeColor = &H0&
           Option1(z).Value = False
           Textd(z).Locked = True:                  Textp(z).Locked = True
           Exit Sub
        End If
        npuntDia = z
        cTurnoAnt = Textd(z).Text            'Lock V/R/L => Autorizacion
        TxTurno = Textd(z).Text
        TxHorario.Text = Textp(z).Text
        Textd(z).ForeColor = &HFF0000
        Textd(z).Locked = False
        If InStr("V_R_L", Trim(Textd(z).Text)) > 0 And Trim(Textd(z).Text) <> "" Then 'Lock V/R/L => Autorizacion
           MsgBox "No, Puede modificar autorizacion.. ", vbExclamation
           Textd(Index).ForeColor = &H0&
           Option1(Index).Value = False
           Textd(Index).Locked = True:                  Textp(Index).Locked = True
           Option1(z).SetFocus
           Exit Sub
        End If
        Textd(z).SetFocus
        Exit Sub
     End If
   Next
End Sub
Private Sub Textd_LostFocus(Index As Integer)
      Textd(Index) = UCase(Textd(Index))
      If Trim(Textd(Index).Text) <> "" And Trim(Textd(Index).Text) <> "*" Then
         If InStr("R_L_V", Trim(Textd(Index).Text)) > 0 Then
            MsgBox "Turno seleccionado, No Autorizado... ", vbCritical, "[ Oficina de Personal ]"
            Textd(Index).Text = cTurnoAnt            'Lock V/R/L => Autorizacion
            Textd(Index).ForeColor = &H0&
            Option1(Index).Value = False
            Textd(Index).Locked = True
            Textp(Index).Text = TxHorario:                  Textp(Index).Locked = True
            Option1(Index).SetFocus
            Exit Sub
         End If
         '-Verifica si Turno Aplica para UPR
         Set rs = New Recordset
         rs.Open "SELECT * FROM CpcCostoHor WHERE ccosto='" & Tx_ccosto & "' and  turno='" & Trim(Textd(Index).Text) & "'", Cnx, adOpenStatic, adLockReadOnly
         If rs.EOF Then
            MsgBox "Turno NO se programa en UPR... ", vbExclamation
            Textd(Index).Text = cTurnoAnt            'Lock V/R/L => Autorizacion
            Textd(Index).ForeColor = &H0&
            Option1(Index).Value = False
            Textd(Index).Locked = True
            Textp(Index).Text = TxHorario:                  Textp(Index).Locked = True
            Option1(Index).SetFocus
            rs.Close
            Exit Sub
         End If
         rs.Close
         '-Asigna cadena al turno
            cvalor = "oooo"
            If Textd(Index) = "MT" Or Textd(Index) = "MX" Then
               cvalor = cvalMt
            ElseIf Left(Textd(Index), 1) = "D" Then
                   cvalor = cvalD
            ElseIf Left(Textd(Index), 1) = "N" Then
                   cvalor = cvalN
            ElseIf Left(Textd(Index), 1) = "M" Then
                   cvalor = cvalM
            ElseIf Left(Textd(Index), 1) = "T" Then
                   cvalor = cvalT
            ElseIf Left(Textd(Index), 1) = "H" Then
                   cvalor = cvalH
            ElseIf Left(Textd(Index), 1) = "C" Then
                   cvalor = cvalC
            ElseIf Left(Textd(Index), 1) = "P" Then
                   cvalor = cvalT
            End If
         
         '-Verifica turnos del Directorio / Complementarias
         i = Index + 1
         cTurnoDIrSig = "oooo": cTurnoDirAnt = "oooo": cTurnoDAnt = "oooo": cTurnoDSig = "oooo"
         cTurnoDir = Mid(cCadTurDIR, (i - 1) * 5 + 1, 4)
         If i > Day(primer) Then
            If i <= Day(ultimo) Then
               cTurnoDirAnt = Mid(cCadTurDIR, (i - 2) * 5 + 1, 4)
               cTurnoDAnt = Mid(cCadTurHRC, (i - 2) * 5 + 1, 4)
            End If
            If i < Day(ultimo) Then
               cTurnoDIrSig = Mid(cCadTurDIR, (i) * 5 + 1, 4)
               cTurnoDSig = Mid(cCadTurHRC, (i) * 5 + 1, 4)
            End If
         Else
            cTurnoDIrSig = Mid(cCadTurDIR, (i) * 5 + 1, 4)
            cTurnoDSig = Mid(cCadTurHRC, (i) * 5 + 1, 4)
         End If
          '--  MsgBox "Dir ant-dia-sig:" + cTurnoDirAnt + "-" + cTurnoDir + "-" + cTurnoDIrSig
         '--Dia complementarias ANT/SIG
          '-- MsgBox "HrC ant-dia-sig:" + cTurnoDAnt + " - " + Left(Textd(Index).Text, 1) + " - " + cTurnoDSig
           
         fsigue = True
         c_a = ""
         
         '-VALIDA turno vs Directorio / complementarias
         If cvalor = cTurnoDir And cTurnoDir <> "oooo" Then
            c_a = " (turno Repetido)"
            fsigue = False
         ElseIf Right(cvalor, 2) = "xx" And Right(cTurnoDir, 2) = "xx" Then
                c_a = " (Cruce noche)"
                fsigue = False
         ElseIf Right(cTurnoDirAnt, 2) = "xx" And InStr("xo_ox_xx", Left(cvalor, 2)) > 0 Then
                 c_a = " (Noche ant.)"
                 fsigue = False
         ElseIf InStr("xo_xx", Left(cvalor, 2)) > 0 And InStr("xo_xx", Left(cTurnoDir, 2)) > 0 Then
                c_a = " (Cruce dia)"
                fsigue = False
         ElseIf Mid(cvalor, 2, 1) = Mid(cTurnoDir, 2, 1) And Mid(cTurnoDir, 2, 1) <> "o" Then
                c_a = " (Cruce dia)"
                fsigue = False
         ElseIf Right(cvalor, 2) = "xx" And InStr("xx_xo_ox", Left(cTurnoDIrSig, 2)) > 0 Then
                c_a = " (hay rol día sig.)"
                fsigue = False
                
         ElseIf Right(cTurnoDAnt, 2) = "xx" And InStr("xo_ox_xx", Left(cvalor, 2)) > 0 Then
                 c_a = " (HrC Noche ant.)"
                 fsigue = False
         ElseIf Right(cvalor, 2) = "xx" And InStr("xo_ox_xx", Left(cTurnoDSig, 2)) > 0 Then
                c_a = " (hay HrC día sig.)"
                fsigue = False
         End If
         
         If fsigue Then
         Else
            MsgBox "Turno " + Textd(Index).Text + " no aplica " + c_a + "... ", vbCritical, "[VERIFIQUE]"
            Textd(Index).Text = cTurnoAnt            'Lock V/R/L => Autorizacion
            Textd(Index).ForeColor = &H0&
            Option1(Index).Value = False
            Textd(Index).Locked = True
            Textp(Index).Text = TxHorario:                  Textp(Index).Locked = True
            Option1(Index).SetFocus
            Exit Sub
         End If
         
         '-NOMBRADOS Medicos D/M/T/N/MT-MX 1,4;  Enf-tec
         '-CAS       Medicos D/M/T/N/MT-MX 1,4;  Enf-tec
         TxTurno = Textd(Index).Text
         If Trim(UCase(Textd(Index).Text)) = "V" Or Trim(UCase(Textd(Index).Text)) = "L" Or Trim(UCase(Textd(Index).Text)) = "*" Then
            Textp(Index).Text = "O"
            TxHorario = "O"
         Else
            If Trim(UCase(Textd(Index).Text)) <> "A" Then
               Set rs_d = New Recordset
               sq_l = "SELECT * FROM CpcCostoHor WHERE Ccosto='" & Txccosto & "' And turno='" & Trim(Textd(Index).Text) & "'"
               '--  MsgBox sq_l
               rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               nson = 0
               If Not rs_d.EOF Then
                  Do While Not rs_d.EOF
                     nson = nson + 1
                     Textp(Index).Text = rs_d!Horario
                     TxHorario = rs_d!Horario
                     '--    MsgBox "horario ->" + TxHorario
                     rs_d.MoveNext
                  Loop
                  If nson > 1 Then
                     Textp(Index).Text = ""
                     TxHorario = ""
                  End If
               Else
                  Textp(Index).Text = ""
               End If
               rs_d.Close
               '--   MsgBox "Horario[" + TxHorario + "]" '<==============
            End If
         End If
         '- Calcula acumulados
         nposicion = Index
         Call Recalcula
         Index = nposicion
         
         '-seleccion  item cboturdir
         If Trim(TxTurno) <> "" Then
            For i = 0 To CboTurDir.ListCount - 1
               npos = InStr(CboTurDir.List(i), ":")
               nlen = Len(CboTurDir.List(i))
               dato = Trim(Right(CboTurDir.List(i), nlen - npos))
               If dato = TxTurno Then
                  CboTurDir.ListIndex = i
                  Exit For
               End If
            Next i
         End If
    
         CboTurDir_KeyPress (13)
         Call Recalcula
         
      Else   'Solo Anula turno ANTERIOR
      
         If InStr("MT_MX_NA_PA", cTurnoAnt) > 0 And Len(cTurnoAnt) = 2 Then
            Select Case cTurnoAnt
            Case "MT", "MX"
               TxMT = Val(TxMT) - 1
            Case "NA"
               TxNa = Val(TxNa) - 1
            Case "PA"
               TxPA = Val(TxPA) - 1
            End Select
         Else
      
            Select Case Trim(cTurnoAnt)
            Case "A"
              'administ -> Asistencial
              TxMT = Val(TxMT) - 1
            Case Else
               If Left(cTurnoAnt, 1) = "M" Then
                  TxM = Val(TxM) - 1
               ElseIf Left(cTurnoAnt, 1) = "T" Then
                      Txt = Val(Txt) - 1
               ElseIf Left(cTurnoAnt, 1) = "D" Then
                      TxD = Val(TxD) - 1
               ElseIf Left(cTurnoAnt, 1) = "N" Then
                      TxN = Val(TxN) - 1
               ElseIf Left(cTurnoAnt, 1) = "H" Then
                      TxM = Val(TxM) - 1
                      TxN = Val(TxN) - 1
               ElseIf Left(cTurnoAnt, 1) = "C" Then
                      Txt = Val(Txt) - 1
                      TxN = Val(TxN) - 1
               ElseIf Left(cTurnoAnt, 1) = "P" Then
                      TxD = Val(TxD) - 1
                      TxN = Val(TxN) - 1
               End If
            End Select
         End If
         TxHoras = Val(TxHoras) - (Val(Left(Trim(Lblhr(Index).Caption), 2)) + (Val(Right(Trim(Lblhr(Index).Caption), 2)) / 60))
         
         Textd(Index).ForeColor = &H0&: Textp(Index).ForeColor = &H0&
         Option1(Index).Value = False
         Textd(Index).Locked = True:                  Textp(Index).Locked = True
         Textd(Index).Text = "":     Textp(Index).Text = "": Lblhr(Index).Caption = ""
         TxTurno = "":   CboTurDir.Text = "":   CboTurDir.Enabled = False
         TxHorario = "":   CboHorario.Text = "":    CboHorario.Enabled = False
         Call Recalcula
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
           TxTurno = Right(CboTurDir.List(j), nlen - nposicion)
           '--- MsgBox "turno[" + TxTurno + "]"
           
           dfi = Day(primer)
           dff = Day(ultimo)
           For i = dfi To dff
              z = i - 1
              If Option1(z).Value = True Then
                 Textd(z).Text = TxTurno
                 Textd(z).ForeColor = &HFF0000
                 Textp(z).SetFocus
                 If Trim(TxHorario) <> "" Then    ' "O"
                    Set rs_d = New Recordset
                    sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & TxHorario & "'"
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
     If TxHorario <> "  " Then
        '-NOMBRADOS Medicos D, N, T, M; Enf-tec D,N,T,J,M ,MT-MX
        '-CAS       Medicos D,N,T,M;     Enf-tec D,N,T,M    ,MT-MX
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
           TxHorario = Right(CboHorario.List(j), nlen - nposicion)
           '- TODOS V-O,*-O
           '-NOMBRADOS Medicos D,N,T,M; Enf-tec 'MT-mx
           '-CAS       Medicos D,N,T,M;      Enf-tec    'MT-MX
           If (Trim(UCase(TxTurno)) = "V" Or Trim(UCase(TxTurno.Text)) = "L" Or Trim(UCase(TxTurno)) = "*") And InStr("O", Trim(UCase(TxHorario))) = 0 Then
              Sw = False
           ElseIf Trim(UCase(TxTurno)) = "A" Then
                  Sw = False
           ElseIf InStr("1DMT4NHCP", Left(UCase(TxTurno), 1)) > 0 Then
                  Set rs_d = New Recordset
                  sq_l = "SELECT * FROM CpCCostoHor WHERE CCosto='" & Trim(Tx_ccosto) + "' and turno='" + TxTurno & "' And  horario='" + TxHorario + "'"
                  rs_d.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  If rs_d.EOF Then
                     Sw = False
                  End If
                  rs_d.Close
                  
           End If
           If Sw Then
           Else
              Beep 500, 350
              MsgBox "HORARIO no Corresponde...  " + Tx_ccosto + "/" + TxTurno + ":" + TxHorario, vbExclamation
              Exit Sub
           End If
           dfi = Day(primer)
           dff = Day(ultimo)
           TxHoras = ""     'calcula Horas Prog
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
                 TxTurno = "":   CboTurDir.Text = "":   CboTurDir.Enabled = False
                 'cmdGraba.SetFocus
              End If
              If Textp(z).Text <> "" Then
                 '--- MsgBox Str(i) + "-" + Textp(z).Text
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
                 '--- MsgBox sq_l + "-" + cvalor
                 
              End If
           Next
           If Val(TxHoras) > 150 Then
              MsgBox "Excede, máximo permitido..", vbExclamation, "Horas del Programa"
              '-Retira reg error
                 
              If InStr("MT_MX_NA_PA", Textd(z)) > 0 And Len(Textd(z)) = 2 Then
                 Select Case Textd(z)
                    Case "MT", "MX"
                       TxMT = Val(TxMT) - 1
                    Case "NA"
                       TxNa = Val(TxNa) - 1
                    Case "PA"
                       TxPA = Val(TxPA) - 1
                 End Select
              Else
          
                  Select Case Trim(Textd(z))
                    Case "A"
                      'administrativo -> Asistencial
                       TxMT = Val(TxMT) - 1
                    Case Else
                       If Left(Textd(z), 1) = "M" Then
                          TxM = Val(TxM) - 1
                       ElseIf Left(Textd(z), 1) = "T" Then
                              Txt = Val(Txt) - 1
                       ElseIf Left(Textd(z), 1) = "D" Then
                              TxD = Val(TxD) - 1
                       ElseIf Left(Textd(z), 1) = "N" Then
                              TxN = Val(TxN) - 1
                       ElseIf Left(Textd(z), 1) = "H" Then
                              TxM = Val(TxM) - 1
                              TxN = Val(TxN) - 1
                       ElseIf Left(Textd(z), 1) = "C" Then
                              Txt = Val(Txt) - 1
                              TxN = Val(TxN) - 1
                       ElseIf Left(Textd(z), 1) = "P" Then
                              TxD = Val(TxD) - 1
                              TxN = Val(TxN) - 1
                       End If
                  End Select
              End If
                 
              TxHoras = Val(TxHoras) - Val(Left(Trim(Lblhr(ndia).Caption), 2)) - (Val(Right(Trim(Lblhr(ndia).Caption), 2)) / 100)
              Textd(ndia).Text = "": Textp(ndia).Text = ""
              Lblhr(ndia).Caption = ""
              Option1(ndia).Value = False
              Textd(ndia).Locked = True:                  Textp(ndia).Locked = True
           End If
           cmdGraba.SetFocus
        End If
        Exit Sub
     End If
   End If
End Sub
Private Sub Cbohorario_LostFocus()
   If Trim(CboHorario.Text) <> "" Then
      Cbohorario_KeyPress (13)
   End If
End Sub

Private Sub Limpiacampos()
   Txtarjeta = "": Txtrabajador = "": Txfecnac = ""
   TxCarEst = "": Txccosto = "": TxcodUPR = "": TxPlH = "": TxCargo = "": TxDscCargo = "": Txactividad = "": TxCadHep = ""
   TxTurno = "": CboTurDir.Text = ""
   TxHoras = "":  TxHorario = "": CboHorario.Text = ""
   TxD = "": TxN = "": TxM = "": Txt = "": TxMT = "": TxNa = "": TxPA = ""
   CboRegPlh.Text = "": TxRegPlh = ""
   For i = 1 To 31
      z = i - 1
      Option1(z).Value = False
      Textd(z).Text = "":      Textd(z).ForeColor = &H80000008
      Textd(z).Visible = True
      Textp(z).Text = "":      Textp(z).ForeColor = &H80000008
      Lblhr(z).Caption = "":   Lblhr(z).ForeColor = &H80000008
   Next
   
End Sub
Private Sub Cmdchk_Click()
   '-Corrige Generales Cargo,Carest,PLH; Fecnac, Grupo, CadHep
   'If fecCierre <= Date Then
   '   MsgBox "Programa CERRADO...  ", vbExclamation
   '   Exit Sub
   'End If
   If IsNull(Txmes) Then
      MsgBox "Defina PERIODO...  ", vbExclamation
      Exit Sub
   End If
   If IsNull(Txtarjeta) Then
      MsgBox "Defina trabajador...  ", vbExclamation
      Exit Sub
   End If
   If wrights = "CON" Or Wccosto <> "" Then
      MsgBox "Operador No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   Call Recalcula
   Set rs = New Recordset
   sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
   rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If rs.EOF Then
   Else
      TxCarEst = IIf(IsNull(rs!carEst), "", rs!carEst)
      TxCargo = IIf(IsNull(rs!Cargo), "", rs!Cargo)
      Txactividad = IIf(IsNull(rs!actividad), "", rs!actividad)
      TxPlH = IIf(IsNull(rs!PLH), "", rs!PLH)
      TxCadHep = IIf(IsNull(rs!CadHep), "", Trim(rs!CadHep))
      If IsDate(rs!fecnac) Then Txfecnac = rs!fecnac
   End If
   rs.Close
   If IsDate(Txfecnac) Then
      If Month(Txfecnac) = Val(Txmes) Then
         LblONOMASTICO.Visible = True
         LblONOMASTICO.Caption = "Onomastico el " + Str(Day(Txfecnac))
      End If
   End If
   Set rs = New Recordset
   rs.Open "SELECT * FROM Cpcargos WHERE  codigo='" & TxCargo & "'", Cnx, adOpenStatic, adLockReadOnly
   If Not rs.EOF Then
      TxDscCargo = Left(rs!descripcion, 15)
   End If
   rs.Close
   
   On Error GoTo final
   Cnx.BeginTrans
   '-Actualiza DIRECTORIO : Plh, Carest, Cargo e Inconsistencia Turno-Horario
   Sw = False
   If Txtarjeta <> "" Then
      sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "' and FLAG='2'"
      Set rs = New Recordset
      rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      If Not rs.EOF Then
         '--D0!cargo = Left(TxDscCargo, 15)
         '--D0. Update
         
         rs!PLH = IIf(IsNull(TxPlH), "", TxPlH.Text)     'Regimen laboral
         rs!carEst = IIf(IsNull(TxCarEst), "", TxCarEst)
         rs!Cargo = IIf(IsNull(TxCargo), "", TxCargo)
         rs!actividad = IIf(IsNull(Txactividad), "", Txactividad)
         '-SABER si una columna existe en una tabla
         SQ1 = "SELECT * FROM  INFORMATION_SCHEMA.COLUMNS AS c1  WHERE  c1.column_name = 'cadhep'  AND c1.table_name = 'cpdirectorio'"
         Set rs_2 = New Recordset
         rs_2.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If Not rs_2.EOF Then
            rs!CadHep = IIf(IsNull(TxCadHep), "", Trim(TxCadHep))
         End If
         rs.Update
         rs_2.Close
         
         'Ubica y corrige inconsistencias Horario-turno
         fecha1 = Format(primer, "yyyymmdd")
         fecha2 = Format(ultimo, "yyyymmdd")
         SQ1 = "SELECT * FROM CpHorCOMP WHERE tarjeta='" & Txtarjeta & "' AND fecha >='" & fecha1 & "' AND fecha<='" & fecha2 & "' ORDER BY Fecha,tarjeta"
         Set rs_d = New Recordset
         rs_d.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs_d.EOF Then
            rs_d.MoveFirst
            Do While Not rs_d.EOF
               i = Day(rs_d!fecha)
               z = i - 1
               Chor = IIf(IsNull(rs_d!Horario), "", rs_d!Horario)
               If Chor <> "" Then
                  cvalor = Lblhr(z)
                  If Trim(cvalor) <> "" And Trim(Textd(z)) = "" Then
                     sq_l = "SELECT * FROM CpcCostoHor WHERE ccosto='" & Txccosto & "' AND horario='" & Chor & "'"
                     Set rs_d1 = New Recordset
                     rs_d1.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                     If Not rs_d1.EOF Then
                        Textd(z) = Trim(rs_d1!turno)
                        rs.Fields(z + 4) = Trim(rs_d1!turno)
                        rs.Update
                        
                        D0.Fields(z + 8) = IIf(IsNull(Textd(z)), "", Textd(z))
                        D0.Update
                        
                     End If
                     rs_d1.Close
                  End If
               End If
               rs_d.MoveNext
            Loop
         End If
         rs_d.Close
         
      End If
      rs.Close
   End If
   Cnx.CommitTrans
   Call Limpiacampos
   Frame1.Caption = "Programa HORAS COMPLEMENRARIAS"
   GoTo Final1
final:
   MsgBox Err.Description & " Error: " & Err.Number & vbCr
   Cnx.RollbackTrans
Final1:
DataGrid1.SetFocus
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
      MsgBox "Operador de Consultas...  ", vbExclamation
      Exit Sub
   End If
   If IsNull(Txmes) Then
      MsgBox "Defina PERIODO...  ", vbExclamation
      Exit Sub
   End If
   If IsNull(TxPlH) Then
      MsgBox "Verifique Relacion Laboral...  ", vbExclamation
      Exit Sub
   End If
   If fecCierre <= Date Then
      Beep 500, 350
      MsgBox "Programa CERRADO...  ", vbExcritical
      Exit Sub
   End If
   If Trim(Txtarjeta) = "" Or Trim(Txtrabajador) = "" Then
      MsgBox "Defina trabajador...  ", vbCritical
      Exit Sub
   End If

   If (Trim(Textd(npuntDia)) = "" Or Trim(Textp(npuntDia)) = "") And Lblhr(npuntDia).Caption <> "" Then
      MsgBox "Verifique Turno/Horario..., día libre Turno es O ", vbCritical
      Exit Sub
   End If
   
   If Txccosto = "11" And nJefServ = 0 And TxDirector = "S" Then
      MsgBox "Ud. no esta autorizado a programar al servidor...", vbCritical, "[Jefe Servicio]"
      Exit Sub
   End If
   
   Call Recalcula
      
   If Val(TxHoras) > 150 Then
      Beep 500, 350
      MsgBox "Excede horas permitidas..!", vbCritical, "`[PROGRAMA]"
      Exit Sub
   End If
   
   If Txtarjeta <> "" Then
      sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "' and FLAG='2'"
      cvalor = ""
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs_d.EOF Then
         cvalor = rs_d!usuario
         If ceditor = "S" Then
         ElseIf cvalor <> wUsuario And cvalor <> wNick Then   '-Firma Usuario
                MsgBox "Operador, No Autorizado..! ", vbCritical
                Call CmdLimpia_Click
                rs_d.Close
                Exit Sub
         End If
      End If
      rs_d.Close
   End If
   
   On Error GoTo final
   Cnx.BeginTrans
   '-Actualiza DIRECTORIO "TURNO"
   Sw = False
   If Txtarjeta <> "" Then
      sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'  and FLAG='2'"
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      If rs_d.EOF Then
         Sw = True
         rs_d.AddNew
         rs_d!tarjeta = Txtarjeta
         rs_d!ano = genaño
         rs_d!mes = Txmes
         rs_d!usuario = wNick      'wUsuario      Firma Usuario
         rs_d!fecha = Date
         rs_d!hora = Format(Time, "hh:mm:ss")
         rs_d!estacion = w_clienName
      Else
         rs_d!usuarioM = wNick      'wUsuario Firma Usuario
         rs_d!fechaM = Date
      End If
      '-Imagen directorio mensual,generado hasta primer dia < Hoy
      cPlantilla0 = ""
      '--  MsgBox "inicio[" + cDirectHoy + "]"
      
      For i = Day(primer) To Day(ultimo)
          z = i - 1
          rs_d.Fields(z + 4) = IIf(IsNull(Textd(z)), "", Textd(z))
          If primer > Date Then
             '-Actualiza Plantilla0
             cPlantilla0 = cPlantilla0 + IIf(IsNull(Textd(z)), "==", Left(Trim(Textd(z)) + "==", 2))
             '-D N P MTMXD*V ==N*...
             ' 1234567890123456
          Else
             '-verifica en cDirectHoy cambio a Textd(z) --> AudDirect
             cvalor = Trim(Mid(cDirectHoy, (i - 1) * 2 + 1, 2))
             c_a = IIf(IsNull(Textd(z)), "==", Left(Trim(Textd(z)) + "==", 2))
             If cvalor <> c_a Then
                '--  MsgBox Trim(Str(i)) + ":de[" + cvalor + "]a[" + c_a + "]"
                   
                SqC = "INSERT INTO CPAudDirect(ano,mes,fecha,ccosto,codupr,tarjeta,deTurno,aTurno,Detalle,registro,estacion,fecregis,horregis) "
                SqC = SqC + " values('" + genaño + "','" + Txmes + "','" + Format(DateValue(primer + z), "yyyymmdd") + "','" + Txccosto + "','" + TxcodUPR + "','" + Txtarjeta + "','" & _
                cvalor + "','" + c_a + "','Directorio Asistenciales','" + wNick + "','" + w_clienName + "','" + Format(Date, "yyyymmdd") + "','" + Format(Time, "hh:mm:ss") + "')"
                Cnx.Execute SqC
             End If
          End If
      Next i
      rs_d!carEst = IIf(IsNull(TxCarEst), "", TxCarEst)
      rs_d!CCosto = IIf(IsNull(Txccosto), "", Txccosto)
      rs_d!CodUpr = IIf(IsNull(TxcodUPR), "", TxcodUPR)
      rs_d!Cargo = IIf(IsNull(TxCargo), "", TxCargo)
      rs_d!actividad = IIf(IsNull(Txactividad), "", Txactividad)
      rs_d!d = IIf(IsNull(TxD), "", TxD)
      rs_d!n = IIf(IsNull(TxN), "", TxN)
      rs_d!m = IIf(IsNull(TxM), "", TxM)
      rs_d!T = IIf(IsNull(Txt), "", Txt)
      rs_d!mT = IIf(IsNull(TxMT), "", TxMT)    '<=======
      rs_d!na = IIf(IsNull(TxNa), "", TxNa)    '<=======
      rs_d!pa = IIf(IsNull(TxPA), "", TxPA)    '<=======
      rs_d!Horas = IIf(IsNull(TxHoras), 0, Val(TxHoras))
      rs_d!Flag = "2"       '------------ASISTENCIAL-Complementarias
      rs_d!PLH = IIf(IsNull(TxPlH), "", TxPlH)    'Regimen laboral
      
      If DateValue(primer) > Date Then
         '-Graba imagen hasta primer dia < Hoy
         rs_d!plantilla0 = cPlantilla0
      End If
      
      '-SABER si una columna existe en una tabla
       SQ1 = "SELECT * FROM  INFORMATION_SCHEMA.COLUMNS AS c1  WHERE  c1.column_name = 'cadhep'  AND c1.table_name = 'cpdirectorio'"
       Set rs_2 = New Recordset
       rs_2.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
       If Not rs_2.EOF Then
          rs_d!CadHep = IIf(IsNull(TxCadHep), "", Trim(TxCadHep.Text))
       End If
       rs_2.Close
      rs_d.Update
      rs_d.Close
      
      If Sw Then
         D0.AddNew
      End If
      D0!tarjeta = Txtarjeta
      D0!nombre = Left(Txtrabajador, 25)
      D0!Cargo = Left(TxDscCargo, 15)
      D0!d = Val(TxD)
      D0!m = Val(TxM)
      D0!T = Val(Txt)
      D0!n = Val(TxN)
      D0!Horas = Val(TxHoras)
      For i = dfi To dff
         z = i - 1
         D0.Fields(z + 8) = IIf(IsNull(Textd(z)), "", Textd(z))
      Next i
      D0!mT = Val(TxMT)    '<=========
      D0!na = Val(TxNa)    '<=========
      D0!pa = Val(TxPA)    '<=========
      D0.Update
   End If
   DataGrid1.Refresh
   '-Actualiza Programa de HORARIOS del grupo
   For i = dfi To dff
      z = i - 1
      f_echa = Format(primer + z, "yyyymmdd")
      If Textp(z).Text <> "" Then
         Set rs = New Recordset
         If Txtarjeta <> "" Then
            sql = "SELECT * FROM CpHorCOMP WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & f_echa & "'"
         End If
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
         If c_a <> rs!Horario Then
            rs!usuario = wNick      'wUsuario     Firma Usuario
            rs!fecRegistro = Date
            rs!HorRegistro = Format(Time, "hh:mm:ss")
            rs!estacion = w_clienName
            If c_a <> "XXXX" Then
               '-cambio horario anterior
               rs!cambio = "X"
            End If
         End If
         rs.Update
         rs.Close
      Else
         If Txtarjeta <> "" Then
            Cnx.Execute "DELETE FROM CpHorCOMP WHERE tarjeta='" & Txtarjeta & "'AND fecha ='" & f_echa & "'"
         End If
      End If
   Next
   Cnx.CommitTrans
   
   Call Limpiacampos
   '-limpia grilla Hrs. Compl del Servidor
   DD.Delete
   DataGrid2.Refresh
   
   Frame1.Caption = "Programa"
   GoTo Final1
final:
   MsgBox Err.Description & " Error: " & Err.Number & vbCr
   Cnx.RollbackTrans
Final1:
End Sub
Private Sub CmdElimina_Click()
    If wrights = "CON" Or ceditor <> "S" Then
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
    sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "' and FLAG='2'"
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
    If MsgBox("Elimina COMPLEMENTARIAS del trabajador ?  ", vbQuestion + vbYesNo + vbDefaultButton2, "DIRECTORIO " + UCase(Trim(TxNombMes))) = vbYes Then
       fecha1 = Format(primer, "yyyymmdd")
       fecha2 = Format(ultimo, "yyyymmdd")
       If Txtarjeta <> "" Then
          Cnx.Execute "DELETE FROM CpHorCOMP WHERE tarjeta='" & Txtarjeta & "'AND fecha >='" & fecha1 & "' AND fecha<='" & fecha2 & "'"
          '-Actualiza DIRECTORIO "TURNO"
          Cnx.Execute "DELETE FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "' and FLAG='2'"
       End If
       DD.Delete    'hras Compl
       D0.Delete    'rol programa
    End If
    
    npuntero = D0.RecordCount
    RowSelected = DataGrid1.Row + 1
    Lblpuntero = "Puntero: " + Trim(Str(RowSelected)) + " de " + Trim(Str(npuntero))
 
    Call Limpiacampos
 End Sub
Private Sub cmdsalir_Click()
   Set D0 = Nothing
   Set DD = Nothing
   Set rs_d = Nothing
   Set rs = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        '- Forms Abiertos
        If UCase(Forms(cont).Name) <> "CPMENU" And UCase(Forms(cont).Name) <> "PROGCOMPLEMENTARIOS" Then
           Unload Forms(cont)
        End If
      End If
   Next
   Me.Hide
   Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set D0 = Nothing
   Set DD = Nothing
   Set rs_d = Nothing
   Set rs = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        '- Forms Abiertos
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "ProgComplementarios" Then
           Unload Forms(cont)
        End If
      End If
   Next
   Me.Hide
   Unload Me
End Sub
Private Sub CalcCierre()
   unafecha = primer
   dfi = Day(primer)
   dff = Day(ultimo)
   For i = dfi To dff
       z = i - 1
       ''If Weekday(unafecha + i - 1) = 1 Then
       If UCase(Left(WeekdayName(Weekday(primer + z)), 3)) = "DOM" Then
          '-feriado
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
                '-feriado
             End If
             rs.MoveNext
          Wend
          rs.Close
       End If
       Textp(z).Enabled = True
   Next
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
       Textd(z).Visible = True
       Textp(z).Visible = True
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
       Textp(z).Enabled = True
   Next
End Sub

Private Sub CboRegPlh_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboRegPlh.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub CboRegPlh_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboRegPlh.Text) Then
         TxRegPlh = ""
         Else
         j = CboRegPlh.ListIndex
         nposicion = InStr(CboRegPlh.List(j), ":")
         nlen = Len(CboRegPlh.List(j))
         TxRegPlh = Right(CboRegPlh.List(j), nlen - nposicion)
      End If
   End If
End Sub
Private Sub CboRegPlh_LostFocus()
  If IsEmpty(CboRegPlh.Text) Then
  Else
     j = CboRegPlh.ListIndex
     nposicion = InStr(CboRegPlh.List(j), ":")
     nlen = Len(CboRegPlh.List(j))
     TxRegPlh = Right(CboRegPlh.List(j), nlen - nposicion)
  End If
End Sub

Private Sub CmdPrint_Click()
   If Txmes <> "" And (Tx_codupr <> "" Or TxUprCCosto <> "" And nJefServ > 0) Then
      '-Programa mensual UPR7
      CargaProgCC
      If Tx_ccosto = "11" Then    'dpto. Enfermeria
         nTipRpt = 112.1
      Else
         nTipRpt = 112
      End If
      InfoCpRS.Left = 0: InfoCpRS.Top = 0
      InfoCpRS.Refresh
      InfoCpRS.Show
      Call Limpiacampos
      CmdPrint.Caption = ""
   End If
   Exit Sub
End Sub
Private Sub CmdAudDirect_Click()
   If Txmes <> "" And Tx_codupr <> "" Then
      '-Lista Cambios del MES en UPR
      nTipRpt = 1011.2
      InfoCpRS.Left = 0: InfoCpRS.Top = 0
      InfoCpRS.Refresh
      InfoCpRS.Show
      Call Limpiacampos
      CmdAudDirect.Caption = ""
   End If
   Exit Sub
End Sub

Private Sub CargaProgCC()
'***  CARGA  CodUpr-MES
    On Error Resume Next
    Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmpDCosto"
    Cnx.Execute ("SELECT * Into #Tabla  FROM CpTmpDCosto WHERE 1 = 2")
    Cnx.Execute "DELETE  FROM #tabla"
    Cnx.CommitTrans
    Set r_d = New ADODB.Recordset
    nposicion = 0
    '-Directorio Mensual de UPR
    W_RegPlh = ""
    If nJefServ > 0 Then
       '-Ccriterio
       sql = "SELECT * FROM CpDirectorio  WHERE ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "' and ccosto='" & TxUprCCosto & "'  AND flag='2' " + cCriterio + " ORDER BY Tarjeta"
       W_RegPlh = "(Jefaturas de Servicio)"
    ElseIf Trim(TxRegPlh) <> "" Then     'Regimen laboral
       sql = "SELECT  * FROM CpDirectorio WHERE ano='" & genaño & "' and mes='" & Txmes & "' and FLAG='2' and codupr='" & Tx_codupr & "' and " & _
       "Plh='" & TxRegPlh & "' ORDER BY ccosto,Tarjeta "
       If InStr("HEP_CAS", Trim(TxRegPlh)) > 0 Then
          W_RegPlh = "RM 0573-92-SA/DM - RD.Nº 030-93-SA-P/" + IIf(Trim(TxRegPlh) = "HEP", "D.L. 276", IIf(Trim(TxRegPlh) = "CAS", "D.L. 1057", ""))
       ElseIf Trim(TxRegPlh) = "3RO" And TxPlH <> "RES" Then
          W_RegPlh = ""
       End If
    Else
       sql = "SELECT  * FROM CpDirectorio WHERE ano='" & genaño & "' and mes='" & Txmes & "' and FLAG='2' and codupr='" & Tx_codupr & "' " & _
       "ORDER BY ccosto,Tarjeta "
       W_RegPlh = "RM 0573-92-SA/DM - RD.Nº 030-93-SA-P"
    End If
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
       CmdPrint.Caption = r_d!tarjeta
       Set rs = New Recordset
       rs.Open "SELECT * FROM cpgeneral WHERE  (baja is NULL) and tarjeta='" & r_d!tarjeta & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
       cNombre = "": cplh = ""
       ndia = 0
       If Not rs.EOF Then
          cNombre = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 30)
          If Trim(TxRegPlh) <> "" Then     'Regimen laboral
             cplh = TxRegPlh
          Else
             cplh = IIf(IsNull(rs!PLH), "", rs!PLH)
          End If
          If IsDate(rs!fecnac) Then
             If Month(rs!fecnac) = Val(Txmes) Then
                ndia = Day(rs!fecnac)
             End If
          End If
       End If
       rs.Close
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
       rs_d!nombre = cNombre
       '-Carga Plantilla 0 del servidor
       cPlantilla0 = IIf(IsNull(r_d!plantilla0), "", Trim(r_d!plantilla0))
       For i = Day(primer) To Day(ultimo)
          z = i - 1
          If ndia = i Then   'Onomastico
             rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "º", "º" + r_d.Fields(z + 4))  '--"¹"
          Else
             '-valor turno Actual
             c_a = IIf(IsNull(r_d.Fields(z + 4)), "==", r_d.Fields(z + 4))
             c_a = Left(Trim(c_a) + "==", 2)
             
             If cPlantilla0 <> "" Then
                '-verifica cambio en plantilla0
                cvalor = Trim(Mid(cPlantilla0, (i - 1) * 2 + 1, 2))
                If cvalor <> c_a Then
                    '-Cambio TurnoDia en Directorio
                    rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "¹", "¹" + r_d.Fields(z + 4))      '--"²"
                Else
                    '-No cambios
                    rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
                End If
             Else
                rs_d.Fields(z + 6) = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
             End If
          End If
          rs_d.Fields(z + 6 + 31) = Left(WeekdayName(Weekday(primer + z)), 1)   'LU/MA/MI/JU/VI/SA/DO
       Next i
       If cCargo <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM cpcargos WHERE codigo='" & cCargo & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If Not rs.EOF Then
             rs_d!DscCargo = IIf(IsNull(rs!descripcion), "", Left(rs!descripcion, 30))
          End If
          rs.Close
       End If
       rs_d!d = IIf(IsEmpty(r_d!d), "", r_d!d)
       rs_d!n = IIf(IsEmpty(r_d!n), "", r_d!n)
       rs_d!m = IIf(IsEmpty(r_d!m), "", r_d!m)
       rs_d!T = IIf(IsEmpty(r_d!T), "", r_d!T)
       rs_d!mT = IIf(IsEmpty(r_d!mT), "", r_d!mT)  '<====
       rs_d!na = IIf(IsEmpty(r_d!na), "", r_d!na)  '<====
       rs_d!pa = IIf(IsEmpty(r_d!pa), "", r_d!pa)  '<====
       rs_d!Horas = IIf(IsEmpty(r_d!Horas), 0, r_d!Horas)
       rs_d!Dia = ndia
       rs_plh = cplh
       rs_d.Update
       r_d.MoveNext
    Loop
    'Cnx.CommitTrans
    rs_d.Close
    r_d.Close
    Exit Sub
End Sub

Private Sub Recalcula()
   dfi = Day(primer)
   dff = Day(ultimo)
   TxHoras = "": TxD = "": TxN = "": TxM = "": Txt = "": TxMT = "": TxNa = "": TxPA = ""
   For i = dfi To dff
     z = i - 1
     If Textd(z).Text <> "" Then
        
        Select Case Trim(Textd(z))
           Case "MT", "MX"
                 TxMT = Val(TxMT) + 1
           Case "NA"
                 TxNa = Val(TxNa) + 1
           Case "PA"
                 TxPA = Val(TxPA) + 1
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
End Sub
