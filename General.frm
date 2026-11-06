VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form General 
   Caption         =   "Registro de Trabajadores"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15225
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   Picture         =   "General.frx":0000
   ScaleHeight     =   10990.87
   ScaleMode       =   0  'User
   ScaleWidth      =   15225
   Begin VB.TextBox Txtservidor 
      Enabled         =   0   'False
      Height          =   195
      Left            =   720
      MaxLength       =   3
      TabIndex        =   105
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox CboTservidor 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.CommandButton CmdNotas 
      Caption         =   "Notas"
      Height          =   519
      Left            =   13920
      Picture         =   "General.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Tx_Codupr 
      Enabled         =   0   'False
      Height          =   195
      Left            =   7320
      MaxLength       =   2
      TabIndex        =   91
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Cbo_Upr 
      Height          =   315
      Left            =   8400
      TabIndex        =   4
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox TxBusDni 
      Height          =   285
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   79
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   14040
      Picture         =   "General.frx":803D8
      Style           =   1  'Graphical
      TabIndex        =   88
      ToolTipText     =   "Limpiar Seleccion"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox TxRegPLH 
      Enabled         =   0   'False
      Height          =   195
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox CboRegPLH 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14160
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   0
      Picture         =   "General.frx":80CA2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox Txbuscar1 
      Height          =   285
      Left            =   10080
      MaxLength       =   15
      TabIndex        =   81
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Txbuscar 
      Height          =   285
      Left            =   3240
      MaxLength       =   5
      TabIndex        =   78
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox Tx_ccosto 
      Height          =   195
      Left            =   7680
      MaxLength       =   2
      TabIndex        =   37
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   735
      Left            =   14040
      Picture         =   "General.frx":818E4
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Cerrar"
      Top             =   9240
      Width           =   855
   End
   Begin VB.CommandButton mnuFicImpSelec 
      Height          =   495
      Left            =   14040
      Picture         =   "General.frx":81A2E
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Selec.Impresora"
      Top             =   3120
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   120
      TabIndex        =   73
      Top             =   3480
      Width           =   13575
      Begin VB.CheckBox ChkDirecColas 
         Alignment       =   1  'Right Justify
         Caption         =   "Directorio-Colas"
         Height          =   195
         Left            =   11760
         TabIndex        =   27
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox TxEMAIL 
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   108
         Top             =   4080
         Width           =   4935
      End
      Begin VB.TextBox TxCodReg 
         Height          =   285
         Left            =   7680
         MaxLength       =   6
         TabIndex        =   26
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox TxNivGuardia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   600
         MaxLength       =   2
         TabIndex        =   102
         Top             =   5640
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox CboNivGuardia 
         Height          =   315
         Left            =   1680
         TabIndex        =   32
         Top             =   5400
         Width           =   5415
      End
      Begin VB.TextBox TxdscCargo 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   100
         Top             =   3600
         Width           =   5175
      End
      Begin VB.TextBox TxEstCivil 
         Enabled         =   0   'False
         Height          =   195
         Left            =   9960
         MaxLength       =   7
         TabIndex        =   96
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox CboEstCivil 
         Height          =   315
         Left            =   11040
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox CboSexo 
         Height          =   315
         Left            =   7440
         TabIndex        =   14
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton CmdLactancia 
         Caption         =   "LACTANCIA"
         Height          =   495
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox ChkDirec 
         Alignment       =   1  'Right Justify
         Caption         =   "DIRECTOR"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   11880
         TabIndex        =   49
         ToolTipText     =   "Director de Unidad"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox TxCodUpR 
         Enabled         =   0   'False
         Height          =   195
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   51
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CboUPR 
         Height          =   315
         Left            =   720
         TabIndex        =   21
         Top             =   2520
         Width           =   5535
      End
      Begin VB.TextBox TxTipoPla 
         Enabled         =   0   'False
         Height          =   195
         Left            =   4440
         MaxLength       =   3
         TabIndex        =   41
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox CboPlh 
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox TxPlh 
         Enabled         =   0   'False
         Height          =   195
         Left            =   600
         MaxLength       =   3
         TabIndex        =   40
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CboGrpGen 
         Height          =   315
         Left            =   7920
         TabIndex        =   22
         Top             =   2520
         Width           =   5295
      End
      Begin VB.TextBox TxFecnac 
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
         Left            =   5640
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox TxRemPrin 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8880
         TabIndex        =   34
         Top             =   6000
         Width           =   855
      End
      Begin VB.TextBox Txfeccese 
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
         Left            =   12120
         TabIndex        =   29
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox TxFecingr 
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
         Left            =   9840
         TabIndex        =   28
         Top             =   3960
         Width           =   1095
      End
      Begin VB.ComboBox Cbocategoria 
         Height          =   315
         Left            =   1200
         TabIndex        =   19
         Top             =   1920
         Width           =   4815
      End
      Begin VB.TextBox TxcodNiv 
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         MaxLength       =   10
         TabIndex        =   61
         Top             =   6240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxCarPla 
         Enabled         =   0   'False
         Height          =   195
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   58
         Top             =   5040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxCarEst 
         Enabled         =   0   'False
         Height          =   195
         Left            =   11880
         MaxLength       =   2
         TabIndex        =   59
         Top             =   5040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxActividad 
         Enabled         =   0   'False
         Height          =   195
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   55
         Top             =   3240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox CboCodNiv 
         Height          =   315
         Left            =   1680
         TabIndex        =   33
         Top             =   6000
         Width           =   2895
      End
      Begin VB.ComboBox CboCargoE 
         Height          =   315
         Left            =   9120
         TabIndex        =   31
         Top             =   4680
         Width           =   4095
      End
      Begin VB.ComboBox CboCargoP 
         Height          =   315
         Left            =   1440
         TabIndex        =   30
         Top             =   4680
         Width           =   4815
      End
      Begin VB.TextBox TxCargo 
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Top             =   3600
         Width           =   495
      End
      Begin VB.ComboBox CboActividad 
         Height          =   315
         Left            =   7920
         TabIndex        =   24
         Top             =   3000
         Width           =   5295
      End
      Begin VB.TextBox TxEstado 
         Enabled         =   0   'False
         Height          =   195
         Left            =   7920
         MaxLength       =   1
         TabIndex        =   39
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Txtarjeta 
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox CboCadHep 
         Height          =   315
         Left            =   960
         TabIndex        =   23
         Top             =   3000
         Width           =   5295
      End
      Begin VB.TextBox TxCadHep 
         Enabled         =   0   'False
         Height          =   195
         Left            =   4560
         MaxLength       =   52
         TabIndex        =   53
         Top             =   3240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxApelMat 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   10800
         MaxLength       =   20
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxRemTot 
         Height          =   285
         Left            =   12360
         MaxLength       =   8
         TabIndex        =   35
         Top             =   6000
         Width           =   855
      End
      Begin VB.TextBox TxApelPat 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   20
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxDocumento 
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Txcategoria 
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         MaxLength       =   1
         TabIndex        =   43
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Txccosto 
         Enabled         =   0   'False
         Height          =   195
         Left            =   6360
         MaxLength       =   4
         TabIndex        =   48
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox CboCcosto 
         Height          =   315
         Left            =   7080
         TabIndex        =   20
         Top             =   1920
         Width           =   4335
      End
      Begin VB.ComboBox CboTipoPla 
         Height          =   315
         Left            =   5040
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox CboEstado 
         Height          =   315
         Left            =   8520
         TabIndex        =   18
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox TxNombres 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   12
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox TxGrpGen 
         Enabled         =   0   'False
         Height          =   225
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   52
         Top             =   2730
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label37 
         Caption         =   "EMAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   107
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "Registro"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6960
         TabIndex        =   106
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label LblTarifaGuardia 
         BackStyle       =   0  'Transparent
         Caption         =   "Lbltarifaguardia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   8520
         TabIndex        =   103
         Top             =   5400
         Width           =   3855
      End
      Begin VB.Label Label34 
         Caption         =   "Tarifa de Guardia"
         Height          =   375
         Left            =   240
         TabIndex        =   101
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil"
         Height          =   255
         Left            =   9960
         TabIndex        =   95
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "Sexo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6960
         TabIndex        =   93
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label26 
         Caption         =   "UPR"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   86
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "Grp Gen"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   6960
         TabIndex        =   85
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label25 
         Caption         =   "PLH"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   600
         TabIndex        =   84
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label24 
         Caption         =   "Tipo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   83
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "F. Nacimiento"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4560
         TabIndex        =   57
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Remuneración Principal"
         Height          =   375
         Left            =   7080
         TabIndex        =   82
         Top             =   6000
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   "F. Cese"
         Height          =   255
         Left            =   11280
         TabIndex        =   76
         Top             =   3960
         Width           =   735
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   13560
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   13560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label12 
         Caption         =   "Remuneración Total"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   10800
         TabIndex        =   71
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Nivel Remunerativo"
         Height          =   375
         Left            =   240
         TabIndex        =   70
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Cargo Estructural"
         Height          =   375
         Left            =   7560
         TabIndex        =   69
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Cargo Planilla"
         Height          =   375
         Left            =   240
         TabIndex        =   68
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Cargo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Actividad"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6960
         TabIndex        =   66
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Nombres"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido Materno"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   9360
         TabIndex        =   54
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido Paterno"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4920
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "Código"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "F.Ingreso"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   9000
         TabIndex        =   64
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Cadena"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   65
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "D.N.I."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   60
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Categoría"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   62
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "UndORG"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   6360
         TabIndex        =   63
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label20 
         Caption         =   "Estado"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   7920
         TabIndex        =   42
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.ComboBox Cbo_ccosto 
      Height          =   315
      Left            =   8400
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton CmdChgReg 
      Height          =   615
      Left            =   14040
      Picture         =   "General.frx":8220C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   13920
      TabIndex        =   72
      Top             =   3720
      Width           =   975
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   615
         Left            =   120
         Picture         =   "General.frx":823E7
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Agregar Registro "
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdElimina 
         Caption         =   "Elimina"
         Height          =   615
         Left            =   120
         Picture         =   "General.frx":82559
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Eliminar"
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton cmdGraba 
         Caption         =   "Graba"
         Height          =   615
         Left            =   120
         Picture         =   "General.frx":826CB
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Grabar"
         Top             =   1440
         Width           =   735
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
         Left            =   120
         Picture         =   "General.frx":8283D
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Imprimir"
         Top             =   2160
         Width           =   765
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   3625
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
      Caption         =   "Servidores"
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
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Servidor"
      Height          =   255
      Left            =   1200
      TabIndex        =   104
      Top             =   600
      Width           =   975
   End
   Begin VB.Label LblDS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LblDS"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   12840
      TabIndex        =   99
      Top             =   9960
      Width           =   2055
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "UBICAR POR ==>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   120
      TabIndex        =   97
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Espere  un momento carga tomará  unos instantes..!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2880
      TabIndex        =   94
      Top             =   2280
      Width           =   8535
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Servicio"
      Height          =   255
      Left            =   7680
      TabIndex        =   90
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "D.N.I."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   89
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg. Laboral"
      Height          =   255
      Left            =   1080
      TabIndex        =   87
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido Paterno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   8400
      TabIndex        =   80
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Tarjeta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   2400
      TabIndex        =   77
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Und. Org."
      Height          =   255
      Left            =   7680
      TabIndex        =   36
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "General"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DD As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim Sw As Boolean
Dim nposicion, ndifdias As Integer
Dim sql, SQL0, csexo, cClavePRG, cUPRold As String
Dim sobtiempo, nhrFin, nhrIni As Double
Private Sub Command1_Click()
     c_grupo = "02": c_secuencia = "021"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub
Private Sub Form_Load()
  Me.Caption = "Registro de Trabajadores - " & genaño & "   -   Sesión: " + wUsuNombre
  LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
  Set rs = New Recordset
  rs.Open "SELECT * FROM Cptablas WHERE tabla='11' and codigo<>'..'  ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboEstado.ListCount - 1
           If CboEstado.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboEstado.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpTablasAux WHERE grupo='REG' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboRegPLH.ListCount - 1
           If CboRegPLH.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboRegPLH.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboPlh.ListCount - 1
           If CboPlh.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboPlh.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  'LLena el Combo TipoPLA
   CboSexo.Clear
   CboSexo.AddItem "Masculino"
   CboSexo.AddItem "Femenino"
   CboSexo.ListIndex = 0
   CboSexo.Text = ""

   Set rs = New Recordset
   rs.Open "SELECT * FROM Tablasaux WHERE Grupo='SCI' and codigo<>'..'  ORDER BY descripcion", Cnx1, adOpenStatic, adLockReadOnly
   rs.MoveFirst
   While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboEstCivil.ListCount - 1
           If CboEstCivil.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboEstCivil.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close

  'LLena el Combo TipoPLA
   CboTservidor.Clear
    CboTservidor.AddItem "ASISTENCIAL"
    CboTservidor.AddItem "ADMINISTRATIVO"
    CboTservidor.ListIndex = 0
    CboTservidor.Text = ""
   
    CboTipoPla.Clear
    CboTipoPla.AddItem "ASISTENCIAL"
    CboTipoPla.AddItem "ADMINISTRATIVO"
    CboTipoPla.ListIndex = 0
    CboTipoPla.Text = ""
   
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpCcosto ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To CboCcosto.ListCount - 1
         If CboCcosto.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("Codigo") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        CboCcosto.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("Codigo")
     End If
     rs.MoveNext
  Wend
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
  rs.Open "SELECT * FROM CpGrupos where general='S' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboGrpGen.ListCount - 1
           If CboGrpGen.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("grupo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboGrpGen.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("grupo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpUndPrgRol ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To CboUPR.ListCount - 1
         If CboUPR.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        CboUPR.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR")
     End If
     rs.MoveNext
  Wend
  rs.MoveFirst
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To Cbo_Upr.ListCount - 1
         If Cbo_Upr.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        Cbo_Upr.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR")
     End If
     rs.MoveNext
  Wend
  rs.Close
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpTablasAux WHERE grupo='CAT' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To Cbocategoria.ListCount - 1
           If Cbocategoria.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          Cbocategoria.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  Set rs = New Recordset
  rs.Open "select * from CpCadena ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboCadHep.ListCount - 1
           If CboCadHep.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("cadena") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboCadHep.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("cadena")
       End If
       rs.MoveNext
  Wend
  rs.Close
  Set rs = New Recordset
  rs.Open "select * from CpActividad ORDER BY Codigo", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboActividad.ListCount - 1
           If CboActividad.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboActividad.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM Cptablas WHERE tabla='28' and codigo<>'..' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboCargoP.ListCount - 1
           If CboCargoP.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboCargoP.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  Set rs = New Recordset
  rs.Open "SELECT * FROM Cptablas WHERE tabla='27' and codigo<>'..'  ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboCargoE.ListCount - 1
           If CboCargoE.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboCargoE.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM Cptablas WHERE tabla='44' and codigo<>'..'  ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboNivGuardia.ListCount - 1
           If CboNivGuardia.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboNivGuardia.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpNIVEL order by codnivel", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboCodNiv.ListCount - 1
           If CboCodNiv.List(i) = rs.Fields("codnivel") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboCodNiv.AddItem rs.Fields("codnivel")
       End If
       rs.MoveNext
  Wend
  rs.Close
  Set DD = New Recordset
  With DD
     .Fields.Append "Tarjeta", adChar, 5
     .Fields.Append "Apelpat", adChar, 20
     .Fields.Append "ApelMat", adChar, 20
     .Fields.Append "Nombres", adChar, 20
     .Fields.Append "estado", adChar, 1
     .Fields.Append "fecnac", adDate
     .Fields.Append "Documento", adChar, 10    '6
     .Fields.Append "actividad", adChar, 4
     .Fields.Append "categoria", adChar, 1
     .Fields.Append "carEst", adChar, 2
     .Fields.Append "grpGen", adChar, 4     '10
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
          Case 0
          .Caption = "Tarjeta"
          .Width = Me.TextWidth(String(6, "S"))
          Case 1
          '.Caption =
          .Width = Me.TextWidth(String(14, "S"))
          Case 2
          '.Caption =
          .Width = Me.TextWidth(String(14, "S"))
          Case 3
          .Width = Me.TextWidth(String(18, "S"))
          Case 4
          .Caption = "Est"
          .Width = Me.TextWidth(String(3, "S"))
          Case 5
          .Width = Me.TextWidth(String(10, "S"))
          Case 6
          .Width = Me.TextWidth(String(10, "S"))
          Case 7
          .Width = Me.TextWidth(String(6, "S"))
          Case 9
          .Width = Me.TextWidth(String(5, "S"))
            Case 10
          .Width = Me.TextWidth(String(5, "S"))
         Case Else
          .Visible = False
          End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
  Next
  Txtarjeta.Locked = True
End Sub

Private Sub CboRegPlh_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboRegPLH.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboRegPlh_LostFocus()
  If IsEmpty(CboRegPLH.Text) Then
  Else
     j = CboRegPLH.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboRegPLH.List(j), ":")
        nlen = Len(CboRegPLH.List(j))
        TxRegPLH.Text = Right(CboRegPLH.List(j), nlen - nposicion)
        Select Case TxRegPLH.Text
          Case "HEP"
             wTituRpt = "(Nombrados/Contratados)"
          Case "CAS"
             wTituRpt = "(C.A.S.)"
          Case "3ER"
             wTituRpt = "(Terceros)"
          Case Else
             wTituRpt = "(Todos)"
        End Select
     End If
     Exit Sub
  End If
End Sub

Private Sub Cbo_ccosto_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Cbo_ccosto.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbo_ccosto_LostFocus()
   If IsEmpty(Cbo_ccosto.Text) Then
   Else
      j = Cbo_ccosto.ListIndex
      If j <> -1 Then
         nposicion = InStr(Cbo_ccosto.List(j), ":")
         nlen = Len(Cbo_ccosto.List(j))
         Tx_ccosto.Text = Right(Cbo_ccosto.List(j), nlen - nposicion)
         Tx_Codupr.Text = "": Cbo_Upr.Text = ""
         
         sql = "SELECT     tab1.CCOSTO, tab1.CODUPR, tab2.Descripcion " & _
         "FROM         CPCCostoUPR AS tab1 INNER JOIN " & _
         "                 CpUndPrgRol AS tab2 ON tab1.CODUPR = tab2.CodUPR " & _
         "WHERE CCOSTO='" & Tx_ccosto + "'  ORDER BY tab2.descripcion "
     
         Cbo_Upr.Clear
         Set rs = New Recordset
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly
         While Not rs.EOF
            ExisteYa = False
            For i = 0 To Cbo_Upr.ListCount - 1
                If Cbo_Upr.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR") Then
                    ExisteYa = True
                    Exit For
                End If
            Next i
            If Not ExisteYa Then
                Cbo_Upr.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR")
            End If
            rs.MoveNext
         Wend
         rs.Close
         
      End If
      Exit Sub
   End If
End Sub
Private Sub Cbotservidor_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboTservidor.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbotservidor_LostFocus()
  If IsEmpty(CboTservidor.Text) Then
  Else
     j = CboTservidor.ListIndex
     If j <> -1 Then
        Txtservidor = Left(CboTservidor.List(j), 3)
     End If
     Exit Sub
  End If
End Sub

Private Sub Cbo_UPR_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Cbo_Upr.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbo_UPR_LostFocus()
   If IsEmpty(Cbo_Upr.Text) Then
   Else
      j = Cbo_Upr.ListIndex
      If j <> -1 Then
         nposicion = InStr(Cbo_Upr.List(j), ":")
         nlen = Len(Cbo_Upr.List(j))
         Tx_Codupr.Text = Right(Cbo_Upr.List(j), nlen - nposicion)
         Tx_ccosto.Text = "": Cbo_ccosto.Text = ""
      End If
      Exit Sub
   End If
End Sub

Private Sub CmdChgReg_Click()
   Txbuscar.BackColor = &H80000005
   TxBusDni.BackColor = &H80000005
   Txbuscar1.BackColor = &H80000005
   sql = "SELECT * FROM CpGeneral WHERE (baja is NULL) "
   If Trim(TxRegPLH) <> "" Then
      sql = sql + " and  Plh='" & TxRegPLH & "' "
   End If
   If Trim(Tx_ccosto) <> "" Then
      sql = sql + " and Ccosto='" & Trim(Tx_ccosto) & "' "
   End If
   If Trim(Tx_Codupr) <> "" Then
      sql = sql + " and  codUpr='" & Trim(Tx_Codupr) & "' "
   End If
   If Trim(Txtservidor) <> "" Then
      sql = sql + " and TipoPla='" & Txtservidor & "' "
   End If
   sql = sql + "  ORDER BY apelpat,apelmat"
   SQL0 = sql
   DataGrid1.Visible = False
   Call cargaDD
   If DD.RecordCount = 0 Then
      DataGrid1.Visible = False
      MsgBox "No Existe datos...  ", vbExclamation
      Call CmdClear_Click
      Exit Sub
   Else
      DD.MoveFirst
   End If
   DataGrid1.Refresh
   DataGrid1.Visible = True
   
End Sub
Private Sub cargaDD()
   nposicion = 0
   Do While Not DD.EOF
      DD.Delete
      DD.MoveFirst
   Loop
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      rs.MoveFirst
      Do While Not rs.EOF
         nposicion = nposicion + 1
         DD.AddNew
         DD!tarjeta = rs!tarjeta
         DD!apelpat = IIf(IsNull(rs!apelpat), "", rs!apelpat)
         DD!apelmat = IIf(IsNull(rs!apelmat), "", rs!apelmat)
         DD!nombres = IIf(IsNull(rs!nombres), "", rs!nombres)
         DD!estado = IIf(IsNull(rs!estado), "", rs!estado)
         If IsDate(rs!fecnac) Then DD!fecnac = rs!fecnac
         DD!documento = IIf(IsNull(rs!documento), "", rs!documento)
         DD!actividad = IIf(IsNull(rs!actividad), "", rs!actividad)
         DD!Categoria = IIf(IsNull(rs!Categoria), "", rs!Categoria)
         DD!carEst = IIf(IsNull(rs!carEst), "", rs!carEst)
         DD!grpgen = IIf(IsNull(rs!grpgen), "", rs!grpgen)
         DD.Update
         rs.MoveNext
      Loop
   End If
   rs.Close
   Call Limpiacampos
End Sub

Private Sub CmdClear_Click()
   Call Limpiacampos
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   DataGrid1.Visible = False
   Do While Not DD.EOF
      DD.Delete
      DD.MoveFirst
   Loop
   DataGrid1.Refresh
   DataGrid1.Visible = True
   CboRegPLH.Text = "": TxRegPLH = ""
   CboTservidor.Text = "": Txtservidor = ""
   Cbo_ccosto.Text = "": Tx_ccosto = ""
End Sub

Private Sub Limpiacampos()
   LblTarifaGuardia.Caption = ""
   Txtarjeta = "": Txbuscar = "": TxBusDni = "": Txbuscar1 = ""
   Txbuscar.BackColor = &H80000005: TxBusDni.BackColor = &H80000005: Txbuscar1.BackColor = &H80000005
   
   TxApelPat = "": TxApelMat = "": TxNombres = "": TxFecnac = "": TxDocumento = ""
   TxEstCivil = "": CboEstCivil.Text = "": TxEstado = "": CboEstado.Text = ""
   Txcategoria = "": Cbocategoria.Text = "": Txccosto = "": CboCcosto.Text = ""
   TxFecingr = "": Txfeccese = "": TxGrpGen = "": CboGrpGen.Text = "": TxCodUpR = "": CboUPR.Text = ""
   TxCadHep = "": CboCadHep.Text = "": TxActividad = "": CboActividad.Text = "": Txcategoria.Text = "": Cbocategoria.Text = ""
   TxCargo = "": TxdscCargo = "": TxPlh = "": CboPlh.Text = ""
   TxCarPla = "": CboCargoP.Text = "": TxCarEst = "": CboCargoE.Text = ""
   TxNivGuardia = "": CboNivGuardia.Text = "":
   TxcodNiv = "": CboCodNiv.Text = "": TxRemPrin = "": TxRemTot = ""
   TxTipoPla = "": CboTipoPla.Text = ""
   CboSexo.Text = ""
   ChkDirec.Value = 0
   csexo = "": TxEMAIL = ""
   ChkDirecColas.Value = 0
   CmdLactancia.BackColor = &H8000000F   'Default => Gris
   'DataGrid1.SetFocus
End Sub
Private Sub Habilitacampos()
   Txtarjeta.Locked = True
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
    Txtarjeta = DataGrid1.Columns(0)
    TxApelPat = IIf(IsEmpty(DataGrid1.Columns(1)), "", Trim(DataGrid1.Columns(1)))
    TxApelMat = IIf(IsEmpty(DataGrid1.Columns(2)), "", Trim(DataGrid1.Columns(2)))
    TxNombres = IIf(IsEmpty(DataGrid1.Columns(3)), "", Trim(DataGrid1.Columns(3)))
    TxEstado = IIf(IsEmpty(DataGrid1.Columns(4)), "", DataGrid1.Columns(4))
    If IsDate(DataGrid1.Columns(5)) Then TxFecnac = DataGrid1.Columns(5)
    TxDocumento = IIf(IsEmpty(DataGrid1.Columns(6)), "", DataGrid1.Columns(6))
    TxActividad = IIf(IsEmpty(DataGrid1.Columns(7)), "", DataGrid1.Columns(7))
    Txcategoria = IIf(IsEmpty(DataGrid1.Columns(8)), "", DataGrid1.Columns(8))
    '-TxCadHep = IIf(IsEmpty(DataGrid1.Columns(9)), "", DataGrid1.Columns(9))
    TxCarEst = IIf(IsEmpty(DataGrid1.Columns(9)), "", DataGrid1.Columns(9))
    TxGrpGen = IIf(IsEmpty(DataGrid1.Columns(10)), "", DataGrid1.Columns(10))
    
    Call OtrosCampos
        
    If Trim(TxEstado) <> "" Then
       For i = 0 To CboEstado.ListCount - 1
         npos = InStr(CboEstado.List(i), ":")
         nlen = Len(CboEstado.List(i))
         dato = Trim(Right(CboEstado.List(i), nlen - npos))
         If dato = Trim(TxEstado) Then
            CboEstado.ListIndex = i
            Exit For
         End If
       Next i
    End If
    If Trim(Txcategoria) <> "" Then
       For i = 0 To Cbocategoria.ListCount - 1
         npos = InStr(Cbocategoria.List(i), ":")
         nlen = Len(Cbocategoria.List(i))
         dato = Trim(Right(Cbocategoria.List(i), nlen - npos))
         If dato = Trim(Txcategoria) Then
            Cbocategoria.ListIndex = i
            Exit For
         End If
       Next i
    End If
    
    If Trim(TxGrpGen) <> "" Then
       For i = 0 To CboGrpGen.ListCount - 1
         npos = InStr(CboGrpGen.List(i), ":")
         nlen = Len(CboGrpGen.List(i))
         dato = ""
         For j = nlen To 1 Step -1
             If Mid$(CboGrpGen.List(i), j, 1) = ":" Then
                npos = j
                Exit For
             End If
         Next j
         dato = Trim(Right(CboGrpGen.List(i), nlen - npos))
         If dato = Trim(TxGrpGen) Then
            CboGrpGen.ListIndex = i
            Exit For
         End If
       Next i
    End If
    
    If Trim(TxCadHep) <> "" Then
       For i = 0 To CboCadHep.ListCount - 1
         npos = InStr(CboCadHep.List(i), ":")
         nlen = Len(CboCadHep.List(i))
         dato = Trim(Right(CboCadHep.List(i), nlen - npos))
         If dato = Trim(TxCadHep) Then
            CboCadHep.ListIndex = i
            Exit For
         End If
       Next i
    End If
    If Trim(TxActividad) <> "" Then
       For i = 0 To CboActividad.ListCount - 1
         npos = InStr(CboActividad.List(i), ":")
         nlen = Len(CboActividad.List(i))
         dato = Trim(Right(CboActividad.List(i), nlen - npos))
         If dato = Trim(TxActividad) Then
            CboActividad.ListIndex = i
            Exit For
         End If
       Next i
    End If
    Txtarjeta.Locked = True
    
    CmdLactancia.BackColor = &H8000000F   'Varon No APLICA: Gris
    If Trim(TxDocumento) <> "" Then
       Set rs = New Recordset
       sql = "SELECT * FROM [General] where DNI='" & Trim(TxDocumento) & "'"
       rs.Open sql, Cnx1, adOpenStatic, adLockReadOnly, adCmdText
       If Not rs.EOF Then
          csexo = rs!sexo0
       End If
       rs.Close
    End If
    If csexo = "F" Then
       CmdLactancia.BackColor = &HC0FFC0     'Mujer  verde    OFF
    End If
    
    If Trim(cClavePRG) <> "" And w_clienteIP = "192.168.1.19" Then
       MsgBox "ClaveProg:[" + cClavePRG + "]"
    End If
    
    
    Sw = False
    Exit Sub
final:
  MsgBox Err.Description & " ErrOr: " & Err.Number & vbCr
End Sub
Private Sub dataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
 If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
  On Error GoTo final
   If KeyCode = 13 Then
     Call DataGrid1_Click
   End If
final:
   MsgBox Err.Description & " ERRor: " & Err.Number & vbCr
End Sub

Private Sub OtrosCampos()
    Set rs = New Recordset
    sql = "SELECT * FROM CPGeneral WHERE (baja is NULL) and  tarjeta='" & Txtarjeta & "'"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
       TxEstCivil = IIf(IsNull(rs!EstCivil), "", rs!EstCivil)
       Txccosto = IIf(IsNull(rs!CCosto), "", rs!CCosto) '11
       TxCodUpR = IIf(IsNull(rs!CodUpr), "", rs!CodUpr) '12
       cUPRold = IIf(IsNull(rs!CodUpr), "", Trim(rs!CodUpr))
       If IsDate(rs!Fecingr) Then TxFecingr = rs!Fecingr  '13
       If IsDate(rs!feccese) Then Txfeccese = rs!feccese '14
       TxCarPla = IIf(IsNull(rs!carPla), "", rs!carPla)     '15
       TxCadHep = IIf(IsNull(rs!CadHep), "", rs!CadHep) '16
       TxNivGuardia = IIf(IsNull(rs!NivGuardia), "", rs!NivGuardia)
       TxcodNiv = IIf(IsNull(rs!CodNiv), "", rs!CodNiv)  '17
       TxRemTot = IIf(IsNull(rs!RemTot), "", rs!RemTot)     '18
       TxCargo = IIf(IsNull(rs!cargo), "", rs!cargo)    '19
       TxPlh = IIf(IsNull(rs!PLH), "", rs!PLH)  '20
       TxTipoPla = IIf(IsNull(rs!tipoPla), "", rs!tipoPla)  '21
       ChkDirec.Value = IIf(rs!Director = "S", 1, 0)    '22
       TxCodReg = IIf(IsNull(rs!CODREG), "", rs!CODREG)  '23   COLEGIATURA
       cClavePRG = IIf(IsNull(rs!Clave), "", rs!Clave)
       TxEMAIL = IIf(IsNull(rs!Email), "", rs!Email)  '25
       ChkDirecColas.Value = IIf(IsNull(rs!DirecColas), 0, IIf(rs!DirecColas = "S", 1, 0)) '26
       ChkDirecColas.Enabled = True
       If TxCodUpR <> "03" Or TxEstado <> "A" Then
          ChkDirecColas.Enabled = False
       End If
       Encclave = ""
       WVdesencripta (cClavePRG)
       cClavePRG = Encclave
       
       If Trim(rs!sexo) <> "" Then
          dato = Trim(rs!sexo)     '23
          For i = 0 To CboSexo.ListCount - 1
             If dato = Left(CboSexo.List(i), 1) Then
                CboSexo.ListIndex = i
                Exit For
             End If
          Next i
       End If
       
    End If
    rs.Close
       
    If Trim(TxEstCivil) <> "" Then
          For i = 0 To CboEstCivil.ListCount - 1
            npos = InStr(CboEstCivil.List(i), ":")
            nlen = Len(CboEstCivil.List(i))
            dato = Trim(Right(CboEstCivil.List(i), nlen - npos))
            If dato = Trim(TxEstCivil) Then
                CboEstCivil.ListIndex = i
                Exit For
            End If
          Next i
    End If
    If Trim(TxPlh) <> "" Then
          For i = 0 To CboPlh.ListCount - 1
            npos = InStr(CboPlh.List(i), ":")
            nlen = Len(CboPlh.List(i))
            dato = Trim(Right(CboPlh.List(i), nlen - npos))
            If dato = Trim(TxPlh) Then
                CboPlh.ListIndex = i
                Exit For
            End If
          Next i
    End If
    If Trim(TxTipoPla) <> "" Then
          For i = 0 To CboTipoPla.ListCount - 1
            dato = Left(CboTipoPla.List(i), 3)
            If dato = Trim(TxTipoPla) Then
                CboTipoPla.ListIndex = i
                Exit For
            End If
          Next i
    End If
    If Trim(Txccosto) <> "" Then
         For i = 0 To CboCcosto.ListCount - 1
           npos = InStr(CboCcosto.List(i), ":")
           nlen = Len(CboCcosto.List(i))
           dato = Trim(Right(CboCcosto.List(i), nlen - npos))
           If dato = Trim(Txccosto) Then
              CboCcosto.ListIndex = i
              Exit For
           End If
         Next i
    End If
    If Trim(TxCodUpR) <> "" Then
         For i = 0 To CboUPR.ListCount - 1
           npos = InStr(CboUPR.List(i), ":")
           nlen = Len(CboUPR.List(i))
           dato = Trim(Right(CboUPR.List(i), nlen - npos))
           If dato = Trim(TxCodUpR) Then
              CboUPR.ListIndex = i
              Exit For
           End If
         Next i
    End If
    If Trim(TxCargo) <> "" Then
          Set rs = New Recordset
          sql = "SELECT * FROM CpCargos WHERE codigo='" & Trim(TxCargo) & "'"
          rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          TxdscCargo = ""
          If Not rs.EOF Then
             TxdscCargo = Trim(rs!descripcion)
          Else
             MsgBox "Cargo no existe...", vbExclamation, "[ ERROR ]"
             TxCargo = ""
          End If
          rs.Close
    End If
    If Trim(TxCarPla) <> "" Then
         For i = 0 To CboCargoP.ListCount - 1
            npos = InStr(CboCargoP.List(i), ":")
            nlen = Len(CboCargoP.List(i))
            dato = Trim(Right(CboCargoP.List(i), nlen - npos))
            If dato = Trim(TxCarPla) Then
              CboCargoP.ListIndex = i
              Exit For
            End If
         Next i
    End If
    If Trim(TxCarEst) <> "" Then
          For i = 0 To CboCargoE.ListCount - 1
            npos = InStr(CboCargoE.List(i), ":")
            nlen = Len(CboCargoE.List(i))
            dato = Trim(Right(CboCargoE.List(i), nlen - npos))
            If dato = Trim(TxCarEst) Then
               CboCargoE.ListIndex = i
               Exit For
            End If
          Next i
    End If
    
    If Trim(TxNivGuardia) <> "" Then
          For i = 0 To CboNivGuardia.ListCount - 1
            npos = InStr(CboNivGuardia.List(i), ":")
            nlen = Len(CboNivGuardia.List(i))
            dato = Trim(Right(CboNivGuardia.List(i), nlen - npos))
            If dato = Trim(TxNivGuardia) Then
               CboNivGuardia.ListIndex = i
               Exit For
            End If
          Next i
          
          Set rs = New Recordset
          sql = "SELECT * FROM CpTarifas WHERE (patindex('%" + Trim(TxNivGuardia) + "%' ,niveles)>0)   ORDER BY descripcion"
          rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          LblTarifaGuardia.Caption = ""
          If Not rs.EOF Then
             LblTarifaGuardia.Caption = Trim(rs!descripcion) + "=> S/." + Trim(Str(rs!DiaOrd))
          End If
          rs.Close
    End If
    
    If Trim(TxcodNiv) <> "" Then
          For i = 0 To CboCodNiv.ListCount - 1
            dato = Trim(CboCodNiv.List(i))
            If dato = Trim(TxcodNiv) Then
               CboCodNiv.ListIndex = i
               Exit For
            End If
          Next i
    End If
    
End Sub

Private Sub cmdNuevo_Click()
   Call Limpiacampos
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   'DataGrid1.Row = 0
   'DataGrid1.Col = 0
   Call Habilitacampos
   Txtarjeta.Locked = False
   Sw = True
End Sub
Private Sub cmdGraba_Click()
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   
       If Trim(TxEMAIL) <> "" Then
          TxEMAIL = UCase(TxEMAIL)
          If InStr(TxEMAIL, "@") = 0 Or InStr(TxEMAIL, "HEP.GOB.PE") = 0 Then
             Beep 500, 350
             MsgBox "Verifique dirección de correos. ", vbCritical, "Error Direccion]"
             TxEMAIL = ""
             Exit Sub
          End If
       End If
   
   
   If Trim(Txtarjeta) = "" Then
      MsgBox "No indico Nro tarjeta...  ", vbExclamation
      Exit Sub
   Else
      If Sw Then
         Set rs = New Recordset
         sql = "SELECT * FROM CpGeneral  WHERE  (baja is NULL) and  tarjeta='" & Txtarjeta & "'"
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            rs.Close
         Else
            MsgBox "Error, Tarjeta Ya existe..! ", vbExclamation
            rs.Close
            Exit Sub
         End If
      End If
   End If

   If IsEmpty(TxApelPat) Or IsEmpty(TxApelMat) Or IsEmpty(TxNombres) Then
      MsgBox "Registre, generales de trabajador..! ", vbExclamation
      Exit Sub
   End If
   
   j = CboSexo.ListIndex
   If j = -1 Then
      MsgBox "Favor, indique SEXO del servidor..! ", vbCritical
      Exit Sub
   End If

   If Trim(TxEstado) = "" Then
      MsgBox "Error, Indique estado trabajador..! ", vbExclamation
      Exit Sub
   End If
   If Trim(Txcategoria) = "" Then
      MsgBox "Defina categoria", vbExclamation
      Exit Sub
   End If
   If Trim(Txccosto) = "" Then
      MsgBox "Verifique C.Costo", vbExclamation
      Exit Sub
   End If
   If Trim(TxGrpGen) = "" And Trim(TxTipoPla) = "ADM" Then
      If MsgBox("No indico Grupo Generico... continua ?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
      Else
        Exit Sub
      End If
   End If
   If Trim(TxCodUpR) = "" And Trim(TxTipoPla) = "ASI" Then
      If MsgBox("No Indico Unidad Prog. Rol.. continua ?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
      Else
         Exit Sub
      End If
   End If
   If Trim(TxCodUpR) <> "" And Trim(TxTipoPla) = "ADM" Then
      Beep 500, 350
      If MsgBox("Ud. indica, UNIDAD PROG. ROL; No es correcto.. continua ?  ", vbQuestion + vbYesNo + vbDefaultButton2, "[Servidor ADMINISTRATIVO]") = vbYes Then
         TxCodUpR = ""
         CboUPR.Text = ""
      Else
         Exit Sub
      End If
   End If
   
   If Trim(TxCadHep) = "" And Trim(TxTipoPla) = "ASI" Then
      MsgBox "Indique Cadena...  ", vbExclamation
      Exit Sub
   End If
   If Trim(TxTipoPla) = "" And TxEstado = "A" Then
      MsgBox "Indique Tipo servidor...  ", vbExclamation
      Exit Sub
   End If
   If Trim(TxActividad) = "" Then
      MsgBox "Ingrese Actividad...  ", vbExclamation
      Exit Sub
   End If
   If Trim(TxCargo) = "" Then
      MsgBox "Indique Cargo..", vbExclamation
      Exit Sub
   End If
   If Val(TxRemTot) = 0 Then
      MsgBox "No indicò Remuneracion Total...  ", vbExclamation
   End If
   If Trim(TxPlh) = "" Then
      MsgBox "Indique clase trabajador...  ", vbExclamation
      Exit Sub
   End If
      
   If Trim(TxDocumento) = "" Then
      MsgBox "No indica Documento Identidad...  ", vbExclamation
   End If
   
   If Trim(TxFecnac) = "" Or Not IsDate(TxFecnac) Then
      Beep 500, 350
      MsgBox "Indicar Fecha Nacimiento...  ", vbExclamation
   End If
   
 
   '''On Error GoTo final
   '''Cnx.BeginTrans
   Set rt = New Recordset
   sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and  tarjeta ='" & Trim(Txtarjeta) & "'"
   rt.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If rt.EOF Then
      rt.AddNew
      rt!tarjeta = Trim(Txtarjeta)
      rt!fecha = Date
      rt!usuario = wUsuario         'Firma Usuario
   End If
   rt!apelpat = UCase(TxApelPat)
   rt!apelmat = UCase(TxApelMat)
   rt!nombres = UCase(TxNombres)
   rt!documento = IIf(IsEmpty(TxDocumento), "", TxDocumento)
   If IsDate(TxFecnac) Then
      rt!fecnac = TxFecnac
   Else
       rt!fecnac = Null
   End If
   rt!sexo = Left(CboSexo.List(j), 1)
   rt!EstCivil = Trim(TxEstCivil)
   rt!estado = Trim(TxEstado)
   rt!actividad = Trim(TxActividad)
   rt!Categoria = Trim(Txcategoria)
   rt!CadHep = Trim(TxCadHep)
   rt!grpgen = Trim(TxGrpGen)
   rt!CCosto = Trim(Txccosto)
   rt!CodUpr = Trim(TxCodUpR)
   If IsDate(TxFecingr) Then rt!Fecingr = TxFecingr
   If IsDate(Txfeccese) Then
      rt!feccese = Txfeccese
   Else
      rt!feccese = Null
   End If
   rt!carPla = Trim(TxCarPla)
   rt!carEst = Trim(TxCarEst)
   rt!NivGuardia = Trim(TxNivGuardia)
   rt!CodNiv = Trim(TxcodNiv)
   rt!RemTot = Val(TxRemTot)
   rt!cargo = Trim(TxCargo)
   rt!CODREG = Trim(TxCodReg)       '--COLEGIATURA
   rt!PLH = Trim(TxPlh)
   rt!tipoPla = Trim(TxTipoPla)
   rt!Director = IIf(ChkDirec.Value = 1, "S", 0)        '--director de UPR
   If Not Sw Then
      rt!usrCambia = wUsuario
      rt!fecCambia = Date
      rt!horCambia = Format(Time, "HH:mm:ss")
   End If
   rt!Email = Trim(TxEMAIL)
   rt!DirecColas = IIf(ChkDirecColas.Value = "1", "S", "")
   rt.Update
   rt.Close
   If Sw Then
      DD.AddNew
      DD!tarjeta = Txtarjeta
   End If
   DD!apelpat = UCase(TxApelPat)
   DD!apelmat = UCase(TxApelMat)
   
   DD!nombres = UCase(TxNombres)
   DD!estado = Trim(TxEstado)
   If IsDate(TxFecnac) Then DD!fecnac = TxFecnac
   DD!documento = Trim(TxDocumento)
   DD!actividad = Trim(TxActividad)
   DD!Categoria = Trim(Trim(Txcategoria))
   DD!carEst = Trim(TxCarEst)
   DD!grpgen = Trim(TxGrpGen)
   DD.Update
    
   '''Cnx.CommitTrans
   GoTo Final1
'''final:
'''   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
'''   Cnx.RollbackTrans
Final1:
   Call Limpiacampos
   Call Habilitacampos
   'DataGrid1.Row = 0
   'DataGrid1.Col =
   DataGrid1.SetFocus
   Sw = False
End Sub
Private Sub CmdElimina_Click()
    If wrights = "CON" Then
       MsgBox "No Autorizado...  ", vbExclamation
       Exit Sub
    End If
    If Trim(Txtarjeta) = "" Then Exit Sub
    Set rs = New Recordset
    sq2 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta ='" & Trim(Txtarjeta) & "'"
    rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then
       rs.Close
    Else
       MsgBox "Tiene Marcación en el Periodo...  ", vbExclamation
       rs.Close
       Exit Sub
    End If
    If MsgBox("Eliminar  registro ?      ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       sql = "UPDATE CpGeneral set baja='X',usrbaja='" + wUsuario + "',fecBaja='" + Format(Date, "yyyy-mm-dd") + "',horBaja='" + Format(Time, "HH:mm:dd") + "' WHERE (baja is NULL) and tarjeta='" & Txtarjeta & "'"
       Cnx.Execute sql
       DD.Delete
    End If
    Call Limpiacampos
    Call Habilitacampos
End Sub
Private Sub cmdsalir_Click()
  Set DD = Nothing
  Set rs_d = Nothing
  Set rs = Nothing
  Set rt = Nothing
  Me.Hide
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set DD = Nothing
  Set rs_d = Nothing
  Set rs = Nothing
  Set rt = Nothing
  Me.Hide
  Unload Me
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
   If Trim(Txtarjeta) = "" Then
      CargaTarjeta (0)
      nTipRpt = 1
      If Tx_Codupr <> "" Then
         nTipRpt = 1.01
      End If
      InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
      InfoCpSQL.Refresh
      InfoCpSQL.Show
   Else
      CargaTarjeta (1)
      nTipRpt = 2
      InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
      InfoCpSQL.Refresh
      InfoCpSQL.Show
   End If
   Exit Sub
End Sub

Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Set rs = New Recordset
      sql = "SELECT * FROM CpGeneral WHERE (baja is NULL) and  tarjeta='" & Trim(Txtarjeta) & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
      Else
         MsgBox "Nro Ya Existe.., NO puede Agregar..! ", vbExclamation
         rs.Close
         Exit Sub
      End If
      rs.Close
   End If
   'If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
   '   KeyAscii = 0
   'End If
End Sub
Private Sub TxDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Trim(TxDocumento) <> "" Then
          Set rs = New Recordset
          sql = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and documento='" & Trim(TxDocumento) & "'"
          rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If Not rs.EOF Then
             If rs!tarjeta <> Trim(Txtarjeta) Then
                Beep 500, 350
                MsgBox "Nro Documento, pertenece a servidor tarjeta " + rs!tarjeta + "..! ", vbCritical
             End If
             rs.Close
             Exit Sub
          End If
          rs.Close
       End If
    End If
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
       KeyAscii = 0
    End If
End Sub
Private Sub TxDocumento_LostFocus()
    Call TxDocumento_KeyPress(13)
End Sub

Private Sub Cbosexo_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboSexo.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbosexo_LostFocus()
  If IsEmpty(CboSexo.Text) Then
  Else
     j = CboSexo.ListIndex
     If j <> -1 Then
        '-SEXO = left(Cbosexo.List(j), 1)
     End If
     Exit Sub
  End If
End Sub
Private Sub CboEstCivil_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboEstCivil.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboEstCivil_LostFocus()
  If IsEmpty(CboEstCivil.Text) Then
  Else
     j = CboEstCivil.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboEstCivil.List(j), ":")
        nlen = Len(CboEstCivil.List(j))
        TxEstCivil = Right(CboEstCivil.List(j), nlen - nposicion)
     End If
     Exit Sub
  End If
End Sub
Private Sub Cboplh_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboPlh.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboPlh_Click()
  If IsEmpty(CboPlh.Text) Then
  Else
     j = CboPlh.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboPlh.List(j), ":")
        nlen = Len(CboPlh.List(j))
        TxPlh.Text = Right(CboPlh.List(j), nlen - nposicion)
     End If
     Exit Sub
  End If
End Sub
Private Sub Cboplh_LostFocus()
    Call CboPlh_Click
End Sub
Private Sub Cbotipopla_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboTipoPla.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbotipopla_LostFocus()
  If IsEmpty(CboTipoPla.Text) Then
  Else
     j = CboTipoPla.ListIndex
     If j <> -1 Then
        TxTipoPla.Text = Left(CboTipoPla.List(j), 3)
     End If
     Exit Sub
  End If
End Sub
Private Sub Cboestado_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboEstado.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cboestado_LostFocus()
  If IsEmpty(CboEstado.Text) Then
  Else
     j = CboEstado.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboEstado.List(j), ":")
        nlen = Len(CboEstado.List(j))
        TxEstado.Text = Right(CboEstado.List(j), nlen - nposicion)
        'Txtexto.Text = Left(Cbogrupos.List(j), nposicion-1)
     End If
     Exit Sub
  End If
End Sub
Private Sub Cbocategoria_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(Cbocategoria.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbocategoria_LostFocus()
  If IsEmpty(Cbocategoria.Text) Then
  Else
     j = Cbocategoria.ListIndex
     If j <> -1 Then
        nposicion = InStr(Cbocategoria.List(j), ":")
        nlen = Len(Cbocategoria.List(j))
        Txcategoria.Text = Trim(Right(Cbocategoria.List(j), nlen - nposicion))
     End If
     Exit Sub
  End If
End Sub
Private Sub Cboccosto_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboCcosto.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboCcosto_Click()
    If IsEmpty(CboCcosto.Text) Then
    
  Else
     j = CboCcosto.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboCcosto.List(j), ":")
        nlen = Len(CboCcosto.List(j))
        Txccosto.Text = Right(CboCcosto.List(j), nlen - nposicion)
     End If
     Exit Sub
  End If
End Sub
Private Sub Cboccosto_LostFocus()
    Call CboCcosto_Click
End Sub

Private Sub CbogrpGen_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboGrpGen.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboGrpGen_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(CboGrpGen.Text) = "" Then
        TxGrpGen = ""
     Else
         j = CboGrpGen.ListIndex
         If j <> -1 Then
            nlen = Len(CboGrpGen.List(j))
            For i = nlen To 1 Step -1
               If Mid(CboGrpGen.List(j), i, 1) = ":" Then
                  nposicion = i
                  Exit For
               End If
            Next
            TxGrpGen = Right(CboGrpGen.List(j), nlen - nposicion)
         End If
     End If
     Exit Sub
  End If
End Sub
Private Sub CbogrpGen_LostFocus()
   Call CboGrpGen_KeyPress(13)
End Sub

Private Sub CboUPR_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboUPR.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboUPR_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(CboUPR.Text) = "" Then
        TxCodUpR = ""
     Else
         j = CboUPR.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboUPR.List(j), ":")
            nlen = Len(CboUPR.List(j))
            TxCodUpR = Right(CboUPR.List(j), nlen - nposicion)
         End If
     End If
     If Trim(cUPRold) <> "" And Trim(cUPRold) <> Trim(TxCodUpR) Then
        MsgBox "Verifique, que servidor no tenga programación vigente en UPR " + cUPRold, vbExclamation, "[CAMBIA grupo UPR]"
     End If
     Exit Sub
  End If
End Sub
Private Sub CboUPR_LostFocus()
    Call CboUPR_KeyPress(13)
End Sub

Private Sub Cbocadhep_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboCadHep.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbocadhep_LostFocus()
  If IsEmpty(CboCadHep.Text) Then
  Else
     j = CboCadHep.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboCadHep.List(j), ":")
        nlen = Len(CboCadHep.List(j))
        TxCadHep.Text = Right(CboCadHep.List(j), nlen - nposicion)
     End If
     Exit Sub
  End If
End Sub
Private Sub Cbocatividad_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboActividad.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cboactividad_LostFocus()
  If IsEmpty(CboActividad.Text) Then
  Else
     j = CboActividad.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboActividad.List(j), ":")
        nlen = Len(CboActividad.List(j))
        TxActividad.Text = Right(CboActividad.List(j), nlen - nposicion)
     End If
     Exit Sub
  End If
End Sub

Private Sub TxCargo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(TxCargo) = "" Then
         Load Ayuda3
         Ayuda3.Tag = 1
         Ayuda3.Show 1
      End If
      Set rs = New Recordset
      sql = "SELECT * FROM CpCargos WHERE codigo='" & Trim(TxCargo) & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      TxdscCargo = ""
      If Not rs.EOF Then
         TxdscCargo = Trim(rs!descripcion)
      Else
         MsgBox "Cargo no existe...", vbExclamation, "[ ERROR ]"
         TxCargo = ""
      End If
      rs.Close
      Exit Sub
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
      KeyAscii = 0
   End If
End Sub
Private Sub TxCODREG_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Trim(TxCOREG) <> "" And Val(TxCodReg) > 0 Then
          TxCodReg = Right("000000" + Trim(Str(Val(TxCodReg))), 6)
       End If
    End If
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
       KeyAscii = 0
    End If
End Sub
Private Sub TxCODREG_LostFocus()
    Call TxCODREG_KeyPress(13)
End Sub

Private Sub Cbocargop_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboCargoP.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbocargop_LostFocus()
  If IsEmpty(CboCargoP.Text) Then
  Else
     j = CboCargoP.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboCargoP.List(j), ":")
        nlen = Len(CboCargoP.List(j))
        TxCarPla.Text = Right(CboCargoP.List(j), nlen - nposicion)
     End If
     Exit Sub
  End If
End Sub

Private Sub Cbocargoe_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboCargoE.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbocargoe_LostFocus()
  If IsEmpty(CboCargoE.Text) Then
  Else
     j = CboCargoE.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboCargoE.List(j), ":")
        nlen = Len(CboCargoE.List(j))
        TxCarEst.Text = Right(CboCargoE.List(j), nlen - nposicion)
        'Txtexto.Text = Left(Cbogrupos.List(j), nposicion-1)
     End If
     Exit Sub
  End If
End Sub

Private Sub CbonivGuardia_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboNivGuardia.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CbonivGuardia_LostFocus()
  If IsEmpty(CboNivGuardia.Text) Then
  Else
     j = CboNivGuardia.ListIndex
     LblTarifaGuardia.Caption = ""
     If j <> -1 Then
        nposicion = InStr(CboNivGuardia.List(j), ":")
        nlen = Len(CboNivGuardia.List(j))
        TxNivGuardia = Right(CboNivGuardia.List(j), nlen - nposicion)
        
        If Trim(TxNivGuardia) <> "" Then
           Set rs = New Recordset
           sql = "SELECT * FROM CpTarifas WHERE (patindex('%" + Trim(TxNivGuardia) + "%' ,niveles)>0)   ORDER BY descripcion"
           rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
           LblTarifaGuardia.Caption = ""
           If Not rs.EOF Then
              LblTarifaGuardia.Caption = Trim(rs!descripcion) + "=> S/." + Trim(Str(rs!DiaOrd))
           End If
           rs.Close
        End If
        
     End If
     Exit Sub
  End If
End Sub

Private Sub Cbocodniv_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboCodNiv.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbocodniv_LostFocus()
  If IsEmpty(CboCodNiv.Text) Then
  Else
     j = CboCodNiv.ListIndex
     If j <> -1 Then
        TxcodNiv = Trim(CboCodNiv.List(j))
     End If
     Exit Sub
  End If
End Sub


Private Sub TxEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Trim(TxEMAIL) <> "" Then
          TxEMAIL = UCase(TxEMAIL)
          If InStr(TxEMAIL, "@") = 0 Or InStr(TxEMAIL, "HEP.GOB.PE") = 0 Then
             MsgBox "Verifique dirección de correos. ", vbCritical, "Error Direccion]"
             Exit Sub
          End If
       End If
    End If
End Sub
Private Sub TxEmail_LostFocus()
    Call TxEmail_KeyPress(13)
End Sub


Private Sub CmdNotas_Click()
   If Trim(Txtarjeta) <> "" Then
      Load CpNotas
      CpNotas.Left = 2835: CpNotas.Top = 4755
      CpNotas.Refresh
      CpNotas.Show
   Else
      MsgBox "No indica Servidor..", vbExclamation
      Exit Sub
   End If
End Sub

Private Sub CmdLactancia_Click()
  If csexo <> "F" Then Exit Sub
  '-CmdLactancia.BackColor = &H8000000F   'Varon No APLICA: Gris
  '-CmdLactancia.BackColor = &HC0FFC0     'Mujer  verde    OFF
  '-CmdLactancia.BackColor = &H8080FF     '       Rojo     ON
  
  '+++++ FECHA BASE +++++++
  'MsgBox DateAdd("yyyy", 10, Date) ' agrega 10 años a la fecha actual
  'MsgBox DateAdd("m", 5, Date) ' Agrega cinco meses a la fecha actual
  'MsgBox DateAdd("D", -1, Date) ' resta 1 dias a la fecha actual ( usa el signo -)
   
  Load Lactancia
  Lactancia.Left = 1725: Lactancia.Top = 4035
  Lactancia.Refresh
  Lactancia.Show
End Sub

Private Sub CargaTarjeta(nopc)
'***  CARGA  registros  ***   0:Todos   1:Orden
    'Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmGeneral"
    If nopc = 0 Then
       sql = SQL0
       If Trim(TxRegPLH) <> "" Then
          Select Case TxRegPLH
               Case "HEP"
                 wTituRpt = "(Nombrados/Contratados)"
               Case "CAS"
                 wTituRpt = "(C.A.S.)"
               Case "3ER"
                 wTituRpt = "(Terceros)"
               Case Else
                 wTituRpt = "(Todos)"
          End Select
       End If
    Else
       sql = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta='" & Trim(Txtarjeta) & "'"
    End If
    On Error Resume Next
    Set rs_d = New Recordset
    rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Set rt = New Recordset
    sq_l = "SELECT * FROM CpTmGeneral"
    rt.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    Do While Not rs_d.EOF
       rt.AddNew
       rt!tarjeta = rs_d!tarjeta
       rt!apelpat = IIf(IsEmpty(rs_d!apelpat), "", UCase(rs_d!apelpat))
       rt!apelmat = IIf(IsEmpty(rs_d!apelmat), "", UCase(rs_d!apelmat))
       rt!nombres = IIf(IsEmpty(rs_d!nombres), "", UCase(rs_d!nombres))
       rt!estado = IIf(IsEmpty(rs_d!estado), "", rs_d!estado)
       rt!PLH = IIf(IsEmpty(rs_d!PLH), "", rs_d!PLH)
       rt!tipoPla = IIf(IsEmpty(rs_d!tipoPla), "", rs_d!tipoPla)
       If IsDate(rs_d!fecnac) Then rt!fecnac = rs_d!fecnac
       rt!documento = IIf(IsEmpty(rs_d!documento), "", rs_d!documento)
       rt!actividad = IIf(IsEmpty(rs_d!actividad), "", rs_d!actividad)
       rt!CadHep = IIf(IsEmpty(rs_d!CadHep), "", rs_d!CadHep)   '¿?
       rt!grpgen = IIf(IsEmpty(rs_d!grpgen), "", rs_d!grpgen)
       rt!CCosto = IIf(IsEmpty(rs_d!CCosto), "", rs_d!CCosto)
       rt!CodUpr = IIf(IsEmpty(rs_d!CodUpr), "", rs_d!CodUpr)
       rt!Categoria = IIf(IsEmpty(rs_d!Categoria), "", rs_d!Categoria)
       If IsDate(rs_d!Fecingr) Then rt!Fecingr = rs_d!Fecingr
       If IsDate(rs_d!feccese) Then rt!feccese = rs_d!feccese
       rt!carPla = IIf(IsEmpty(rs_d!carPla), "", rs_d!carPla)
       rt!carEst = IIf(IsEmpty(rs_d!carEst), "", rs_d!carEst)
       rt!NivGuardia = IIf(IsEmpty(rs_d!NivGuardia), "", rs_d!NivGuardia)
       rt!CodNiv = IIf(IsEmpty(rs_d!CodNiv), "", rs_d!CodNiv)
       rt!RemTot = IIf(IsEmpty(rs_d!RemTot), 0, rs_d!RemTot)
       rt!cargo = IIf(IsNull(rs_d!cargo), "", rs_d!cargo)
       rt!PLH = IIf(IsNull(rs_d!PLH), "", rs_d!PLH)
       rt!tipoPla = IIf(IsNull(rs_d!tipoPla), "", rs_d!tipoPla)
       cCodigo = IIf(IsNull(rs_d!actividad), "", rs_d!actividad)
       CmdPrint.Caption = rt!tarjeta
       If Trim(cCodigo) <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM CpActividad WHERE  codigo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dscactividad = "" + UCase(Left(rs!descripcion, 35))
          End If
          rs.Close
       End If
       cCodigo = IIf(IsNull(rs_d!CadHep), "", rs_d!CadHep)
       If Trim(cCodigo) <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM Cpcadena WHERE cadena='" & Trim(cCodigo) & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dsccadena = "" + Left(rs!descripcion, 150)
          End If
          rs.Close
       End If
       cCodigo = IIf(IsNull(rs_d!grpgen), "", rs_d!grpgen)
       If Val(cCodigo) > 0 Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM CPGrupos WHERE grupo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dscgrpGen = "" + UCase(Left(rs!descripcion, 35))
          End If
          rs.Close
       End If
       cCodigo = IIf(IsNull(rs_d!CCosto), "", rs_d!CCosto)
       If IsEmpty(cCodigo) Then
       Else
          Set rs = New Recordset
          rs.Open "SELECT * FROM CpCcosto WHERE codigo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dscccosto = "" + UCase(Left(rs!descripcion, 35))
          End If
          rs.Close
       End If
       cCodigo = IIf(IsNull(rs_d!CodUpr), "", rs_d!CodUpr)
       If IsEmpty(cCodigo) Then
       Else
          Set rs = New Recordset
          rs.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dscUPR = "" + UCase(Left(rs!descripcion, 55))
          End If
          rs.Close
       End If
       cCodigo = IIf(IsNull(rs_d!Categoria), "", rs_d!Categoria)
       If Trim(cCodigo) <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM CpTablasaux WHERE grupo='CAT' and codigo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dsccategoria = "" + UCase(Left(rs!descripcion, 20))
          End If
          rs.Close
       End If
       cCodigo = IIf(IsNull(rs_d!cargo), "", rs_d!cargo)
       If Trim(cCodigo) <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM CpCargos WHERE codigo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!DscCargo = "" + UCase(Left(rs!descripcion, 35))
          End If
          rs.Close
       End If
       cCodigo = IIf(IsNull(rs_d!carPla), "", rs_d!carPla)
       If Trim(cCodigo) <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM Cptablas WHERE tabla='28' and codigo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!DsccarPla = "" + Left(rs!descripcion, 30)
          End If
          rs.Close
       End If
       cCodigo = IIf(IsNull(rs_d!carEst), "", rs_d!carEst)
       If Trim(cCodigo) <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM Cptablas WHERE tabla='27' and codigo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dsccarest = "" + Left(rs!descripcion, 30)
          End If
          rs.Close
       End If
       rt.Update
       rs_d.MoveNext
    Loop
    'Cnx.CommitTrans
    CmdPrint.Caption = ""
    rt.Close
    rs_d.Close
End Sub

Private Sub TxBUscar_KeyPress(KeyAscii As Integer)
  If nposicion = 0 Then Exit Sub
  If KeyAscii = 13 Then
     Dim cBuscar As String
     cpuntero = DataGrid1.Bookmark
     If Txbuscar = "" Then Exit Sub
     cBuscar = "tarjeta like '" & Txbuscar & "*'"
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
Private Sub TxBusDni_KeyPress(KeyAscii As Integer)
  If nposicion = 0 Then Exit Sub
  If KeyAscii = 13 Then
     Dim cBusDni As String
     cpuntero = DataGrid1.Bookmark
     If TxBusDni = "" Then Exit Sub
     cBusDni = "Documento like '" & TxBusDni & "*'"
     DD.MoveFirst
     DD.Find (cBusDni)
     If DD.EOF Then
        TxBusDni.BackColor = &HFF&
     Else
        TxBusDni.BackColor = &H80000005
        DataGrid1.Refresh
        TxBusDni.Text = ""
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
     cBuscar = "Apelpat like '" & Txbuscar1 & "*'"
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
'+++++++++++++++++++++
' varfecha = "10/12/2005"
' "Select * from mytabla where fecha=#" & cdate(varfecha) &"#"

