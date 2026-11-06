VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Autorizaciones 
   Caption         =   "Autorizacion, Motivos Personales"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "Autorizaciones.frx":0000
   ScaleHeight     =   9120
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMINERIA 
      Height          =   855
      Left            =   10920
      Picture         =   "Autorizaciones.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "Minerìa de Datos"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdProducc 
      Caption         =   "Producción"
      Height          =   375
      Left            =   8400
      TabIndex        =   65
      ToolTipText     =   "Reg. del Dia"
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton CmdQuery 
      Caption         =   "Consulta"
      Height          =   375
      Left            =   6600
      TabIndex        =   62
      ToolTipText     =   "Check tarjeta"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Tx_diaHra 
      Enabled         =   0   'False
      Height          =   285
      Left            =   11640
      MaxLength       =   1
      TabIndex        =   61
      Top             =   6480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox Cbo_diaHra 
      Height          =   315
      Left            =   10560
      TabIndex        =   60
      Top             =   6480
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton mnuFicImpSelec 
      Height          =   495
      Left            =   9720
      Picture         =   "Autorizaciones.frx":806C5
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Selec.Impresora"
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "Graba"
      Height          =   615
      Left            =   10320
      Picture         =   "Autorizaciones.frx":80EA3
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Grabar"
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   615
      Left            =   11160
      Picture         =   "Autorizaciones.frx":81015
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Eliminar"
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   615
      Left            =   9480
      Picture         =   "Autorizaciones.frx":81187
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Agregar Registro "
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "Autorizaciones.frx":812F9
      Style           =   1  'Graphical
      TabIndex        =   42
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
      Left            =   10680
      Picture         =   "Autorizaciones.frx":81F3B
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Imprimir"
      Top             =   6840
      Width           =   765
   End
   Begin VB.CommandButton cmdChgReg 
      Height          =   615
      Left            =   9840
      Picture         =   "Autorizaciones.frx":82805
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Ubicar Registros"
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "[Registro]"
      Height          =   3255
      Left            =   120
      TabIndex        =   14
      Top             =   5520
      Width           =   9255
      Begin VB.TextBox TxTM 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   69
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxPa 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   68
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Txna 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   67
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxFecha 
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
         Left            =   2040
         TabIndex        =   66
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Limpia"
         Height          =   615
         Left            =   8400
         Picture         =   "Autorizaciones.frx":829E0
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Limpia Plantilla"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox Txplh 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7680
         MaxLength       =   3
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxT 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   59
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxM 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   58
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxN 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   57
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxD 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   56
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxHoras 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   55
         Top             =   2880
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxcodUpr 
         Height          =   285
         Left            =   6960
         MaxLength       =   2
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox Txtipopla 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         MaxLength       =   3
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame FrameH 
         Height          =   975
         Left            =   5280
         TabIndex        =   50
         Top             =   1440
         Width           =   3855
         Begin VB.OptionButton Option2 
            Caption         =   "Ingreso"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   23
            Top             =   600
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Salida"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   24
            Top             =   600
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Otros"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   25
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox TxTipo 
            Enabled         =   0   'False
            Height          =   195
            Left            =   600
            MaxLength       =   1
            TabIndex        =   26
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Txhorafin 
            Height          =   285
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Txhoraini 
            Height          =   285
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
         Begin VB.Shape Shape1 
            Height          =   15
            Left            =   3360
            Top             =   360
            Width           =   15
         End
         Begin VB.Label Label10 
            Caption         =   "Tipo  :"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Hora (hhmm)        Del"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Al  "
            Height          =   255
            Left            =   2760
            TabIndex        =   51
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame FrameD 
         Height          =   735
         Left            =   120
         TabIndex        =   48
         Top             =   1920
         Width           =   3975
         Begin VB.ComboBox CboTurno 
            Height          =   315
            Left            =   1560
            TabIndex        =   27
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "Programar turno en Directorio"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox TxCarEst 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7440
         MaxLength       =   2
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Txccosto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         MaxLength       =   2
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Txcargo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         MaxLength       =   4
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Txturno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         MaxLength       =   2
         TabIndex        =   28
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CboDH 
         Height          =   315
         Left            =   5640
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox TxDiaHora 
         Enabled         =   0   'False
         Height          =   195
         Left            =   5280
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Txtrabajador 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   7
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox Txtarjeta 
         Height          =   285
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Txcodigo 
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         MaxLength       =   4
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxFechaFin 
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
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   20
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox TxfechaIni 
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.ComboBox CboOcurrencia 
         Height          =   315
         Left            =   1200
         TabIndex        =   16
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox TxDocumento 
         Height          =   285
         Left            =   5400
         MaxLength       =   30
         TabIndex        =   29
         Top             =   2760
         Width           =   3615
      End
      Begin VB.Label LblMarcas 
         Alignment       =   2  'Center
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   5640
         TabIndex        =   64
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label LblHorario 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6600
         TabIndex        =   63
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "Tarbajador"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Por :"
         Height          =   255
         Left            =   5280
         TabIndex        =   40
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin"
         Height          =   255
         Left            =   2520
         TabIndex        =   39
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Fec Inicio dd/mm/aaaa"
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   15
         Left            =   3480
         TabIndex        =   31
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label8 
         Caption         =   "Documento"
         Height          =   255
         Left            =   4320
         TabIndex        =   30
         Top             =   2760
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   10920
      Picture         =   "Autorizaciones.frx":82F12
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Cerrar"
      Top             =   8160
      Width           =   735
   End
   Begin VB.TextBox Tx_trabajador 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   35
      TabIndex        =   4
      Top             =   600
      Width           =   3615
   End
   Begin VB.TextBox Tx_tarjeta 
      Height          =   285
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   195
      Left            =   4440
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   2175
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
      Height          =   285
      Left            =   8400
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7435
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
      Caption         =   "Transacciones delpresente MES"
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
      Left            =   8640
      TabIndex        =   70
      Top             =   8760
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA REGISTRO"
      Height          =   255
      Left            =   6840
      TabIndex        =   38
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label_2 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   255
      Left            =   1200
      TabIndex        =   37
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   1200
      TabIndex        =   36
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Autorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DD As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim Sw As Boolean
Dim nson, nposicion, ndifdias, nmesA, nmesI, nmesF, nano As Integer
Dim sql, cmes, ccadena, cmeses As String
Dim sobtiempo, nhrFin, nhrIni As Double
Dim cvalor, p_ing, p_ing1, p_sal, p_sal1, p_horas, cdiaant As String
Dim unafecha As Date
Dim fecCierre As Date
Dim lDejaProg  As Boolean

Private Sub Command1_Click()
     c_grupo = "02": c_secuencia = "024"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub

Private Sub Form_Load()
  Me.Caption = "Autorización, Motivos Personales - " & genaño & "   -   Sesión: " + wUsuNombre
  LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
  cmeses = "ENERO    FEBRERO  MARZO    ABRIL    MAYO     JUNIO    "
  cmeses = cmeses + "JULIO    AGOSTO   SETIEMBREOCTUBRE  NOVIEMBREDICIEMBRE"
  
  CboDH.Clear
  CboDH.AddItem "Dia :D"
  CboDH.AddItem "Hora:H"
  CboDH.Text = ""
  'CboDH.ListIndex = 0
  
  Cbo_diaHra.Clear
  Cbo_diaHra.AddItem "Dia :D"
  Cbo_diaHra.AddItem "Hora:H"
  Cbo_diaHra.Text = ""
  'Cbo_diaHra.ListIndex = 0
  
  CboTurno.Clear
  CboTurno.AddItem "Libre     :*"
  CboTurno.AddItem "Vacaciones:V"
  CboTurno.AddItem "Licencia  :L"
  CboTurno.AddItem "Descanso  :R"
  CboTurno.ListIndex = 0
  CboTurno.Text = ""
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM cptablas WHERE tabla='32' ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
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
  
  Txmes = genmes
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
  rs.Open "SELECT * FROM CpOcurrencias ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboOcurrencia.ListCount - 1
           If CboOcurrencia.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboOcurrencia.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  Set DD = New Recordset
  With DD
     .Fields.Append "tarjeta", adChar, 5
     .Fields.Append "Codigo", adChar, 4
     .Fields.Append "DiaHora", adChar, 1
     .Fields.Append "Fechaini", adDate
     .Fields.Append "FechaFin", adDate
     .Fields.Append "horaIni", adChar, 4
     .Fields.Append "horaFin", adChar, 4
     .Fields.Append "Tipo", adChar, 1
     .Fields.Append "Documento", adChar, 30
     .Fields.Append "Turno", adChar, 2
     .Fields.Append "Fecha", adDate     '10
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
          '.Caption = "codigo"
          .Width = Me.TextWidth(String(7, "S"))
          Case 2
          '.Caption ="DiaHora"
          .Width = Me.TextWidth(String(6, "S"))
          Case 3
          '.Caption ="fechaini"
          .Width = Me.TextWidth(String(10, "S"))
          Case 4
          '.Caption ="fechafin"
          .Width = Me.TextWidth(String(10, "S"))
          Case 5
          '.Caption ="HoraIni"
          .Width = Me.TextWidth(String(7, "S"))
          Case 6
          '.Caption ="HoraFin"
          .Width = Me.TextWidth(String(7, "S"))
          Case 7
          '.Caption ="Tipo"
          .Width = Me.TextWidth(String(6, "S"))
          Case 8
          '.Caption ="Documento"
           .Width = Me.TextWidth(String(30, "S"))
          Case Else
           .Visible = False
          End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
  Next
  DataGrid1.Enabled = False
  Call Deshabilita
End Sub
Private Sub Cbomes_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboMes.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbomes_LostFocus()
   If IsEmpty(CboMes.Text) Then
      Txmes = ""
   Else
      j = CboMes.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboMes.List(j), ":")
         nlen = Len(CboMes.List(j))
         Txmes = Right(CboMes.List(j), nlen - nposicion)
         unafecha = "01/" + Txmes + "/" + Trim(Str(genaño))
         Call ObtenerPrimerUltimoDia(unafecha)
         fecCierre = ultimo
        
         Set rs = New Recordset
         rs.Open "SELECT * FROM CpCronograma WHERE ano='" & genaño & "' and Mes='" & Txmes & "'", Cnx, adOpenStatic, adLockReadOnly   'CpHorarios
         If rs.EOF Then
            MsgBox "No Existe CRONOGRAMA del MES... Verifique", vbExclamation
         Else
            If IsDate(rs!cierremes) Then fecCierre = rs!cierremes
            If fecCierre <= Date Then
               Autorizaciones.Caption = "Autorizaciones, Motivos Personales - Programa CERRADO.  "
            Else
               Autorizaciones.Caption = "Autorizaciones, Motivos personales"
            End If
         End If
         rs.Close
      End If
      Exit Sub
   End If
End Sub
Private Sub Cbomes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboMes.Text) Then
         Txmes = ""
      Else
         j = CboMes.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboMes.List(j), ":")
            nlen = Len(CboMes.List(j))
            Txmes = Right(CboMes.List(j), nlen - nposicion)
            unafecha = "01/" + Txmes + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)
         End If
      End If
      Exit Sub
   End If
End Sub
Private Sub CmdProducc_Click()
   nson = 0
   If IsDate(TxDia) Then
      Call CargaProducc
      If nson = 0 Then
         MsgBox "No registro, Documentos..", vbExclamation
         Exit Sub
      End If
      W_lbldh = "Dias"
      nTipRpt = 3.01
      InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
      InfoCpSQL.Refresh
      InfoCpSQL.Show
   End If
   Exit Sub
End Sub

Private Sub Tx_tarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Tx_tarjeta) = "" Or (Trim(Tx_tarjeta) <> "" And EsTexto(Tx_tarjeta)) Then
         If Trim(Tx_tarjeta) <> "" And EsTexto(Tx_tarjeta) Then
            Ayuda1.TxUbiAP = Trim(Tx_tarjeta)
         End If
         Load Ayuda1
         Ayuda1.Tag = 4
         Ayuda1.Show 1
      Else
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Tx_tarjeta & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "Tarjeta, No Ubicada", vbExclamation
            Tx_tarjeta = ""
         Else
            Tx_trabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
         End If
         rs.Close
      End If
      If Txmes <> "" And Trim(Tx_tarjeta) <> "" Then
         Call CmdChgReg_Click
      End If
      Exit Sub
   End If
End Sub

Private Sub CmdQuery_Click()
     Load AutorizaQ
     AutorizaQ.Left = 2595: AutorizaQ.Top = 3285
     AutorizaQ.Refresh
     AutorizaQ.Show
End Sub

Private Sub CmdMINERIA_Click()
   Load AutorizaMIN
   AutorizaMIN.Left = 2595: AutorizaMIN.Top = 4985
   AutorizaMIN.Refresh
   AutorizaMIN.Show
End Sub

Private Sub CmdChgReg_Click()
   If Not IsDate(TxDia) And Txmes = "" Then Exit Sub
   If Trim(Tx_tarjeta) <> "" And Trim(Tx_trabajador) = "" Then
      Set rs = New Recordset
      sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Tx_tarjeta & "'"
      rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "No Ubicado", vbExclamation
         '  Tx_tarjeta.Text = ""
         Exit Sub
      Else
         Tx_trabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
      End If
      rs.Close
   End If
   If IsDate(TxDia) Then
      fecha1 = Format(TxDia, "yyyymmdd")
      Txmes = ""
      CboMes.Text = ""
   ElseIf Trim(Txmes) <> "" Then
          fecha1 = Format(primer, "yyyymmdd")
          fecha2 = Format(ultimo, "yyyymmdd")
   End If
   If Trim(Tx_tarjeta) = "" Then
      If IsDate(TxDia) Then   'registro
         sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and  fecha='" & fecha1 & "' ORDER BY tarjeta,FechaIni "
      ElseIf Trim(Txmes) <> "" Then
             sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and  fechaIni<='" & fecha2 & "' And  fechaFin>='" & fecha1 & "' ORDER BY tarjeta,fechaini "
      End If
   Else
      If IsDate(TxDia) Then
         sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and fecha='" & fecha1 & "' AND tarjeta='" & Tx_tarjeta & "' ORDER BY tarjeta,fechaIni "
      ElseIf Trim(Txmes) <> "" Then
             sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and fechaIni<='" & fecha2 & "' And  fechaFin>='" & fecha1 & "' and Tarjeta='" & Tx_tarjeta & "' ORDER BY tarjeta,fechaini "
      End If
   End If
   DataGrid1.Enabled = True
   
   Call cargaDD
End Sub
Private Sub cargaDD()
   nson = 0
   DD.Close
   
  DataGrid1.Visible = False
  Set DD = New Recordset
  With DD
     .Fields.Append "tarjeta", adChar, 5
     .Fields.Append "Codigo", adChar, 4
     .Fields.Append "DiaHora", adChar, 1
     .Fields.Append "Fechaini", adDate
     .Fields.Append "FechaFin", adDate
     .Fields.Append "horaIni", adChar, 4
     .Fields.Append "horaFin", adChar, 4
     .Fields.Append "Tipo", adChar, 1
     .Fields.Append "Documento", adChar, 30
     .Fields.Append "Turno", adChar, 2
     .Fields.Append "Fecha", adDate     '10
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
          '.Caption = "codigo"
          .Width = Me.TextWidth(String(7, "S"))
          Case 2
          '.Caption ="DiaHora"
          .Width = Me.TextWidth(String(6, "S"))
          Case 3
          '.Caption ="fechaini"
          .Width = Me.TextWidth(String(10, "S"))
          Case 4
          '.Caption ="fechafin"
          .Width = Me.TextWidth(String(10, "S"))
          Case 5
          '.Caption ="HoraIni"
          .Width = Me.TextWidth(String(7, "S"))
          Case 6
          '.Caption ="HoraFin"
          .Width = Me.TextWidth(String(7, "S"))
          Case 7
          '.Caption ="Tipo"
          .Width = Me.TextWidth(String(6, "S"))
          Case 8
          '.Caption ="Documento"
           .Width = Me.TextWidth(String(30, "S"))
          Case Else
           .Visible = False
          End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
  Next

   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      rs.MoveFirst
      Do While Not rs.EOF
         f_ok = False
         If Val(Txmes) > 0 Then
            cAAMM = genaño + Txmes
            If Left(Format(rs!fechaIni, "YYYYMMDD"), 6) <= cAAMM And Left(Format(rs!fechaFin, "YYYYMMDD"), 6) >= cAAMM Then
               f_ok = True
            End If
         Else
            f_ok = True
         End If
         If f_ok Then
            nson = nson + 1
            DD.AddNew
            DD!tarjeta = IIf(IsNull(rs!tarjeta), "", rs!tarjeta)
            DD!Codigo = IIf(IsNull(rs!Codigo), "", rs!Codigo)
            DD!diahora = IIf(IsNull(rs!diahora), "", rs!diahora)
            If IsDate(rs!fechaIni) Then DD!fechaIni = rs!fechaIni
            If IsDate(rs!fechaFin) Then DD!fechaFin = rs!fechaFin
            DD!horaini = IIf(IsNull(rs!horaini), "", rs!horaini)
            DD!horafin = IIf(IsNull(rs!horafin), "", rs!horafin)
            DD!tipo = IIf(IsNull(rs!tipo), "", rs!tipo)
            DD!documento = IIf(IsNull(rs!documento), "", rs!documento)
            DD!turno = IIf(IsNull(rs!turno), "", rs!turno)
            If IsDate(rs!fecha) Then DD!fecha = rs!fecha
            DD.Update
         End If
         rs.MoveNext
      Loop
   End If
   rs.Close
   DataGrid1.Refresh
   DataGrid1.Visible = True
   Call Deshabilita
   Call Limpiacampos
   If nson = 0 Then
      MsgBox "No registra autorizaciones..", vbExclamation
   End If
End Sub
Private Sub Limpiacampos()
   Txtarjeta = "": Txtrabajador = ""
   CboOcurrencia.Text = "": Txcodigo = "": Txfecha = ""
   CboDH.Text = "": TxDiaHora = ""
   TxfechaIni = "": TxFechaFin = ""
   LblHorario.Caption = "": LblMarcas.Caption = ""
   Option2(0).Value = False: Option2(1).Value = False: Option2(2).Value = False: TxTipo = ""
   Txhoraini = "": Txhorafin = ""
   TxTurno = "": CboTurno.Text = ""
   TxDocumento = ""
   'DataGrid1.SetFocus
End Sub
Private Sub Deshabilita()
   Txtarjeta.Enabled = False
   CboOcurrencia.Enabled = False
   TxfechaIni.Enabled = False: TxFechaFin.Enabled = False
   CboDH.Enabled = False
   FrameH.Enabled = False: FrameD.Enabled = False
   TxDocumento.Locked = True
End Sub
Private Sub Habilita()
   Txtarjeta.Enabled = True
   CboOcurrencia.Enabled = True
   TxfechaIni.Enabled = True: TxFechaFin.Enabled = True
   CboDH.Enabled = True
   FrameH.Enabled = True: FrameD.Enabled = True
   TxDocumento.Locked = False
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   With DD      ' Adodc1.Recordset
      If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
         .Sort = .Fields(ColIndex).[Name] & " Desc"
      Else
         .Sort = .Fields(ColIndex).[Name] & " Asc"
      End If
   End With
   Call Deshabilita
   Call Limpiacampos
End Sub
Private Sub DataGrid1_Click()
    If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
    Call dataGrid1_KeyDown(13, 0)
End Sub
Private Sub dataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   On Error GoTo final
   If KeyCode = 13 Then
      Call Deshabilita
      Call Limpiacampos
      Txtarjeta = IIf(IsEmpty(DataGrid1.Columns(0)), "", DataGrid1.Columns(0))
      Txcodigo = IIf(IsEmpty(DataGrid1.Columns(1)), "", DataGrid1.Columns(1))
      TxDiaHora = IIf(IsEmpty(DataGrid1.Columns(2)), "", DataGrid1.Columns(2))
      If IsDate(DataGrid1.Columns(3)) Then TxfechaIni = DataGrid1.Columns(3)
      If IsDate(DataGrid1.Columns(4)) Then TxFechaFin = DataGrid1.Columns(4)
      Txhoraini = IIf(IsEmpty(DataGrid1.Columns(5)), "", DataGrid1.Columns(5))
      Txhorafin = IIf(IsEmpty(DataGrid1.Columns(6)), "", DataGrid1.Columns(6))
      TxTipo = IIf(IsEmpty(DataGrid1.Columns(7)), "", DataGrid1.Columns(7))
      TxDocumento = IIf(IsEmpty(DataGrid1.Columns(8)), "", DataGrid1.Columns(8))
      TxTurno = IIf(IsEmpty(DataGrid1.Columns(9)), "", DataGrid1.Columns(9))
      If IsDate(DataGrid1.Columns(10)) Then Txfecha = DataGrid1.Columns(10)
      If Trim(Txtarjeta) <> "" Then
         Set rs = New Recordset
         rs.Open "SELECT * FROM CpGeneral WHERE (baja is NULL) and  tarjeta='" & Txtarjeta & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs.EOF Then
            Txtrabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
            Txccosto = IIf(IsNull(rs!CCosto), "", rs!CCosto)
            TxCodUpR = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)
            TxCargo = IIf(IsNull(rs!Cargo), "", rs!Cargo)
            TxCarEst = IIf(IsNull(rs!carEst), "", rs!carEst)
            TxPlH = IIf(IsNull(rs!PLH), "", rs!PLH)
            TxTipoPla = IIf(IsNull(rs!tipoPla), "", rs!tipoPla)
         End If
         rs.Close
         Txtarjeta.Enabled = False
      End If
      If Trim(TxDiaHora) <> "" Then
        For i = 0 To CboDH.ListCount - 1
          npos = InStr(CboDH.List(i), ":")
          nlen = Len(CboDH.List(i))
          dato = Trim(Right(CboDH.List(i), nlen - npos))
          If dato = Trim(TxDiaHora) Then
             CboDH.ListIndex = i
             Exit For
          End If
        Next i
      End If
      If Trim(Txcodigo) <> "" Then
        For i = 0 To CboOcurrencia.ListCount - 1
          npos = InStr(CboOcurrencia.List(i), ":")
          nlen = Len(CboOcurrencia.List(i))
          dato = Trim(Right(CboOcurrencia.List(i), nlen - npos))
          If dato = Trim(Txcodigo) Then
             CboOcurrencia.ListIndex = i
             Exit For
          End If
        Next i
      End If
      If Trim(TxDiaHora) = "H" Then
        FrameD.Enabled = False
        If Trim(TxTipo) = "I" Then
            Option2(0).Value = True
         ElseIf Trim(TxTipo) = "S" Then
                Option2(1).Value = True
         ElseIf Trim(TxTipo) = "O" Then
                Option2(2).Value = True
         End If
         TxTurno = "": CboTurno.Text = ""
      ElseIf Trim(TxDiaHora) = "D" Then
             Txhoraini = "": Txhorafin = ""
             FrameH.Enabled = False
             If Trim(TxTurno) <> "" Then
                For i = 0 To CboTurno.ListCount - 1
                  npos = InStr(CboTurno.List(i), ":")
                  nlen = Len(CboTurno.List(i))
                  dato = Trim(Right(CboTurno.List(i), nlen - npos))
                  If dato = Trim(TxTurno) Then
                     CboTurno.ListIndex = i
                     Exit For
                  End If
                Next i
             End If
      End If
      Call Deshabilita
      
      p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": cdiaant = "N": c2dias = "N"
      m_ing = "": m_sal = "": m_ing1 = "": m_sal1 = ""
      fechaB = Format(TxfechaIni, "yyyymmdd")
      Fechab0 = Format(DateValue(TxfechaIni) - 1, "yyyymmdd")
      sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
      Set rs = New Recordset
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         rs.Close
         sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & Fechab0 & "'"
         Set rs = New Recordset
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         cdiaant = "S"
      End If
      If Not rs.EOF Then
         cHorario = IIf(IsNull(rs!Horario), "", rs!Horario)
         If cHorario <> "" Then
            rs.Close
            Set rs = New Recordset
            sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & cHorario & "'"
            rs.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs.EOF Then
               p_ing = IIf(IsNull(rs!ingreso), "", rs!ingreso)
               p_sal = IIf(IsNull(rs!salida), "", rs!salida)
               p_ing1 = IIf(IsNull(rs!Ingreso1), "", rs!Ingreso1)
               p_sal1 = IIf(IsNull(rs!salida1), "", rs!salida1)
               p_horas = IIf(IsNull(rs!Horas), "0000", rs!Horas)
               LblHorario = p_ing + "-" + p_sal + " / " + p_ing1 + "-" + p_sal1 + " " + IIf(cdiaant <> "", "(D/A)", "")
               If Val(p_horas) > 0 Then
                  '-Marcas en DIA Programado
                  If (p_ing > p_sal And Trim(p_sal1) = "") Or (p_ing > p_sal1 And Trim(p_ing1) <> "") Then
                     c2dias = "S"
                  End If
                  Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
                  If cdiaant = "S" Then       'ubica marcas del dia anterior
                     If c2dias = "S" Then
                        Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha >='" & Fechab0 & "' AND fecha <='" & fechaB & "' ORDER BY fecha,hora"
                     End If
                  Else
                     If c2dias = "S" Then
                        Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha >='" & fechaB & "' AND fecha <='" & FechamF & "' ORDER BY fecha,hora"
                     End If
                  End If
                  Set rs_d = New Recordset
                  rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  If Not rs_d.EOF Then
                     Do While Not rs_d.EOF
                        If Trim(m_ing) = "" Then
                           m_ing = rs_d!hora
                        ElseIf Trim(m_sal) = "" Then
                               m_sal = rs_d!hora
                        ElseIf Trim(m_ing1) = "" Then
                               m_ing1 = rs_d!hora
                        ElseIf Trim(m_sal1) = "" Then
                               m_sal1 = rs_d!hora
                        End If
                        rs_d.MoveNext
                     Loop
                  End If
                  rs_d.Close
                  LblMarcas = "Marcas de reloj: " + m_ing + "-" + m_sal + " / " + m_ing1 + "-" + m_sal1 + " " + IIf(c2dias <> "", "(N/D)", "")
               End If
            End If
         End If
      End If
      rs.Close
      Sw = False
      Exit Sub
   End If
final:
   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
End Sub
Private Sub CmdClear_Click()
  Call Limpiacampos
  cmdsalir.SetFocus
End Sub
Private Sub cmdNuevo_Click()
   Call Limpiacampos
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If Val(genaño) > Year(Date) Then
      Beep 500, 350
      Beep 500, 350
      MsgBox "NO permite Transacciones; hacerlo en AÑO ACTUAL... ", vbCritical
      Exit Sub
   End If
   Call Habilita
   If Tx_tarjeta <> "" Then
      Txtarjeta = Tx_tarjeta
      Txtarjeta.Enabled = False
      Txfecha = Date
      Set rs = New Recordset
      rs.Open "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta='" & Txtarjeta & "'", Cnx, adOpenStatic, adLockReadOnly
      If Not rs.EOF Then
         Txtrabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
         Txccosto = IIf(IsNull(rs!CCosto), "", rs!CCosto)
         TxCodUpR = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)
         TxCargo = IIf(IsNull(rs!Cargo), "", rs!Cargo)
         TxCarEst = IIf(IsNull(rs!carEst), "", rs!carEst)
         TxPlH = IIf(IsNull(rs!PLH), "", rs!PLH)
         TxTipoPla = IIf(IsNull(rs!tipoPla), "", rs!tipoPla)
      End If
      rs.Close
      CboOcurrencia.SetFocus
   Else
      Txtarjeta.Enabled = True
      Txtarjeta.SetFocus
   End If
   Sw = True
End Sub
Private Sub cmdGraba_Click()
   ccadena = ""
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If fecCierre <= Date Then      'And wDerEspec <> "S"
      MsgBox "Programa CERRADO...  ", vbExclamation
      Exit Sub
   End If
   If Not Sw Then
      MsgBox "No se permite Cambios... elimine y registre!  ", vbCritical
      Exit Sub
   End If
   If Trim(Txtarjeta) = "" Then
      MsgBox "No indico, Nro. tarjeta...  ", vbExclamation
      Exit Sub
   Else
      Set rs = New Recordset
      sql = "SELECT * FROM CpGeneral  WHERE (baja is NULL) and  tarjeta='" & Txtarjeta & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "Error, Tarjeta No existe..! ", vbExclamation
         rs.Close
         Exit Sub
      Else
         rs.Close
      End If
   End If
   
   If Not IsDate(TxfechaIni) Or Not IsDate(TxFechaFin) Or DateValue(TxfechaIni) > DateValue(TxFechaFin.Text) Then
      MsgBox "Verifique fechas...  ", vbExclamation
      Exit Sub
   End If
   If Trim(TxDiaHora) = "" Then
      MsgBox "Defina autorización,  Por Dia/hora ", vbExclamation
      Exit Sub
   End If
   If Trim(Txcodigo) = "" Then
      MsgBox "Indique Motivo ..", vbExclamation
      Exit Sub
   End If
    
   If TxDiaHora = "H" Then
      If Val(Txhorafin) > 2400 Then
         MsgBox "Verifique, hora fin..", vbExclamation
         Exit Sub
      Else
         p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = ""
         'verifica Horario programado
         fechaB = Format(TxfechaIni, "yyyymmdd")
         sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
         Set rs = New Recordset
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         cerror = "S"       '<============
         If rs.EOF Then
            'Verifica caso autoriza sobre turno de 2Dias: 1930-0730 autoriza 0430-0630
            rs.Close
            fechaB = Format(DateValue(TxfechaIni) - 1, "yyyymmdd")
            sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
            Set rs = New Recordset
            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         End If
         If Not rs.EOF Then
            cHorario = IIf(IsNull(rs!Horario), "", rs!Horario)
            rs.Close
            If cHorario <> "" Then
               Set rs = New Recordset
               sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & cHorario & "'"
               rs.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not rs.EOF Then
                  p_ing = IIf(IsNull(rs!ingreso), "", rs!ingreso)
                  p_sal = IIf(IsNull(rs!salida), "", Trim(rs!salida))
                  p_ing1 = IIf(IsNull(rs!Ingreso1), "", Trim(rs!Ingreso1))
                  p_sal1 = IIf(IsNull(rs!salida1), "", Trim(rs!salida1))
                  p_horas = IIf(IsNull(rs!Horas), "0000", rs!Horas)
                  rs.Close
                  
                  'Evalua Marcas
                  If p_ing < p_sal And p_ing1 = "" Then
                     'P:0730-1930
                     If p_ing <= Txhoraini Then     'And FinOk(Txhorafin, p_ing, p_horas)
                        'm:0730-0930
                        cerror = "N"
                     End If
                  ElseIf p_ing > p_sal And p_ing1 = "" Then     '2dias
                         'P:1930-0730
                         If Txhoraini < Txhorafin Then
                            'm:1930-2030
                            If p_ing <= Txhoraini Then    'And FinOk(Txhorafin, p_ing, p_horas)
                               cerror = "N"
                            ElseIf p_ing >= Txhoraini Then    'And FinOk(Txhorafin, p_ing, p_horas)
                                   'm:0230-0330
                                   cerror = "N"
                            End If
                         ElseIf Txhoraini > Txhorafin Then
                                'm:2130-0230
                                If p_ing <= Txhoraini Then    'And FinOk(Txhorafin, p_ing, p_horas)
                                   cerror = "N"
                                End If
                         End If
                  ElseIf p_ing < p_sal And p_ing1 <> "" Then
                         'P:0800-1400 / 2000-0800,  1400-2000 / 2000-0800
                         If p_ing <= Txhoraini Then   'And FinOk(Txhorafin, p_ing, p_horas)
                            'm:0800-0900
                            cerror = "N"
                         ElseIf p_ing >= Txhorafin Then   'And FinOk(Txhorafin, p_ing, p_horas)
                            'm:0700-0800
                            cerror = "N"
                         End If
                  End If
               End If
               If cerror = "S" Then
                  MsgBox "Verifique, Rango Horas.. ", vbExclamation
               End If
            Else
               MsgBox "Verifique, Sin Horario..", vbExclamation
            End If
            If cerror = "S" Then
               Exit Sub
            End If
         Else
            MsgBox "Verifique, No programo Horario..", vbExclamation
            rs.Close
            Exit Sub
         End If
      End If
   End If
   If TxDiaHora = "D" And Trim(TxTurno) = "" Then
      MsgBox "Autorización, sin cambiar TURNO del Directorio ..", vbExclamation, "[IMPORTANTE]"
   End If
   
   
   If TxDiaHora = "D" And Left(Txcodigo, 2) = "VA" Then
      nndias = DateDiff("d", DateValue(TxfechaIni), DateValue(TxFechaFin)) + 1
      Call CHKDias(Txtarjeta, Txcodigo)
      If nndias > 30 Then
         MsgBox "Son " + Trim(Str(nndias)) + " dìas, EXCEDE el limite.. ", vbCritical, "[VACACIONES]"
         Exit Sub
      Else
         MsgBox "Son " + Trim(Str(nndias)) + " dìas, ACUMULADOS.. ", vbInformation, "[VACACIONES]"
      End If
   End If
   
   fecha1 = Format(TxfechaIni, "yyyymmdd")
   L_ok = True
   If Trim(TxDiaHora) = "D" And Trim(TxTurno) <> "" Then
      '-verifica Programa Directorio,  ==> mismo mes de APLICACION; tarjeta,txmes,txfechaini,txfechafin
      If Month(TxFechaFin) = Val(Txmes) And Month(TxfechaIni) = Month(TxFechaFin) Then
         sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
         Set rs_d = New Recordset
         rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         'Verifica si existe, Registros Programacion
         nson = 0
         If Not rs_d.EOF Then
            For i = Day(TxfechaIni) To Day(TxFechaFin)
               z = i - 1
               If Trim(rs_d.Fields(z + 4)) = "" Or Trim(rs_d.Fields(z + 4)) = "*" Or IsNull(rs_d.Fields(z + 4)) Then
               Else
                  ccadena = ccadena + Trim(Str(i)) + "_"
                  nson = nson + 1
               End If
            Next i
         End If
         If nson > 0 Then
         
            Msgvalue = MsgBox("Hay turno en DIRECTORIO los dìas: " + ccadena + "   CONTINUA ? ", vbQuestion + vbYesNo + vbDefaultButton2, "[Verifique PERIODO]")
            If Msgvalue = vbNo Then
               L_ok = False
            End If
            
         End If
         rs_d.Close
      ElseIf Month(TxfechaIni) <> Month(TxFechaFin) Or Month(TxFechaFin) <> Val(Txmes) Then
             nmesI = Month(TxfechaIni)
             nmesF = Month(TxFechaFin)
             nmesA = 0
             ndifdias = DateValue(TxFechaFin) - DateValue(TxfechaIni) + 1
             For j = 0 To ndifdias - 1
                unafecha = DateValue(TxfechaIni) + j
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
                   '''-  MsgBox sql + "->" + Str(ndifdias)
                   Set rs_d = New Recordset
                   rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
                   'Verifica si existen Registros previos
                   nson = 0
                   If Not rs_d.EOF Then
                      '''For i = Day(TxfechaIni) To Day(TxFechaFin)
                      For i = Day(primer) To Day(ultimo)
                         z = i - 1
                         If Trim(rs_d.Fields(z + 4)) = "" Or Trim(rs_d.Fields(z + 4)) = "*" Or IsNull(rs_d.Fields(z + 4)) Then
                         Else
                            ccadena = ccadena + Trim(Str(i)) + "_"
                            nson = nson + 1
                         End If
                      Next i
                   End If
                   If nson > 0 Then
                      Msgvalue = MsgBox("Hay DIRECTORIO los dìas: " + ccadena + "  CONTINUA ? ", vbQuestion + vbYesNo + vbDefaultButton2, "[APLICA DIRECTORIO]")
                      If Msgvalue = vbNo Then
                         L_ok = False
                      End If
                   End If
                End If
             Next j
      End If
      If Not L_ok Then
         Call Deshabilita
         Call Limpiacampos
         'DataGrid1.SetFocus
         Txtarjeta.Enabled = False
         Sw = False
         Exit Sub
      End If
   End If
   
   fecha0 = Format(Txfecha, "yyyymmdd")
'''------   On Error GoTo final
 '''----  Cnx.BeginTrans
   Dim rt As Recordset
   Set rt = New Recordset
   sq2 = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' and fecha='" & fecha0 & "' and fechaIni='" & fecha1 & "' and codigo='" & Txcodigo & "'"
   rt.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If Sw And rt.EOF Then
      rt.AddNew
      rt!tarjeta = Txtarjeta
      rt!fecha = Txfecha
      rt!Codigo = Txcodigo
      If IsDate(TxfechaIni) Then rt!fechaIni = TxfechaIni
   Else
      MsgBox "Existe Autorización... ", vbCritical, "[VERIFICAR fechas]"
      rt.Close
      Exit Sub
   End If
   rt!diahora = TxDiaHora
   If IsDate(TxFechaFin) Then rt!fechaFin = TxFechaFin
   If Trim(TxDiaHora) = "D" Then
      rt!horaini = ""
      rt!horafin = ""
      rt!tipo = ""
   Else
      rt!horaini = IIf(IsNull(Txhoraini), "", Txhoraini)
      rt!horafin = IIf(IsNull(Txhorafin), "", Txhorafin)
      rt!tipo = IIf(IsNull(TxTipo), "", TxTipo)
   End If
   rt!documento = IIf(IsNull(TxDocumento), "", TxDocumento)
   rt!turno = IIf(IsNull(TxTurno), "", TxTurno)
   rt!hora = Format(Time, "hh:mm:ss")
   rt!usuario = wNick
   rt!estacion = w_clienName
   rt.Update
   rt.Close
   If Sw Then
      DD.AddNew
      DD!tarjeta = Txtarjeta
      DD!fecha = Txfecha
   End If
   DD!Codigo = Txcodigo
   DD!diahora = TxDiaHora
   If IsDate(TxfechaIni) Then DD!fechaIni = TxfechaIni
   If IsDate(TxFechaFin) Then DD!fechaFin = TxFechaFin
   If Trim(TxDiaHora) = "D" Then
      DD!horaini = ""
      DD!horafin = ""
      DD!tipo = ""
   Else
      DD!horaini = IIf(IsNull(Txhoraini), "", Txhoraini)
      DD!horafin = IIf(IsNull(Txhorafin), "", Txhorafin)
      DD!tipo = IIf(IsNull(TxTipo), "", TxTipo)
   End If
   DD!documento = IIf(IsNull(TxDocumento), "", TxDocumento)
   DD!turno = IIf(IsNull(TxTurno), "", TxTurno)
   DD.Update
   '-DIA, procesa Directorio
   If Trim(TxDiaHora) = "D" And Trim(TxTurno) <> "" Then
      '-Programa cambio Turno en Directorio  ==> tarjeta,txmes,txfechaini,txfechafin
      If Month(TxfechaIni) = Val(Txmes) And Month(TxfechaIni) = Month(TxFechaFin) Then
         sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
         Set rs_d = New Recordset
         rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If rs_d.EOF Then
            Sw = True
            rs_d.AddNew
            rs_d!tarjeta = Txtarjeta
            rs_d!ano = genaño
            rs_d!mes = Txmes
            rs_d!usuario = wNick      'Firma Usuario
            rs_d!fecha = Date
            rs_d!hora = Format(Time, "hh:mm:ss")
            rs_d!estacion = w_clienName
         Else
            rs_d!usuarioM = wNick
            rs_d!fechaM = Date
         End If
         TxHoras = "": TxD = "": TxN = "": TxM = "": Txt = "":  TxMT = "": TxNa = "": TxPA = ""
         nhoras = 0: nd = 0: nn = 0: nm = 0: nt = 0: nmt = 0: nna = 0: nPA = 0
         For i = Day(TxfechaIni) To Day(TxFechaFin)
             z = i - 1
             If Trim(rs_d.Fields(z + 4)) = "" Or Trim(rs_d.Fields(z + 4)) = "*" Or IsNull(rs_d.Fields(z + 4)) Then
                '--Libre, Registra vacación
                rs_d.Fields(z + 4) = TxTurno
             Else
             
               '--->>OJO NO RETIRA PTOGRAMACION ---
               lDejaProg = False
               If lDejaProg Then
                '--->> OJO: RETIRA horario PROGRAMADO ¿? <<---
                fecha1 = Format(DateValue(TxfechaIni) + z, "yyyymmdd")
                sql1 = "UPDATE CpHor" & genaño & " SET Horario='O' ,fecregistro='" + Format(Date, "yyyymmdd") + "',Horregistro='" + Format(Time, "hh:mm:ss") + "',usuario='" + wNick + "',estacion='" + w_clienName + "' WHERE tarjeta='" + Txtarjeta & "' and fecha='" & fecha1 & "'"
                Cnx.Execute sql1
                
                rs_d.Fields(z + 4) = TxTurno
             
                If Trim(rs_d.Fields(z + 4)) = "D" Or InStr("D*D1D2D3D4", Trim(rs_d.Fields(z + 4))) > 0 Or Trim(rs_d.Fields(z + 4)) = "MT" Or Trim(rs_d.Fields(z + 4)) = "MX" Then
                   nd = nd + 1
                End If
                If Trim(rs_d.Fields(z + 4)) = "N" Or InStr("N*N2N3N4", Trim(rs_d.Fields(z + 4))) > 0 Then
                    nn = nn + 1
                End If
                If Trim(rs_d.Fields(z + 4)) = "M" Or InStr("M1M2M3", Trim(rs_d.Fields(z + 4))) > 0 Then
                    nm = nm + 1
                End If
                If Trim(rs_d.Fields(z + 4)) = "T" Or InStr("T1T2", Trim(rs_d.Fields(z + 4))) > 0 Then
                   nt = nt + 1
                End If
                If Trim(rs_d.Fields(z + 4)) = "H" Or InStr("H*H1", Trim(rs_d.Fields(z + 4))) > 0 Then
                   nm = nm + 1
                   nn = nn + 1
                End If
                If Trim(rs_d.Fields(z + 4)) = "C" Or InStr("C*C1", Trim(rs_d.Fields(z + 4))) > 0 Then
                   nt = nt + 1
                   nn = nn + 1
                End If
                If Trim(rs_d.Fields(z + 4)) = "P" Or InStr("P*P1", Trim(rs_d.Fields(z + 4))) > 0 Then
                   nd = nd + 1
                   nn = nn + 1
                End If
                If Trim(rs_d.Fields(z + 4)) = "MT" Or Trim(rs_d.Fields(z + 4)) = "MX" Then
                   nmt = nmt + 1
                End If
                If Trim(rs_d.Fields(z + 4)) = "NA" Then
                   nPA = nna + 1
                End If
                If Trim(rs_d.Fields(z + 4)) = "PA" Then
                   nPA = nPA + 1
                End If
               End If
             End If
             
         Next i
         If nd + nn + nm + nt + nmt + nna + nPA > 0 Then
            TxD = IIf(IsNull(rs_d!d), 0, Val(rs_d!d)) - nd
            TxN = IIf(IsNull(rs_d!n), 0, Val(rs_d!n)) - nn
            TxM = IIf(IsNull(rs_d!m), 0, Val(rs_d!m)) - nm
            Txt = IIf(IsNull(rs_d!T), 0, Val(rs_d!T)) - nt
            TxMT = IIf(IsNull(rs_d!mT), 0, Val(rs_d!mT)) - nmt
            TxNa = IIf(IsNull(rs_d!na), 0, Val(rs_d!na)) - nna
            TxPA = IIf(IsNull(rs_d!pa), 0, Val(rs_d!pa)) - nPA
            
            TxHoras = IIf(IsNull(rs_d!Horas), 0, rs_d!Horas) - (nPA * 24 + (nd + nn + nmt + nna) * 12 + (nm + nt) * 6)
            rs_d!d = Trim(TxD)
            rs_d!n = Trim(TxN)
            rs_d!m = Trim(TxM)
            rs_d!T = Trim(Txt)
            rs_d!mT = Trim(TxMT)
            rs_d!na = Trim(TxNa)
            rs_d!pa = Trim(TxPA)
            rs_d!Horas = Val(TxHoras)
         End If
         rs_d!carEst = IIf(IsNull(TxCarEst), "", TxCarEst)
         rs_d!CCosto = IIf(IsNull(Txccosto), "", Txccosto)
         rs_d!CodUpr = IIf(IsNull(TxCodUpR), "", TxCodUpR)
         rs_d!Cargo = IIf(IsNull(TxCargo), "", TxCargo)
         rs_d!PLH = IIf(IsNull(TxPlH), "", TxPlH.Text)
         rs_d!Flag = IIf(TxTipoPla = "ADM", "1", IIf(TxTipoPla = "ASI", "0", "0"))   'administrativo/Asistencial
         rs_d.Update
         rs_d.Close
         
      ElseIf Month(TxfechaIni) <> Month(TxFechaFin) Or Month(TxFechaFin) <> Val(Txmes) Then
             nmesI = Month(TxfechaIni)
             nmesF = Month(TxFechaFin)
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
                      rs_d.AddNew
                      rs_d!tarjeta = Txtarjeta
                      rs_d!ano = cano
                      rs_d!mes = cmes
                      rs_d!usuario = wNick      'Firma Usuario
                      rs_d!fecha = Date
                      rs_d!hora = Format(Time, "hh:mm:ss")
                      rs_d!estacion = w_clienName
                   Else
                      rs_d!usuarioM = wNick
                      rs_d!fechaM = Date
                   End If
                   TxHoras = "": TxD = "": TxN = "": TxM = "": Txt = "": TxMT = "": TxNa = "": TxPA = ""
                   nhoras = 0: nd = 0: nn = 0: nm = 0: nt = 0: nmt = 0: nna = 0: nPA = 0
                   nson = 0
                   For i = Day(primer) To Day(ultimo)
                       z = i - 1
                       nson = nson + 1
                       If Trim(rs_d.Fields(z + 4)) = "" Or Trim(rs_d.Fields(z + 4)) = "*" Or IsNull(rs_d.Fields(z + 4)) Then
                          '---REGISTRA VACACIÒN ---
                          rs_d.Fields(z + 4) = TxTurno
                       Else
                          
                         '----OJO deja PROGRAMA---
                         lDejaProg = False
                         If lDejaProg Then
                          '--->> OJO: RETIRA horario PROGRAMADO ¿? <<---
                          fecha1 = Format(primer + z, "yyyymmdd")
                          sql1 = "UPDATE CpHor" & genaño & " SET Horario='O',fecregistro='" + Format(Date, "yyyymmdd") + "',Horregistro='" + Format(Time, "hh:mm:ss") + "',usuario='" + wNick + "',estacion='" + w_clienName + "' WHERE tarjeta='" + Txtarjeta & "' and fecha='" & fecha1 & "'"
                          Cnx.Execute sql1
                
                          rs_d.Fields(z + 4) = TxTurno
                          
                          If Trim(rs_d.Fields(z + 4)) = "D" Or InStr("D*D1D2D3D4", Trim(rs_d.Fields(z + 4))) > 0 Or Trim(rs_d.Fields(z + 4)) = "MT" Or Trim(rs_d.Fields(z + 4)) = "MX" Then
                             nd = nd + 1
                          End If
                          If Trim(rs_d.Fields(z + 4)) = "N" Or InStr("N*N2N3N4", Trim(rs_d.Fields(z + 4))) > 0 Then
                             nn = nn + 1
                          End If
                          If Trim(rs_d.Fields(z + 4)) = "MT" Or Trim(rs_d.Fields(z + 4)) = "MX" Then
                             nmt = nmt + 1
                          End If
                          If Trim(rs_d.Fields(z + 4)) = "NA" Then
                             nna = nna + 1
                          End If
                          If Trim(rs_d.Fields(z + 4)) = "PA" Then
                             nPA = nPA + 1
                          End If
                          If Trim(rs_d.Fields(z + 4)) = "M" Or InStr("M1M2M3", Trim(rs_d.Fields(z + 4))) > 0 Then
                             nm = nm + 1
                          End If
                          If Trim(rs_d.Fields(z + 4)) = "T" Or InStr("T1T2", Trim(rs_d.Fields(z + 4))) > 0 Then
                             nt = nt + 1
                          End If
                          If Trim(rs_d.Fields(z + 4)) = "H" Or InStr("H*H1", Trim(rs_d.Fields(z + 4))) > 0 Then
                             nm = nm + 1
                             nn = nn + 1
                          End If
                          If Trim(rs_d.Fields(z + 4)) = "C" Or InStr("C*C1", Trim(rs_d.Fields(z + 4))) > 0 Then
                             nt = nt + 1
                             nn = nn + 1
                          End If
                          If Trim(rs_d.Fields(z + 4)) = "P" Or InStr("P*P1", Trim(rs_d.Fields(z + 4))) > 0 Then
                             nd = nd + 1
                             nn = nn + 1
                          End If
                         End If
                       End If
                   Next i
                   j = j + (nson - 1)
                   If nd + nn + nm + nt + nmt + nna + nPA > 0 Then
                      TxD = IIf(IsNull(rs_d!d), 0, Val(rs_d!d)) - nd
                      TxN = IIf(IsNull(rs_d!n), 0, Val(rs_d!n)) - nn
                      TxM = IIf(IsNull(rs_d!m), 0, Val(rs_d!m)) - nm
                      Txt = IIf(IsNull(rs_d!T), 0, Val(rs_d!T)) - nt
                      TxMT = IIf(IsNull(rs_d!mT), 0, Val(rs_d!mT)) - nmt
                      TxNa = IIf(IsNull(rs_d!na), 0, Val(rs_d!na)) - nna
                      TxPA = IIf(IsNull(rs_d!pa), 0, Val(rs_d!pa)) - nPA
                      
                      TxHoras = IIf(IsNull(rs_d!Horas), 0, rs_d!Horas) - (pa * 24 + (nd + nn + nmt + nna) * 12 + (nm + nt) * 6)
                      rs_d!d = Trim(TxD)
                      rs_d!n = Trim(TxN)
                      rs_d!m = Trim(TxM)
                      rs_d!T = Trim(Txt)
                      rs_d!mT = Trim(TxMT)
                      rs_d!na = Trim(TxNa)
                      rs_d!pa = Trim(TxPA)
                      rs_d!Horas = Val(TxHoras)
                   End If
                   rs_d!carEst = IIf(IsNull(TxCarEst), "", TxCarEst)
                   rs_d!CCosto = IIf(IsNull(Txccosto), "", Txccosto)
                   rs_d!CodUpr = IIf(IsNull(TxCodUpR), "", TxCodUpR)
                   rs_d!Cargo = IIf(IsNull(TxCargo), "", TxCargo)
                   rs_d!PLH = IIf(IsNull(TxPlH), "", TxPlH)
                   rs_d!Flag = IIf(TxTipoPla = "ADM", "1", IIf(TxTipoPla = "ASI", "0", "0"))   'administrativo/Asistencial
                   rs_d.Update
                   rs_d.Close
                End If
             Next j
      End If
   End If
'---   Cnx.CommitTrans
   
   L_ok = True
   If Trim(TxDiaHora) = "D" And Trim(TxTurno) <> "" Then
      'Valida Cambio Turno en Directorio  ==> tarjeta,txfecha,txfechaini,txfechafin
      If Month(TxfechaIni) = Val(Txmes) And Month(TxfechaIni) = Month(TxFechaFin) Then
         sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
         Set rs_d = New Recordset
         rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         For i = Day(TxfechaIni) To Day(TxFechaFin)
            z = i - 1
            If Trim(rs_d.Fields(z + 4)) = TxTurno Then
            Else
               '--->>PERMITE VACACIÒN en día PROGRAMADO <<--- L_ok = False
            End If
         Next
         rs_d.Close
      ElseIf Month(TxfechaIni) <> Month(TxFechaFin) Or Month(TxFechaFin) = Val(Txmes) Then
             nmesI = Month(TxfechaIni)
             nmesF = Month(TxFechaFin)
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
                   nson = 0
                   For i = Day(primer) To Day(ultimo)
                      z = i - 1
                      If Trim(rs_d.Fields(z + 4)) = TxTurno Then
                      Else
                         '--->>PERMITE VACACIÒN en dia PROGRAMADO  <<--- L_ok = False
                      End If
                   Next
                   j = j + (nson - 1)
                End If
             Next j
      End If
      If L_ok Then
      Else
         MsgBox "<E> Transaccion Invalida...Procese nuevamente.  ", vbExclamation
      End If
   End If
   GoTo Final1
'''final:
'''   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
'''   Cnx.RollbackTrans
Final1:
   Call Deshabilita
   Call Limpiacampos
   'DataGrid1.SetFocus
   Txtarjeta.Enabled = False
   Sw = False
End Sub
Private Sub CmdElimina_Click()
    If wrights = "CON" Then
       MsgBox "No Autorizado...  ", vbExclamation
       Exit Sub
    End If
    If Trim(Txtarjeta) = "" Then Exit Sub
    If Txmes <> "" Then
       If fecCierre <= Date Then      'And wDerEspec <> "S"
          MsgBox "Programa CERRADO...  " + Format(fecCierre, "dd/mm/yyyy"), vbExclamation
          Exit Sub
       End If
    End If
    
    fecha0 = Format(Txfecha, "yyyymmdd")
    fecha1 = Format(TxfechaIni, "yyyymmdd")
    fecha2 = Format(TxFechaFin, "yyyymmdd")
    If MsgBox("Eliminar  registro ?      ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       If TxDiaHora = "D" Then
          'Cnx.Execute "DELETE FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "'"
          cvalor = "UPDATE CpAutori" & genaño & " SET anulado='X',usuAnula='" + wNick + "',fecAnula='" + Format(Date, "yyyymmdd") + "',horAnula='" + Format(Time, "hh:mm:ss") + "',estAnula='" + w_clienName + "' "
          sql = cvalor + " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "'"
          If Year(TxFechaFin) > Val(genaño) Then
             Cnx.Execute (sql)
             cano = Trim(Str(Year(TxFechaFin)))
             'Cnx.Execute "DELETE FROM CpAutori" & cano & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "'"
             cvalor = "UPDATE CpAutori" & cano & " SET anulado='X',usuAnula='" + wNick + "',fecAnula='" + Format(Date, "yyyymmdd") + "',horAnula='" + Format(Time, "hh:mm:ss") + "',estAnula='" + w_clienName + "' "
             sql = cvalor + "  WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "'"
          ElseIf Year(Txfecha) < Val(genaño) Then
                 Cnx.Execute (sql)
                 cano = Trim(Str(Year(Txfecha)))
                 'Cnx.Execute "DELETE FROM CpAutori" & cano & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "'"
                 cvalor = "UPDATE CpAutori" & cano & " SET anulado='X',usuAnula='" + wNick + "',fecAnula='" + Format(Date, "yyyymmdd") + "',horAnula='" + Format(Time, "hh:mm:ss") + "',estAnula='" + w_clienName + "' "
                 sql = cvalor + " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "'"
          End If
          Cnx.Execute (sql)
       ElseIf TxDiaHora = "H" Then
              'Cnx.Execute "DELETE FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "' AND HoraIni='" & Txhoraini & "'AND HoraFin='" & Txhorafin & "'"
              cvalor = "UPDATE CpAutori" & genaño & " SET anulado='X',usuAnula='" + wNick + "',fecAnula='" + Format(Date, "yyyymmdd") + "',horAnula='" + Format(Time, "hh:mm:ss") + "',estAnula='" + w_clienName + "' "
              sql = cvalor + " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "' AND HoraIni='" & Txhoraini & "'AND HoraFin='" & Txhorafin & "'"
              If Year(TxFechaFin) > Val(genaño) Then
                 Cnx.Execute (sql)
                 cano = Trim(Str(Year(TxFechaFin)))
                 'Cnx.Execute "DELETE FROM CpAutori" & cano & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaINI='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "' AND HoraIni='" & Txhoraini & "'AND HoraFin='" & Txhorafin & "'"
                 cvalor = "UPDATE CpAutori" & cano & " SET anulado='X',usuAnula='" + wNick + "',fecAnula='" + Format(Date, "yyyymmdd") + "',horAnula='" + Format(Time, "hh:mm:ss") + "',estAnula='" + w_clienName + "' "
                 sql = cvalor + " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaINI='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "' AND HoraIni='" & Txhoraini & "'AND HoraFin='" & Txhorafin & "'"
              ElseIf Year(Txfecha) < Val(genaño) Then
                     Cnx.Execute (sql)
                     cano = Trim(Str(Year(Txfecha)))
                     'Cnx.Execute "DELETE FROM CpAutori" & cano & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaINI='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "' AND HoraIni='" & Txhoraini & "'AND HoraFin='" & Txhorafin & "'"
                     cvalor = "UPDATE CpAutori" & cano & " SET anulado='X',usuAnula='" + wNick + "',fecAnula='" + Format(Date, "yyyymmdd") + "',horAnula='" + Format(Time, "hh:mm:ss") + "',estAnula='" + w_clienName + "' "
                     sql = cvalor + " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaINI='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "' AND HoraIni='" & Txhoraini & "'AND HoraFin='" & Txhorafin & "'"
              End If
              Cnx.Execute (sql)
      End If
      
      
       DD.Delete
       If Trim(TxTurno) <> "" Then
          '-Retira Autorizacion, Turno en Directorio mismo año de AFECTACION
          If Month(TxfechaIni) = Val(Txmes) And Month(TxfechaIni) = Month(TxFechaFin) Then
             sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
             Set rs_d = New Recordset
             rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
             If Not rs_d.EOF Then
                For i = Day(TxfechaIni) To Day(TxFechaFin)
                   z = i - 1
                   'if Trim(rs_D.Fields(z + 4)) = Trim(TxTurno.Text) Then
                   rs_d.Fields(z + 4) = ""
                   'End If
                Next i
                rs_d.Update
             End If
             rs_d.Close
          ElseIf Month(TxfechaIni) <> Month(TxFechaFin) Or Month(TxFechaFin) <> Val(genaño) Then
                 nmesI = Month(TxfechaIni)
                 nmesF = Month(TxFechaFin)
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
                       'If Trim(rs_D.Fields(z + 4)) = Trim(TxTurno.Text) Then
                       rs_d.Fields(z + 4) = ""
                       'End If
                    Next i
                    j = j + (nson - 1)
                    rs_d.Update
                    rs_d.Close
                 Next j
          End If
       End If
       
       Call CmdChgReg_Click
       
    End If
    Call Deshabilita
    Call Limpiacampos
End Sub
Private Sub cmdsalir_Click()
  Set DD = Nothing
  Set rs_d = Nothing
  Set rs = Nothing
  Me.Hide
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set DD = Nothing
  Set rs_d = Nothing
  Set rs = Nothing
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


Private Sub Cbo_DiaHra_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Cbo_diaHra.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbo_DiaHra_LostFocus()
   '-OUT
   If IsEmpty(Cbo_diaHra.Text) Then
   Else
      j = Cbo_diaHra.ListIndex
      If j <> -1 Then
         nposicion = InStr(Cbo_diaHra.List(j), ":")
         nlen = Len(Cbo_diaHra.List(j))
         Tx_diaHra = Trim(Right(Cbo_diaHra.List(j), nlen - nposicion))
      End If
      Exit Sub
   End If
End Sub


Private Sub CmdPrint_Click()
   nson = 0
   If IsDate(TxDia) Then
      c_periodo = "(" + Format(TxDia, "dd/mm/yyyy") + ")"
      CargaAutoriza (1)
   ElseIf Trim(Txmes) <> "" Then
          c_periodo = "(" + Trim(Mid(cmeses, (Val(Txmes) - 1) * 9 + 1, 9)) + ")"
          CargaAutoriza (0)
   End If
   If nson = 0 Then
      MsgBox "Seleccion, no registra autorizaciones..", vbExclamation
      Exit Sub
   End If
   W_lbldh = "Total"
   nTipRpt = 3
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Exit Sub
End Sub






Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Txtarjeta) = "" Or (Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta)) Then
         If Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta) Then
            Ayuda1.TxUbiAP = Trim(Txtarjeta)
         End If
         Load Ayuda1
         Ayuda1.Tag = 4.1
         Ayuda1.Show 1
      Else
         Set rt = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
         rt.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rt.EOF Then
            MsgBox "No Ubicado", vbExclamation
            Txtarjeta.Text = ""
         Else
            Txtrabajador = Left(Trim(rt!apelpat) + " " + Trim(rt!apelmat) + " " + Trim(rt!nombres), 35)
            TxCarEst = rt!carEst
            TxCargo = IIf(IsNull(rt!Cargo), "", rt!Cargo)
            Txccosto = IIf(IsNull(rt!CCosto), "", rt!CCosto)
            TxCodUpR = IIf(IsNull(rt!CodUpr), "", rt!CodUpr)
            TxPlH = IIf(IsNull(rt!PLH), "", rt!PLH)
            TxTipoPla = IIf(IsNull(rt!tipoPla), "", rt!tipoPla)
         End If
         rt.Close
      End If
      CboOcurrencia.SetFocus
      Exit Sub
   End If
End Sub
Private Sub CboOcurrencia_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboOcurrencia.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboOcurrencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If IsEmpty(CboOcurrencia.Text) Then
   Else
      j = CboOcurrencia.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboOcurrencia.List(j), ":")
         nlen = Len(CboOcurrencia.List(j))
         Txcodigo = Trim(Right(CboOcurrencia.List(j), nlen - nposicion))
         
         Set rs = New Recordset
         Sq0 = "SELECT * FROM Cpocurrencias where codigo='" + Txcodigo + "'"
         rs.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then
            TxTurno = IIf(IsNull(rs!turnodirectorio), "", Trim(rs!turnodirectorio))
            TxDiaHora = "D"
            For i = 0 To CboDH.ListCount - 1
              npos = InStr(CboDH.List(i), ":")
              nlen = Len(CboDH.List(i))
              dato = Trim(Right(CboDH.List(i), nlen - npos))
              If dato = Trim(TxDiaHora) Then
                 CboDH.ListIndex = i
                 Exit For
              End If
            Next i
            Call CboDH_KeyPress(13)
         End If
         rs.Close
         If TxTurno <> "" Then
            For i = 0 To CboTurno.ListCount - 1
              npos = InStr(CboTurno.List(i), ":")
              nlen = Len(CboTurno.List(i))
              dato = Trim(Right(CboTurno.List(i), nlen - npos))
              If dato = Trim(TxTurno) Then
                 CboTurno.ListIndex = i
                 Exit For
              End If
            Next i
         Else
            CboTurno.Text = ""
            CboDH.Text = "": TxDiaHora = ""
         End If
         
      End If
      Exit Sub
   End If
 End If
End Sub
Private Sub CboOcurrencia_LostFocus()
   CboOcurrencia_KeyPress (13)
End Sub
Private Sub CboDH_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboDH.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboDH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If IsEmpty(CboDH.Text) Then
   Else
      j = CboDH.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboDH.List(j), ":")
         nlen = Len(CboDH.List(j))
         TxDiaHora = Trim(Right(CboDH.List(j), nlen - nposicion))
         If Trim(TxDiaHora) = "D" Then
            Option2(0).Value = False: Option2(1).Value = False: Option2(2).Value = False
            Txhoraini = "": Txhorafin = ""
            FrameH.Enabled = False
         ElseIf Trim(TxDiaHora) = "H" Then
                TxTurno = "": CboTurno.Text = ""
                FrameD.Enabled = False
         End If
      End If
      Exit Sub
   End If
 End If
End Sub
Private Sub CboDH_LostFocus()
   Call CboDH_KeyPress(13)
   
End Sub
Private Sub Txfechaini_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsDate(TxfechaIni.Text) Then Exit Sub
      If Year(TxfechaIni.Text) < Val(genaño) Then
         MsgBox "Periodo Cerrado...", vbCritical
         Exit Sub
      End If
      If Year(TxfechaIni.Text) > Val(genaño) Then
         If MsgBox("Verifique AÑO... Debe haber Tablas, Prosigue ?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            c_ano = Left(Format(DateValue(TxfechaIni), "yyyymmdd"), 4)
            NvaTblaAU = "CpAutori" & c_ano
            Dim rs As Recordset
            Set rs = New ADODB.Recordset
            Sq0 = "SELECT * FROM " & NvaTblaAU
            On Error GoTo nohay
            rs.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            rs.Close
            GoTo Sale
nohay:
            MsgBox "Acción no Permitida...", vbCritical
            TxfechaIni.SetFocus
            Exit Sub
Sale:
         Else
            TxfechaIni = ""
            TxfechaIni.SetFocus
            Exit Sub
         End If
         
      End If
      If IsEmpty(TxFechaFin.Text) Or TxFechaFin.Text = "" Then
         TxFechaFin.Text = TxfechaIni.Text
         TxFechaFin.SetFocus
      End If
      Exit Sub
   End If
End Sub
Private Sub TxfechaIni_LostFocus()
   If Not IsDate(TxfechaIni.Text) Then Exit Sub
   If Year(TxfechaIni.Text) < Val(genaño) Then
         MsgBox "Periodo Cerrado...", vbCritical
         Exit Sub
   End If
   If Year(TxfechaIni.Text) > Val(genaño) Then
         If MsgBox("Verifique AÑO... Debe haber Tablas, Prosigue ?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            c_ano = Left(Format(DateValue(TxfechaIni), "yyyymmdd"), 4)
            NvaTblaAU = "CpAutori" & c_ano
            Dim rs As Recordset
            Set rs = New ADODB.Recordset
            Sq0 = "SELECT * FROM " & NvaTblaAU
            On Error GoTo nohay
            rs.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            rs.Close
            GoTo Sale
nohay:
            MsgBox "Acción no permitida...", vbCritical
            TxfechaIni.SetFocus
            Exit Sub
Sale:
         Else
            TxfechaIni = ""
            TxfechaIni.SetFocus
            Exit Sub
         End If
         
   End If
   If IsEmpty(TxFechaFin.Text) Or TxFechaFin.Text = "" Then
      TxFechaFin.Text = TxfechaIni.Text
      TxFechaFin.SetFocus
   End If
   If TxDiaHora = "H" Then
      FrameH.Enabled = True
    Else
      FrameD.Enabled = True
    End If
      
   Exit Sub
End Sub
Private Sub TxfechaFin_LostFocus()
   If Not IsDate(TxFechaFin.Text) Then Exit Sub
   If Year(TxFechaFin.Text) < Val(genaño) Then
      MsgBox "verifique Fecha periodo incorrecto"
      Exit Sub
   End If
   
   If TxDiaHora = "D" Then
      '-Verifica no haya Bloquea comprometido
      fecha1 = Format(TxfechaIni, "yyyymmdd")
      fecha2 = Format(TxFechaFin, "yyyymmdd")
      Set rt = New Recordset
      '-sq2 = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "'"
      sq2 = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "'  AND DiaHora='" & TxDiaHora & "' AND fechaIni<='" & fecha2 & "' AND fechaFin>='" & fecha1 & "'"
      rt.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      If Not rt.EOF Then
         Beep 500, 350
         Beep 500, 350
         MsgBox "Existe autorización ... " + rt!Codigo + ",  del : " + Format(rt!fechaIni, "dd/mm/yyyy") + "  al : " + Format(rt!fechaFin, "dd/mm/yyyy"), vbCritical, "[ VERIFIQUE ]"
         TxfechaIni = "": TxFechaFin = ""
         CboTurno.Text = ""
         TxTurno = ""
         TxfechaIni.SetFocus
         rt.Close
         Exit Sub
      End If
      rt.Close
   End If
   
   If IsEmpty(TxFechaFin) Or DateValue(TxFechaFin) < DateValue(TxfechaIni) Then
      MsgBox "Verifique fecha Final", vbExclamation
      TxfechaIni.SetFocus
   Else
      p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = ""
      m_ing = "": m_sal = "": m_ing1 = "": m_sal1 = "": cdiaant = "N": c2dias = "N"
      fechaB = Format(TxfechaIni, "yyyymmdd")
      Fechab0 = Format(DateValue(TxfechaIni) - 1, "yyyymmdd")
      FechamF = Format(DateValue(TxfechaIni) + 1, "yyyymmdd")
      sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
      Set rs = New Recordset
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         rs.Close
         sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & Fechab0 & "'"
         Set rs = New Recordset
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         cdiaant = "S"     'Rige Horario diaanterior
      End If
      If Not rs.EOF Then
         cHorario = IIf(IsNull(rs!Horario), "", rs!Horario)
         If cHorario <> "" Then
            rs.Close
            Set rs = New Recordset
            sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & cHorario & "'"
            rs.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs.EOF Then
               p_ing = IIf(IsNull(rs!ingreso), "", rs!ingreso)
               p_sal = IIf(IsNull(rs!salida), "", rs!salida)
               p_ing1 = IIf(IsNull(rs!Ingreso1), "", rs!Ingreso1)
               p_sal1 = IIf(IsNull(rs!salida1), "", rs!salida1)
               p_horas = IIf(IsNull(rs!Horas), "0000", rs!Horas)
               LblHorario = p_ing + "-" + p_sal + " / " + p_ing1 + "-" + p_sal1 + " " + IIf(cdiaant <> "", "(D/A)", "")
               If Val(p_horas) > 0 Then
                 '-Marcas en DIA Programado
                 If (p_ing > p_sal And Trim(p_sal1) = "") Or (p_ing > p_sal1 And Trim(p_ing1) <> "") Then
                    c2dias = "S"
                 End If
                 Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
                 If cdiaant = "S" Then       'ubica marcas del dia anterior
                    If c2dias = "S" Then
                       Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha >='" & Fechab0 & "' AND fecha <='" & fechaB & "' ORDER BY fecha,hora"
                    End If
                 Else
                    If c2dias = "S" Then
                       Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha >='" & fechaB & "' AND fecha <='" & FechamF & "' ORDER BY fecha,hora"
                    End If
                 End If
                 Set rs_d = New Recordset
                 rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                 If Not rs_d.EOF Then
                    '-ASIGNA marcas basado en Horario
                    Do While Not rs_d.EOF
                       If Trim(m_ing) = "" Then
                           m_ing = rs_d!hora
                       ElseIf Trim(m_sal) = "" Then
                               m_sal = rs_d!hora
                       ElseIf Trim(m_ing1) = "" Then
                              m_ing1 = rs_d!hora
                       ElseIf Trim(m_sal1) = "" Then
                              m_sal1 = rs_d!hora
                       End If
                       rs_d.MoveNext
                    Loop
                 End If
                 rs_d.Close
                 LblMarcas = "Marcas de reloj: " + m_ing + "-" + m_sal + " / " + m_ing1 + "-" + m_sal1 + " " + IIf(c2dias = "S", "(N/D)", "")
               End If
            End If
         End If
      End If
      rs.Close
   End If
   Exit Sub
End Sub

Private Sub Option2_Click(Index As Integer)
   Select Case Index
      Case 0
         TxTipo.Text = "I"
      Case 1
          TxTipo.Text = "S"
      Case 2
          TxTipo.Text = "O"
   End Select
End Sub
Private Sub TxHoraIni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Val(Txhoraini.Text) > 2400 Or Right(Txhoraini, 2) > "59" Then
        MsgBox ("Verifique formato Hora HHMM")
     Else
     
        If TxDiaHora = "H" Then
           '-Bloquea duplicado
           fecha1 = Format(TxfechaIni, "yyyymmdd")
           fecha2 = Format(TxFechaFin, "yyyymmdd")
           Set rt = New Recordset
           sq2 = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta ='" & Trim(Txtarjeta) & "' AND fechaIni='" & fecha1 & "' AND fechaFin='" & fecha2 & "' AND codigo='" & Txcodigo & "' AND DiaHora='" & TxDiaHora & "' AND HoraIni='" & Txhoraini & "'"
           rt.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
           If Not rt.EOF Then
              Beep 500, 350
              Beep 500, 350
              MsgBox "verifique.. Autorización ya existe...", vbCritical
              Txhoraini = "": Txhorafin = ""
              Txhoraini.SetFocus
              rt.Close
              Exit Sub
           End If
           rt.Close
        End If
           
        Txhorafin.SetFocus
        Exit Sub
     End If
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
       KeyAscii = 0
   End If
End Sub
Private Sub Txhoraini_LostFocus()
     TxHoraIni_KeyPress (13)
End Sub

Private Sub TxHorafin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Val(Txhorafin) > 2400 Or Right(Txhorafin, 2) > "59" Then
         MsgBox ("Verifique Hora Final HHMM ")
         Txhorafin = ""
         Exit Sub
      End If
      If Val(Txhorafin) < Val(Txhoraini) Then
         If MsgBox("Hora final no puede ser menor a hora de inicio.. Corrige ?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            Txhorafin = ""
            Exit Sub
         End If
      End If
      
         p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = ""
         'Verifica Horario programado
         fechaB = Format(TxfechaIni, "yyyymmdd")
         sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
         Set rs = New Recordset
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         cerror = "S"
         If rs.EOF Then
            'Verifica caso autoriza sobre turno de 2Dias: 1930-0730 autotiza 0430-0630
            rs.Close
            fechaB = Format(DateValue(TxfechaIni) - 1, "yyyymmdd")
            sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
            Set rs = New Recordset
            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         End If
         If Not rs.EOF Then
            cHorario = IIf(IsNull(rs!Horario), "", rs!Horario)
            If cHorario <> "" Then
               rs.Close
               Set rs = New Recordset
               sq_l = "SELECT * FROM CpHorarioNew WHERE Codigo='" & cHorario & "'"
               rs.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not rs.EOF Then
                  p_ing = IIf(IsNull(rs!ingreso), "", rs!ingreso)
                  p_sal = IIf(IsNull(rs!salida), "", rs!salida)
                  p_ing1 = IIf(IsNull(rs!Ingreso1), "", rs!Ingreso1)
                  p_sal1 = IIf(IsNull(rs!salida1), "", rs!salida1)
                  p_horas = IIf(IsNull(rs!Horas), "0000", rs!Horas)
               End If
            Else
               MsgBox "Verifique, Sin Horario..", vbExclamation
            End If
            rs.Close
            Exit Sub
         Else
            MsgBox "Verifique, No programo Horario..", vbExclamation
         End If
         rs.Close
      
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
      KeyAscii = 0
   End If
End Sub
Private Sub Txhorafin_LostFocus()
   TxHorafin_KeyPress (13)
End Sub

Private Sub Cboturno_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboTurno.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cboturno_LostFocus()
   If IsEmpty(CboTurno.Text) Then
   Else
      j = CboTurno.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboTurno.List(j), ":")
         nlen = Len(CboTurno.List(j))
         TxTurno = Right(CboTurno.List(j), nlen - nposicion)
      End If
      Exit Sub
   End If
End Sub
Private Sub Cboturno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboTurno.Text) Then
      Else
         j = CboTurno.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboTurno.List(j), ":")
            nlen = Len(CboTurno.List(j))
            TxTurno = Right(CboTurno.List(j), nlen - nposicion)
            
            'If Trim(TxTurno) = "V" Then
            '   nndias = 0
               '-conteo dias para motivo Codigo   VA17..VAaa
            '   Call ChkConteoD(Txcodigo)
            'End If
            
         End If
      End If
      Exit Sub
   End If
End Sub

Private Sub CargaAutoriza(nopc)
'***  CARGA  registros  ***   0:Mes   1:dia
    'Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmpTabla"
    Set rs_d = New Recordset
    sq_l = "SELECT * FROM CpTmptabla"
    rs_d.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    nson = 0
    If nopc = 0 Then
       fecha1 = Format(primer, "yyyymmdd")
       fecha2 = Format(ultimo, "yyyymmdd")
       If Trim(Tx_tarjeta) = "" Then
          sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and fechaini>='" & fecha1 & "' AND fechaINI<='" & fecha2 & "'  ORDER BY tarjeta,FechaIni"
          If Trim(Tx_diaHra) <> "" Then
             sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and fechaini>='" & fecha1 & "' AND fechaINI<='" & fecha2 & "' and diahora='" & Tx_diaHra.Text & "' ORDER BY tarjeta,FechaIni"
          End If
       Else
           sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta='" & Trim(Tx_tarjeta) & "' AND fechaini>='" & fecha1 & "' AND fechaINI<='" & fecha2 & "' ORDER BY tarjeta,FechaIni"
           If Trim(Tx_diaHra) <> "" Then
              sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta='" & Trim(Tx_tarjeta) & "' AND fechaini>='" & fecha1 & "' AND fechaINI<='" & fecha2 & "' and diahora='" & Tx_diaHra & "' ORDER BY tarjeta,FechaIni"
           End If
       End If
    Else
       fecha1 = Format(TxDia, "yyyymmdd")
       If Trim(Tx_tarjeta) = "" Then
          sql = "SELECT * FROM CpAutori" & genaño & "  WHERE (anulado is NULL) and fechaini<='" & fecha1 & "' AND fechaFin>='" & fecha1 & "' ORDER BY tarjeta,FechaIni"
          If Trim(Tx_diaHra) <> "" Then
             sql = "SELECT * FROM CpAutori" & genaño & "  WHERE (anulado is NULL) and fechaini<='" & fecha1 & "' AND fechaFin>='" & fecha1 & "' and DiaHora='" & Tx_diaHra & "' ORDER BY tarjeta,FechaIni"
          End If
       Else
          sql = "SELECT * FROM CpAutori" & genaño & "  WHERE (anulado is NULL) and tarjeta='" & Trim(Tx_tarjeta) & "' AND fechaini<='" & fecha1 & "' AND fechafin>='" & fecha1 & "' ORDER BY tarjeta,FechaIni"
          If Trim(Tx_diaHra) <> "" Then
             sql = "SELECT * FROM CpAutori" & genaño & "  WHERE  (anulado is NULL) and tarjeta='" & Trim(Tx_tarjeta) & "' AND fechaini<='" & fecha1 & "' AND fechafin>='" & fecha1 & "' And DiaHora='" & Tx_diaHra & "' ORDER BY tarjeta,FechaIni"
          End If
       
       
       End If
    End If
    Set r_d = New ADODB.Recordset
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not r_d.EOF
       nson = nson + 1
       rs_d.AddNew
       rs_d!tarjeta = r_d!tarjeta
       Set rs = New Recordset
       sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & r_d!tarjeta & "'"
       rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       If rs.EOF Then
       Else
          rs_d!trabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
       End If
       rs.Close
       rs_d!Codigo = IIf(IsEmpty(r_d!Codigo), "", r_d!Codigo)
       rs_d!diahora = IIf(IsEmpty(r_d!diahora), "", r_d!diahora)
       If IsDate(r_d!fechaIni) Then rs_d!fechaIni = r_d!fechaIni
       If IsDate(r_d!fechaFin) Then rs_d!fechaFin = r_d!fechaFin
       rs_d!documento = IIf(IsEmpty(r_d!documento), "", r_d!documento)
       If r_d!diahora = "D" Then
          rs_d!Valor = r_d!fechaFin - r_d!fechaIni + 1
          rs_d!turno = IIf(IsEmpty(r_d!turno), "", r_d!turno)
       End If
       m_hini = "": m_hfin = "": m_tot = 0
       If r_d!diahora = "H" Then
          rs_d!horaini = IIf(IsEmpty(r_d!horaini), "", r_d!horaini)
          rs_d!horafin = IIf(IsEmpty(r_d!horafin), "", r_d!horafin)
          m_hini = IIf(IsEmpty(r_d!horaini), "", r_d!horaini)
          m_hfin = IIf(IsEmpty(r_d!horafin), "", r_d!horafin)
          m_tot = (Val(Left(m_hfin, 2)) - Val(Left(m_hini, 2))) * 60 + Val(Right(m_hfin, 2)) - Val(Right(m_hini, 2))   'minutos
          rs_d!Valor = m_tot    'mints
       End If
       rs_d.Update
       r_d.MoveNext
    Loop
   ' Cnx.CommitTrans
    rs_d.Close
    r_d.Close
End Sub

Private Sub CargaProducc()
    'Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmpTabla"
    Set rs_d = New Recordset
    sq_l = "SELECT * FROM CpTmptabla"
    rs_d.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    nson = 0
    fecha1 = Format(DateValue(TxDia.Text), "yyyymmdd")
    sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and fecha='" & fecha1 & "' ORDER BY Usuario,tarjeta"
    Set r_d = New ADODB.Recordset
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not r_d.EOF
       nson = nson + 1
       rs_d.AddNew
       rs_d!usuario = r_d!usuario
       rs_d!fecha = r_d!fecha
       rs_d!tarjeta = r_d!tarjeta
       Set rs = New Recordset
       sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & r_d!tarjeta & "'"
       rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       If rs.EOF Then
       Else
          rs_d!trabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
       End If
       rs.Close
       rs_d!Codigo = IIf(IsEmpty(r_d!Codigo), "", r_d!Codigo)
       rs_d!diahora = IIf(IsEmpty(r_d!diahora), "", r_d!diahora)
       If IsDate(r_d!fechaIni) Then rs_d!fechaIni = r_d!fechaIni
       If IsDate(r_d!fechaFin) Then rs_d!fechaFin = r_d!fechaFin
       rs_d!documento = IIf(IsEmpty(r_d!documento), "", r_d!documento)
       If r_d!diahora = "D" Then
          rs_d!Valor = r_d!fechaFin - r_d!fechaIni + 1
          rs_d!turno = IIf(IsEmpty(r_d!turno), "", r_d!turno)
       End If
       m_hini = "": m_hfin = "": m_tot = 0
       If r_d!diahora = "H" Then
          rs_d!horaini = IIf(IsEmpty(r_d!horaini), "", r_d!horaini)
          rs_d!horafin = IIf(IsEmpty(r_d!horafin), "", r_d!horafin)
          m_hini = IIf(IsEmpty(r_d!horaini), "", r_d!horaini)
          m_hfin = IIf(IsEmpty(r_d!horafin), "", r_d!horafin)
          m_tot = (Val(Left(m_hfin, 2)) - Val(Left(m_hini, 2))) * 60 + Val(Right(m_hfin, 2)) - Val(Right(m_hini, 2))   'minutos
          rs_d!Valor = m_tot    'mints
       End If
       rs_d.Update
       r_d.MoveNext
    Loop
   ' Cnx.CommitTrans
    rs_d.Close
    r_d.Close
End Sub


