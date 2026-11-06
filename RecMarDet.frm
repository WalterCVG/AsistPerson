VERSION 5.00
Begin VB.Form RecMarDet 
   Caption         =   "Detalle Diario"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "RecMarDet.frx":0000
   ScaleHeight     =   4275
   ScaleWidth      =   11985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   11040
      Picture         =   "RecMarDet.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Cerrar"
      Top             =   3600
      Width           =   735
   End
   Begin VB.Frame frame1 
      Enabled         =   0   'False
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11535
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   4920
         TabIndex        =   28
         Top             =   960
         Width           =   6495
         Begin VB.TextBox TxMotivos 
            Height          =   285
            Left            =   5520
            TabIndex        =   41
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox Txdescuenta 
            Height          =   285
            Left            =   2040
            MaxLength       =   1
            TabIndex        =   40
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox TxHorafin 
            Height          =   285
            Left            =   4440
            MaxLength       =   4
            TabIndex        =   23
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox TxHoraini 
            Height          =   285
            Left            =   3480
            MaxLength       =   4
            TabIndex        =   22
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox Txdiahora 
            Height          =   285
            Left            =   3000
            MaxLength       =   1
            TabIndex        =   21
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox TxHrasFalt 
            Height          =   285
            Left            =   4320
            TabIndex        =   25
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox TxHrasTrab 
            Height          =   285
            Left            =   1320
            TabIndex        =   24
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox TxAutorizac 
            Height          =   285
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   20
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Txdescuento 
            Height          =   285
            Left            =   4440
            TabIndex        =   18
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox Txtardanza 
            Height          =   285
            Left            =   3480
            TabIndex        =   17
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox TxhrasProg 
            Height          =   285
            Left            =   1320
            TabIndex        =   16
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "Minutos"
            Height          =   255
            Left            =   5520
            TabIndex        =   39
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "A"
            Height          =   255
            Left            =   4440
            TabIndex        =   38
            Top             =   840
            Width           =   135
         End
         Begin VB.Label Label14 
            Caption         =   "De"
            Height          =   255
            Left            =   3480
            TabIndex        =   37
            Top             =   840
            Width           =   255
         End
         Begin VB.Label Label1 
            Caption         =   "Diahora"
            Height          =   255
            Left            =   2400
            TabIndex        =   36
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Horas  Falta"
            Height          =   375
            Left            =   3480
            TabIndex        =   35
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Horas Trabajadas"
            Height          =   495
            Left            =   240
            TabIndex        =   34
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Autorizacion"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Descuento"
            Height          =   255
            Left            =   4440
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Minutos"
            Height          =   255
            Left            =   3480
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Tardanza"
            Height          =   255
            Left            =   2400
            TabIndex        =   30
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Horas Prog."
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.TextBox Txmsal1 
         Height          =   285
         Left            =   3840
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Txming1 
         Height          =   285
         Left            =   2880
         TabIndex        =   13
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Txmsal 
         Height          =   285
         Left            =   1680
         TabIndex        =   12
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Txming 
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox Txpsal1 
         Height          =   285
         Left            =   3840
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Txping1 
         Height          =   285
         Left            =   2880
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Txturno 
         Height          =   285
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
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
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Txpsal 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Txping 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   600
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox TxHorario 
         Height          =   285
         Left            =   8400
         MaxLength       =   4
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Txgrupo 
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
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Horario"
         Height          =   255
         Left            =   7800
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Marcas Reloj"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Turno"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Programa   "
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label label12 
         Caption         =   "Grupo  "
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "RecMarDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim Sw As Boolean
Dim sql As String
Private Sub Form_Load()
  Frame1.Enabled = False
  Txfecha.Text = RecMarcasRS.Txfecha.Text
  fecha = Format(Txfecha.Text, "yyyymmdd")
  Set rs = New Recordset
  'SQL = "SELECT * FROM CpTmrecord where tarjeta='" & RecordMarcas.Txtarjeta.Text & "' and fecha='" & fecha & "'"
  '-Uso Temporal
  sql = "SELECT * FROM #tabla1 WHERE tarjeta='" & RecMarcasRS.Txtarjeta.Text & "' and fecha='" & fecha & "'"
  
  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  If Not rs.EOF Then
      TxGrupo.Text = IIf(IsNull(rs!grupo), "", rs!grupo)
      TxTurno.Text = IIf(IsNull(rs!turno), "", rs!turno)
      TxHorario.Text = IIf(IsNull(rs!Horario), "", rs!Horario)
      Txping.Text = IIf(IsNull(rs!ping), "", rs!ping)
      Txpsal.Text = IIf(IsNull(rs!psal), "", rs!psal)
      Txping1.Text = IIf(IsNull(rs!ping1), "", rs!ping1)
      Txpsal1.Text = IIf(IsNull(rs!psal1), "", rs!psal1)
      TxhrasProg.Text = 0 + rs!hrasprog
      Txming.Text = IIf(IsNull(rs!ming), "", rs!ming)
      Txmsal.Text = IIf(IsNull(rs!msal), "", rs!msal)
      Txming1.Text = IIf(IsNull(rs!ming1), "", rs!ming1)
      Txmsal1.Text = IIf(IsNull(rs!msal1), "", rs!msal1)
      Txtardanza.Text = 0 + rs!tardanza
      Txdescuento.Text = 0 + rs!DESCUENTO
      TxAutorizac.Text = IIf(IsNull(rs!autorizac), "", rs!autorizac)
      Txdescuenta.Text = "" + rs!descuenta
      TxDiaHora.Text = IIf(IsNull(rs!diahora), "", rs!diahora)
      Txhoraini.Text = IIf(IsNull(rs!horaini), "", rs!horaini)
      Txhorafin.Text = IIf(IsNull(rs!horafin), "", rs!horafin)
      TxMotivos.Text = IIf(IsNull(rs!motpersonal), rs!Comision, rs!motpersonal)
      TxHrasTrab.Text = 0 + rs!HrasTrab
      MSG = ""
      VerHrsMint (Val(TxhrasProg.Text))
      TxhrasProg.Text = MSG
      'TxTipo.Text = IIf(IsNull(rs!tipo), "", rs!tipo)
      If Val(TxHrasTrab.Text) > 0 Then
         MSG = ""
         VerHrsMint (Val(TxHrasTrab.Text))
         TxHrasTrab.Text = MSG
      End If
      If Val(Txhorafin.Text) - Val(Txhoraini.Text) > 0 Then
         MSG = ""
         VerHrsMint (Val(TxMotivos.Text))
         TxMotivos.Text = MSG
      End If
      TxHrasFalt.Text = 0 + rs!hrasfalt
  End If
  rs.Close
  
End Sub
 
Private Sub cmdsalir_Click()
  Set rs = Nothing
  Me.Hide
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  Me.Hide
  Unload Me
End Sub

