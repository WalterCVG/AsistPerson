VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form EditaMarcas 
   Caption         =   "Edicion Marcas"
   ClientHeight    =   9105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "EditaMarcas.frx":0000
   ScaleHeight     =   9105
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Limpia "
      Height          =   615
      Left            =   9720
      Picture         =   "EditaMarcas.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Limpia Plantilla"
      Top             =   7440
      Width           =   615
   End
   Begin VB.CommandButton CmdClear 
      Height          =   615
      Left            =   11160
      Picture         =   "EditaMarcas.frx":80798
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Otra Consulta"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "EditaMarcas.frx":81062
      Style           =   1  'Graphical
      TabIndex        =   35
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
      Left            =   8640
      Picture         =   "EditaMarcas.frx":81CA4
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Imprimir"
      Top             =   120
      Width           =   765
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
      Left            =   7200
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   195
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Tx_tarjeta 
      Height          =   285
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Tx_trabajador 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   5535
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   9480
      TabIndex        =   20
      Top             =   6360
      Width           =   2175
      Begin VB.CommandButton cmdGraba 
         Caption         =   "Graba"
         Height          =   615
         Left            =   1080
         Picture         =   "EditaMarcas.frx":8256E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Grabar"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdElimina 
         Caption         =   "Elimina"
         Height          =   615
         Left            =   1200
         Picture         =   "EditaMarcas.frx":826E0
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Eliminar"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   615
         Left            =   120
         Picture         =   "EditaMarcas.frx":82852
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Agregar Registro "
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   10440
      Picture         =   "EditaMarcas.frx":829C4
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Cerrar"
      Top             =   8280
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "[Registro]"
      Height          =   2175
      Left            =   240
      TabIndex        =   9
      Top             =   6240
      Width           =   8895
      Begin VB.TextBox Txobservacion 
         Height          =   285
         Left            =   5160
         MaxLength       =   40
         TabIndex        =   19
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox TxGrpGen 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7920
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxMotivo 
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Cbomotivo 
         Height          =   315
         Left            =   1200
         TabIndex        =   18
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox TxTipo 
         Enabled         =   0   'False
         Height          =   195
         Left            =   6240
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Cbotipo 
         Height          =   315
         Left            =   6720
         TabIndex        =   16
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Txhora 
         Height          =   285
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   14
         Top             =   720
         Width           =   615
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
         Left            =   840
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Txtrabajador 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   35
         TabIndex        =   11
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox Txtarjeta 
         Height          =   285
         Left            =   840
         MaxLength       =   5
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Observacion"
         Height          =   255
         Left            =   4200
         TabIndex        =   39
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "D/A : horario inicio dia anterior"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label LblHorario 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2040
         TabIndex        =   37
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Grp Gen"
         Height          =   255
         Left            =   7200
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Justificacion"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   15
         Left            =   3480
         TabIndex        =   29
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   6360
         TabIndex        =   28
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Hora"
         Height          =   255
         Left            =   4800
         TabIndex        =   27
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdchgreg 
      Height          =   615
      Left            =   9960
      Picture         =   "EditaMarcas.frx":82E06
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ubicar Registros"
      Top             =   600
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4575
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8070
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
      Caption         =   "Registros Encontrados"
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
      Left            =   8160
      TabIndex        =   40
      Top             =   8760
      Width           =   2055
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   1200
      TabIndex        =   33
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label_2 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   255
      Left            =   1200
      TabIndex        =   32
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      Height          =   255
      Left            =   6600
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "EditaMarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DD As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim Sw As Boolean
Dim nposicion, ndifdias As Integer
Dim sql, o_tarjeta As String
Dim sobtiempo, nhrFin, nhrIni As Double
Dim unafecha As Date
Private Sub Command1_Click()
     c_grupo = "02": c_secuencia = "027"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub

Private Sub Form_Load()
  Me.Caption = "Edición Marcas - " & genaño & "   -   Sesión: " + wUsuNombre
  LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
  Set rs = New Recordset
  rs.Open "SELECT * FROM cptablas WHERE tabla='32' AND codigo<>'..'  ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
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
  rs.Open "SELECT * FROM Cptablas WHERE tabla='34' and codigo<>'..' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To Cbotipo.ListCount - 1
           If Cbotipo.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          Cbotipo.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  rs.Open "SELECT * FROM Cptablas WHERE tabla='38' and codigo<>'..' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To Cbomotivo.ListCount - 1
           If Cbomotivo.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          Cbomotivo.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  Set DD = New Recordset
  With DD
     .Fields.Append "fecha", adDate, 5
     .Fields.Append "tarjeta", adChar, 5
     .Fields.Append "hora", adChar, 4
     .Fields.Append "tipo", adChar, 1
     .Fields.Append "motivo", adChar, 2
     .Fields.Append "observacion", adChar, 40
     .Fields.Append "usuario", adChar, 10
     .Fields.Append "Fecregistro", adDate
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
          Case 0
          '.Caption = "fecha"
          .Width = Me.TextWidth(String(10, "S"))
          Case 1
          '.Caption ="tarjeta"
          .Width = Me.TextWidth(String(7, "S"))
          Case 2
          '.Caption ="hora"
          .Width = Me.TextWidth(String(5, "S"))
          Case 3
          .Width = Me.TextWidth(String(4, "S"))
          Case 4
          .Width = Me.TextWidth(String(6, "S"))
          Case 5
          .Width = Me.TextWidth(String(10, "S"))
          Case 7
          .Width = Me.TextWidth(String(10, "S"))
          End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
  Next
  
End Sub
Private Sub Cbomes_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboMes.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbomes_LostFocus()
   If IsEmpty(CboMes.Text) Then
   Else
      j = CboMes.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboMes.List(j), ":")
         nlen = Len(CboMes.List(j))
         Txmes.Text = Right(CboMes.List(j), nlen - nposicion)
         unafecha = "01/" + Txmes.Text + "/" + Trim(Str(genaño))
         Call ObtenerPrimerUltimoDia(unafecha)
      End If
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
            Txmes.Text = Right(CboMes.List(j), nlen - nposicion)
            unafecha = "01/" + Txmes.Text + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)
         End If
      End If
      Exit Sub
   End If
End Sub
Private Sub Tx_tarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Tx_tarjeta.Text) = "" Or (Trim(Tx_tarjeta.Text) <> "" And EsTexto(Tx_tarjeta.Text)) Then
         If Trim(Tx_tarjeta.Text) <> "" And EsTexto(Tx_tarjeta.Text) Then
            Ayuda1.TxUbiAP = Trim(Tx_tarjeta.Text)
         End If
         Load Ayuda1
         Ayuda1.Tag = 1
         Ayuda1.Show 1
      Else
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Tx_tarjeta.Text & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "Tarjeta, No Ubicada", vbExclamation
            Tx_tarjeta.Text = ""
         Else
            Tx_trabajador.Text = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
         End If
         rs.Close
      End If
      Do While Not DD.EOF
         DD.Delete
         DD.MoveFirst
      Loop
      Set DataGrid1.DataSource = DD
      
       Call CmdChgReg_Click
       
      Exit Sub
   End If
End Sub
Private Sub CmdChgReg_Click()
   If Not IsDate(TxDia.Text) And Txmes.Text = "" Then Exit Sub
   If Not IsDate(TxDia.Text) And Tx_tarjeta.Text = "" Then Exit Sub
   If Txmes.Text <> "" And Tx_tarjeta.Text = "" Then Exit Sub
   If Trim(Txmes.Text) <> "" Then
      fecha1 = Format(primer, "yyyymmdd")
      fecha2 = Format(ultimo, "yyyymmdd")
   End If
   If IsDate(TxDia.Text) Then
      fechaB = Format(TxDia.Text, "yyyymmdd")
   End If
   If Trim(Tx_tarjeta.Text) = "" Then
      Tx_trabajador.Text = ""
      If Trim(Txmes.Text) <> "" Then
        If Not IsDate(TxDia.Text) Then
           sql = "SELECT * FROM CpMarca" & genaño & " WHERE fecha>='" & fecha1 & "' And  fecha<='" & fecha2 & "' ORDER BY fecha,hora "
        Else
           sql = "SELECT * FROM CpMarca" & genaño & " WHERE fecha='" & fechaB & "' ORDER BY fecha,hora "
        End If
      ElseIf IsDate(TxDia.Text) Then
             sql = "SELECT * FROM CpMarca" & genaño & " WHERE fecha='" & fechaB & "' ORDER BY fecha,hora "
      End If
   Else
      If Trim(Txmes.Text) <> "" Then
         If Not IsDate(TxDia.Text) Then
            sql = "SELECT * FROM CpMarca" & genaño & " WHERE fecha>='" & fecha1 & "' AND fecha<='" & fecha2 & "' AND tarjeta='" & Tx_tarjeta.Text & "' ORDER BY fecha,hora "
         Else
            sql = "SELECT * FROM CpMarca" & genaño & " WHERE fecha='" & fechaB & "' AND tarjeta='" & Tx_tarjeta.Text & "' ORDER BY fecha,hora "
         End If
      End If
   End If
   Call cargaDD
End Sub
Private Sub cargaDD()
   Do While Not DD.EOF
      DD.Delete
      DD.MoveFirst
   Loop
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If rs.EOF Then
      MsgBox "No registra Marcación...  ", vbExclamation
   Else
      rs.MoveFirst
      Do While Not rs.EOF
         DD.AddNew
         If IsDate(rs!fecha) Then DD!fecha = rs!fecha
         DD!tarjeta = IIf(IsNull(rs!tarjeta), "", rs!tarjeta)
         DD!hora = IIf(IsNull(rs!hora), "", rs!hora)
         DD!tipo = IIf(IsNull(rs!tipo), "", rs!tipo)
         DD!motivo = IIf(IsNull(rs!motivo), "", rs!motivo)
         DD!observacion = IIf(IsNull(rs!observacion), "", rs!observacion)
         DD!usuario = IIf(IsNull(rs!usuario), "", rs!usuario)
         If IsDate(rs!fecRegistro) Then DD!fecRegistro = rs!fecRegistro
         DD.Update
         rs.MoveNext
      Loop
   End If
   rs.Close
   Call Limpiacampos
End Sub
Private Sub CmdClear_Click()
  CboMes.Text = "": Txmes.Text = ""
  TxDia.Text = "":
  Tx_tarjeta.Text = "": Tx_trabajador.Text = ""
  Do While Not DD.EOF
     DD.Delete
     DD.MoveFirst
  Loop
  Set DataGrid1.DataSource = DD
End Sub

Private Sub Limpiacampos()
   Txtarjeta.Text = "": TxTrabajador.Text = ""
   TxGrpGen.Text = ""
   Txfecha.Text = "": Txhora.Text = "": TxTipo.Text = "": Cbotipo.Text = ""
   LblHorario.Caption = ""
   TxMotivo.Text = "": Cbomotivo.Text = ""
   Txobservacion.Text = "":
   'DataGrid1.SetFocus
End Sub
Private Sub Habilitacampos()
   Txtarjeta.Locked = False
   Txhora.Locked = False
   Cbotipo.Enabled = True:   TxTipo.Locked = False
   Cbomotivo.Enabled = True:   TxMotivo.Locked = False
   Txobservacion.Locked = False
End Sub

Private Sub Desabilitacampos()
   Txtarjeta.Locked = True
   Txhora.Locked = True
   Cbotipo.Enabled = False:   TxTipo.Locked = True
   Cbomotivo.Enabled = False:   TxMotivo.Locked = True
   Txobservacion.Locked = True
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
    'On Error GoTo final
    Call Limpiacampos
    'Call Desabilitacampos
    If IsDate(DataGrid1.Columns(0)) Then Txfecha.Text = DataGrid1.Columns(0)
    Txtarjeta.Text = IIf(IsEmpty(DataGrid1.Columns(1)), "", DataGrid1.Columns(1))
    Txhora.Text = IIf(IsEmpty(DataGrid1.Columns(2)), "", DataGrid1.Columns(2))
    TxTipo.Text = IIf(IsEmpty(DataGrid1.Columns(3)), "", DataGrid1.Columns(3))
    TxMotivo.Text = IIf(IsEmpty(DataGrid1.Columns(4)), "", DataGrid1.Columns(4))
    Txobservacion.Text = IIf(IsEmpty(DataGrid1.Columns(5)), "", DataGrid1.Columns(5))
    If Trim(TxTipo.Text) <> "" Then
       For i = 0 To Cbotipo.ListCount - 1
         npos = InStr(Cbotipo.List(i), ":")
         nlen = Len(Cbotipo.List(i))
         dato = Trim(Right(Cbotipo.List(i), nlen - npos))
         If dato = Trim(TxTipo.Text) Then
            Cbotipo.ListIndex = i
            Exit For
         End If
       Next i
    End If
    If Trim(TxMotivo.Text) <> "" Then
       For i = 0 To Cbomotivo.ListCount - 1
         npos = InStr(Cbomotivo.List(i), ":")
         nlen = Len(Cbomotivo.List(i))
         dato = Trim(Right(Cbomotivo.List(i), nlen - npos))
         If dato = Trim(TxMotivo.Text) Then
            Cbomotivo.ListIndex = i
            Exit For
         End If
       Next i
    End If
    If Trim(Txtarjeta.Text) <> "" Then
       Set rs = New Recordset
       rs.Open "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta='" & Txtarjeta.Text & "'", Cnx, adOpenStatic, adLockReadOnly
       If Not rs.EOF Then
          TxTrabajador.Text = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
          TxGrpGen.Text = IIf(IsNull(rs!grpgen), "", Trim(rs!grpgen))
          fecha1 = Format(Txfecha.Text, "yyyymmdd")
          Set rs_d = New Recordset
          'Valida programacion de tarjeta / grupo generico
          rs_d.Open "SELECT * FROM CpHor" & genaño & " WHERE fecha ='" & fecha1 & "' and tarjeta='" & Txtarjeta.Text & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs_d.EOF Then
             TxGrpGen.Text = ""
          End If
          rs_d.Close
       End If
       rs.Close
       Txtarjeta.Enabled = False
    
       p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": cdiaant = ""
       fechaB = Format(Txfecha, "yyyymmdd")
       sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
       Set rs = New Recordset
       rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       If rs.EOF Then
          rs.Close
          fechaB = Format(DateValue(Txfecha) - 1, "yyyymmdd")
          sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
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
                LblHorario.ForeColor = &HC00000     'azul
                If cdiaant = "S" Then
                   LblHorario.ForeColor = &HFF&  'rojo
                End If
                LblHorario = p_ing + "-" + p_sal + " / " + p_ing1 + "-" + p_sal1
             End If
          End If
       End If
    
    End If
    Sw = False
    Exit Sub
'final:
'   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
End Sub
Private Sub dataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
 If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
 '  On Error GoTo final
   If KeyCode = 13 Then
      Call Limpiacampos
      'Call Desabilitacampos
      If IsDate(DataGrid1.Columns(0)) Then Txfecha.Text = DataGrid1.Columns(0)
      Txtarjeta.Text = IIf(IsEmpty(DataGrid1.Columns(1)), "", DataGrid1.Columns(1))
      Txhora.Text = IIf(IsEmpty(DataGrid1.Columns(2)), "", DataGrid1.Columns(2))
      TxTipo.Text = IIf(IsEmpty(DataGrid1.Columns(3)), "", DataGrid1.Columns(3))
      TxMotivo.Text = IIf(IsEmpty(DataGrid1.Columns(4)), "", DataGrid1.Columns(4))
      Txobservacion.Text = IIf(IsEmpty(DataGrid1.Columns(5)), "", DataGrid1.Columns(5))
      If Trim(TxTipo.Text) <> "" Then
         For i = 0 To Cbotipo.ListCount - 1
           npos = InStr(Cbotipo.List(i), ":")
           nlen = Len(Cbotipo.List(i))
           dato = Trim(Right(Cbotipo.List(i), nlen - npos))
           If dato = Trim(TxTipo.Text) Then
              Cbotipo.ListIndex = i
              Exit For
           End If
         Next i
      End If
      If Trim(TxMotivo.Text) <> "" Then
         For i = 0 To Cbomotivo.ListCount - 1
           npos = InStr(Cbomotivo.List(i), ":")
           nlen = Len(Cbomotivo.List(i))
           dato = Trim(Right(Cbomotivo.List(i), nlen - npos))
           If dato = Trim(TxMotivo.Text) Then
              Cbomotivo.ListIndex = i
              Exit For
           End If
         Next i
      End If
      If Trim(Txtarjeta.Text) <> "" Then
         Set rs = New Recordset
         rs.Open "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta='" & Txtarjeta.Text & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs.EOF Then
            TxTrabajador.Text = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
            TxGrpGen.Text = IIf(IsNull(rs!grpgen), "", rs!grpgen)
            fecha1 = Format(Txfecha.Text, "yyyymmdd")
            Set rs_d = New Recordset
            rs_d.Open "SELECT * FROM CpHor" & genaño & " WHERE fecha ='" & fecha1 & "' and tarjeta='" & Txtarjeta.Text & "'", Cnx, adOpenStatic, adLockReadOnly
            If Not rs_d.EOF Then
               TxGrpGen.Text = ""
            End If
            rs_d.Close
         End If
         rs.Close
         
         p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": cdiaant = ""
         fechaB = Format(Txfecha, "yyyymmdd")
         sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
         Set rs = New Recordset
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            rs.Close
            fechaB = Format(DateValue(Txfecha) - 1, "yyyymmdd")
            sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
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
                  LblHorario.ForeColor = &HC00000     'azul
                  If cdiaant = "S" Then
                     LblHorario.ForeColor = &HFF&  'rojo
                  End If
                  LblHorario = p_ing + "-" + p_sal + " / " + p_ing1 + "-" + p_sal1
               End If
            End If
         End If
         
      End If
      Txtarjeta.Enabled = False
      Sw = False
      Exit Sub
   End If
'final:
'   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
End Sub
Private Sub cmdNuevo_Click()
   Call Limpiacampos
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If Trim(Tx_tarjeta.Text) <> "" Then
      Txtarjeta.Text = Tx_tarjeta.Text
      Txtarjeta_KeyPress (13)
   Else
      Txtarjeta.SetFocus
   End If
   Sw = True
End Sub
Private Sub cmdGraba_Click()
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If Trim(Txtarjeta) = "" Then
      MsgBox "Ingrese Nro tarjeta...  ", vbExclamation
      Exit Sub
   Else
   
      If Not IsDate(Txfecha) Then
         MsgBox "Verifique fecha...  ", vbExclamation
         Exit Sub
      End If
      If Year(Txfecha) < Val(genaño) Then
         MsgBox "Verifique fecha...  ", vbExclamation
         Exit Sub
      End If
   
      Set rs = New Recordset
      sql = "SELECT * FROM CpGeneral  WHERE  (baja is NULL) and tarjeta='" & Txtarjeta & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "Error, Tarjeta Ya existe..! ", vbExclamation
         rs.Close
         Exit Sub
      Else
         rs.Close
      End If
   End If
   If Not IsDate(Txfecha) Then
      MsgBox "Verifique fecha...  ", vbExclamation
      Exit Sub
   End If
   If Trim(TxTipo) = "" Then
      MsgBox "Defina Tipo", vbExclamation
      Exit Sub
   End If
   If Trim(TxMotivo) = "" Then
      MsgBox "Indique Motivo de edición..", vbExclamation
      Exit Sub
   End If
   fecha1 = Format(Txfecha, "yyyymmdd")
   On Error GoTo final
   Cnx.BeginTrans
   Set rt = New Recordset
   sq2 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta ='" & Trim(Txtarjeta) & "' and fecha='" & fecha1 & "' and hora='" & Txhora.Text & "'"
   rt.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If rt.EOF Then
      rt.AddNew
      rt!tarjeta = UCase(Txtarjeta)
      If IsDate(Txfecha) Then rt!fecha = Txfecha
      rt!hora = Txhora
   End If
   rt!tipo = TxTipo
   rt!motivo = TxMotivo
   rt!observacion = Trim(Txobservacion)
   rt!usuario = wNick       'wUsuario      Firma Usuario
   rt!fecRegistro = Date
   rt!HorRegistro = Format(Time, "hh:mm:ss")
   rt!estacion = w_clienName
   rt.Update
   rt.Close
   If Sw Then
      DD.AddNew
      DD!tarjeta = Txtarjeta
   End If
   If IsDate(Txfecha) Then DD!fecha = Txfecha
   DD!hora = Txhora
   DD!tipo = TxTipo
   DD!motivo = TxMotivo
   DD!observacion = Txobservacion
   DD!usuario = wUsuario
   DD!fecRegistro = Date
   DD.Update
   Cnx.CommitTrans
   GoTo Final1
final:
   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
   Cnx.RollbackTrans
Final1:
   Call Limpiacampos
   'Call Desabilitacampos
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
    
    If Not IsDate(Txfecha) Then
         MsgBox "Verifique fecha...  ", vbExclamation
         Exit Sub
    End If
    If Year(Txfecha) < Val(genaño) Then
       MsgBox "Verifique fecha...  ", vbExclamation
       Exit Sub
    End If
    
    fecha1 = Format(Txfecha, "yyyymmdd")
    Set rs = New Recordset
    sq2 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta ='" & Trim(Txtarjeta) & "' and fecha='" & fecha1 & "' and hora='" & Txhora.Text & "' and tipo='" & TxTipo.Text & "'"
    rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then
       MsgBox "No Tiene Marcas...  ", vbExclamation
       rs.Close
       Exit Sub
    Else
       rs.Close
    End If
    If MsgBox("Eliminar  registro ?      ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       Cnx.Execute "DELETE FROM CpMarca" & genaño & " WHERE tarjeta ='" & Trim(Txtarjeta.Text) & "' and fecha='" & fecha1 & "' and hora='" & Txhora & "' and tipo='" & TxTipo & "'"
       DD.Delete
    End If
    Call Limpiacampos
    'Call Desabilitacampos
End Sub
Private Sub CmdLimpia_Click()
   Call Limpiacampos
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
   If IsDate(TxDia) Then
      CargaMdia (0)
      nTipRpt = 4
      InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
      InfoCpSQL.Refresh
      InfoCpSQL.Show
   Else
      MsgBox "No indica DIA..", vbExclamation
   End If
   Exit Sub
End Sub

Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Txtarjeta) = "" Or (Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta)) Then
         If Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta) Then
            Ayuda1.TxUbiAP = Trim(Txtarjeta)
         End If
         Load Ayuda1
         Ayuda1.Tag = 2
         Ayuda1.Show 1
      Else
         Set rt = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
         rt.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rt.EOF Then
            MsgBox "No Ubicado", vbExclamation
            Txtarjeta = ""
         Else
            TxTrabajador = Left(Trim(rt!apelpat) + " " + Trim(rt!apelmat) + " " + Trim(rt!nombres), 35)
            TxGrpGen = IIf(IsNull(rt!grpgen), "", rt!grpgen)
         End If
         rt.Close
      End If
      Txfecha.SetFocus
      Exit Sub
   End If
End Sub
Private Sub Txfecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Not IsDate(Txfecha) Then
         MsgBox "Verifique fecha...  ", vbExclamation
         Exit Sub
      End If
      If Year(Txfecha) < Val(genaño) Then
         MsgBox "Verifique fecha...  ", vbExclamation
         Exit Sub
      End If
      p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": cdiaant = ""
      fechaB = Format(Txfecha, "yyyymmdd")
      sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
      Set rs = New Recordset
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         rs.Close
         fechaB = Format(DateValue(Txfecha) - 1, "yyyymmdd")
         sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
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
               LblHorario.ForeColor = &HC00000     'azul
               If cdiaant = "S" Then
                  LblHorario.ForeColor = &HFF&  'rojo
               End If
               LblHorario = p_ing + "-" + p_sal + " / " + p_ing1 + "-" + p_sal1
            End If
         End If
      End If
   End If
End Sub
Private Sub Txfecha_LostFocus()

      If Not IsDate(Txfecha) Then
         MsgBox "Verifique fecha...  ", vbExclamation
         Exit Sub
      End If
      If Year(Txfecha) < Val(genaño) Then
         MsgBox "Verifique fecha...  ", vbExclamation
         Exit Sub
      End If
      
      p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": cdiaant = ""
      fechaB = Format(Txfecha, "yyyymmdd")
      sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
      Set rs = New Recordset
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         rs.Close
         fechaB = Format(DateValue(Txfecha) - 1, "yyyymmdd")
         sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & Txtarjeta & "' AND fecha ='" & fechaB & "'"
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
               LblHorario.ForeColor = &HC00000     'azul
               If cdiaant = "S" Then
                  LblHorario.ForeColor = &HFF&  'rojo
               End If
               LblHorario = p_ing + "-" + p_sal + " / " + p_ing1 + "-" + p_sal1
            End If
            Txhora.SetFocus
         End If
      End If
   
End Sub

Private Sub Cbotipo_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Cbotipo.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbotipo_LostFocus()
   If IsEmpty(Cbotipo.Text) Then
   Else
      j = Cbotipo.ListIndex
      If j <> -1 Then
         nposicion = InStr(Cbotipo.List(j), ":")
         nlen = Len(Cbotipo.List(j))
         TxTipo.Text = Trim(Right(Cbotipo.List(j), nlen - nposicion))
         'TxDscEstab.Text = Left(CboEstab.List(j), nposicion - 1)
      End If
      'Cbomotivo.SetFocus
      Exit Sub
   End If
End Sub
Private Sub Cbomotivo_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Cbomotivo.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbomotivo_LostFocus()
   If IsEmpty(Cbomotivo.Text) Then
   Else
      j = Cbomotivo.ListIndex
      If j <> -1 Then
         nposicion = InStr(Cbomotivo.List(j), ":")
         nlen = Len(Cbomotivo.List(j))
         TxMotivo.Text = Trim(Right(Cbomotivo.List(j), nlen - nposicion))
         'TxDscEstab.Text = Left(CboEstab.List(j), nposicion - 1)
      End If
      'cmdGraba.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Cbocargo_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboCargo.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbocargo_LostFocus()
  If IsEmpty(CboCargo.Text) Then
  Else
     j = CboCargo.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboCargo.List(j), ":")
        nlen = Len(CboCargo.List(j))
        Txcargo.Text = Right(CboCargo.List(j), nlen - nposicion)
     End If
     Exit Sub
  End If
End Sub

Private Sub CargaMdia(nopc)
'***  CARGA  registros  ***   0:Todos   1:Orden
   '' Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmRecord"
    nposicion = 0
    If nopc = 0 Then
       On Error Resume Next
       unafecha = TxDia.Text
       pridia = unafecha
       ultdia = unafecha + 1
       fecha = Format(unafecha, "yyyymmdd")
       fecha1 = Format(unafecha + 1, "yyyymmdd")
       Set rs = New Recordset
       SQ1 = "SELECT * FROM CpMarca" & genaño & " WHERE fecha >='" & fecha & "' and  fecha <='" & fecha1 & " ' ORDER BY fecha,tarjeta,hora"
       rs.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       If rs.EOF Then
       Else
          Do While Not rs.EOF
             nposicion = nposicion + 1
             cTarjeta = rs!tarjeta
             bfecha = Format(rs!fecha, "yyyymmdd")
             Cmdprint.Caption = bfecha + ": " + cTarjeta
             cHora = rs!hora
             ndato = 0
             Set rt = New Recordset
             sq2 = "SELECT * FROM CpTmRecord WHERE fecha='" & bfecha & "' and tarjeta='" & cTarjeta & "'"
             rt.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
             If rt.EOF Then
                rt.AddNew
                rt!fecha = rs!fecha
                rt!tarjeta = cTarjeta
             End If
             If IsNull(rt!ming) Or Trim(rt!ming) = "" Or (Val(rs!hora) - Val(rt!ming) < 5) Then
                rt!ming = rs!hora
             ElseIf IsNull(rt!msal) Or Trim(rt!msal) = "" Or (Val(rs!hora) - Val(rt!msal) < 5) Then
                    rt!msal = rs!hora
             ElseIf IsNull(rt!ming1) Or Trim(rt!ming1) = "" Or (Val(rs!hora) - Val(rt!ming1) < 5) Then
                    rt!ming1 = rs!hora
             ElseIf IsNull(rt!msal1) Or Trim(rt!msal1) = "" Or (Val(rs!hora) - Val(rt!msal1) < 5) Then
                    rt!msal1 = rs!hora
             End If
             rt.Update
             rt.Close
             rs.MoveNext
          Loop
       End If
       rs.Close
       Cmdprint.Caption = ""
       o_tarjeta = "TARJE"
       Set rt = New Recordset
       SQ1 = "SELECT * FROM CpTmRecord  ORDER BY fecha,tarjeta "
       rt.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       Do While Not rt.EOF
          On Error Resume Next
          cTarjeta = rt!tarjeta
          bfecha = Format(rt!fecha, "yyyymmdd")
          Cmdprint.Caption = bfecha + ": " + cTarjeta
          If o_tarjeta <> rt!tarjeta Then
             cGrpGen = "": cvalor = ""
             Set rs = New Recordset
             sql = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta='" & cTarjeta & "'"
             rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs.EOF Then
                cGrpGen = IIf(IsNull(rs!grpgen), "", rs!grpgen)
             End If
             rs.Close
             o_tarjeta = rt!tarjeta
          End If
          'Obtiene HORARIO prog
          Set rs_d = New Recordset
          sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & bfecha & "'"
          rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If rs_d.EOF Then
          Else
             cvalor = IIf(IsNull(rs_d!Horario), "", rs_d!Horario)
          End If
          rs_d.Close
          If cvalor = "" And Trim(cGrpGen) <> "" Then
             Set rs_d = New Recordset
             sql = "SELECT * FROM CpHor" & genaño & " WHERE grupo='" & cGrpGen & "' AND fecha ='" & bfecha & "'"
             rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If rs_d.EOF Then
             Else
                cvalor = IIf(IsNull(rs_d!Horario), "", rs_d!Horario)
             End If
             rs_d.Close
          End If
          p_ping = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": ntardanza = 0
          m_ing = "": m_sal = "": m_ing1 = "": m_sal1 = ""
          If Trim(cvalor) <> "" Then
             Set rs_d = New Recordset
             sql = "SELECT * FROM CpHorarioNew  WHERE Codigo='" & cvalor & "'"
             rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If rs_d.EOF Then
                rs_d.Close
                Set rs_d = New Recordset
                sql = "SELECT * FROM CpHorarios  WHERE Codigo='" & cvalor & "'"
                rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
             End If
             If Not rs_d.EOF Then
                p_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
                p_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                p_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
                p_sal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
                p_horas = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
                nValorH = Val(Left(p_horas, 2))
                nValorM = Val(Right(p_horas, 2))
                'Evalua MARCAS del dia       fecha  hora =>  hhmm
                If Val(p_horas) > 0 Then
                   m_ing = rt!ming
                   m_sal = rt!msal
                   m_ing1 = rt!ming1
                   m_sal1 = rt!msal1
                   'Corrige doble marcacion
                   If Val(p_ing1) + Val(p_sal1) = 0 And Val(rt!msal) > 0 And Val(rt!ming1) > 0 Then
                      m_sal = IIf(rt!ming1 > rt!msal, rt!ming1, rt!msal)
                   End If
                   If Val(p_ing1) + Val(p_sal1) = 0 And Val(rt!msal) > 0 And Val(rt!ming1) > 0 Then
                      m_ing1 = IIf(rt!msal < rt!ming1, rt!sal, rt!ming1)
                   End If
                   If p_ing < rt!ming Then
                      ntardanza = Val(Left(m_ing, 2)) - Val(Left(p_ing, 2)) + (Val(Right(m_ing, 2)) - Val(Right(p_ing, 2))) / 60
                   ElseIf p_ing1 < m_ing1 And Val(p_sal1) > 0 Then
                          ntardanza = Val(Left(m_ing1, 2)) - Val(Left(p_ing1, 2)) + (Val(Right(m_ing1, 2)) - Val(Right(p_ing1, 2))) / 60
                   End If
                   nValorH = Int(ntardanza)                      ' 1.166   => 1
                   nValorM = ntardanza - nValorH                 '         =>  .166
                   ntardanza = nValorH + ((nvalor) * 0.6)        ' 1.10   0=> 1:10
                End If
             End If
             rs_d.Close
             Set rs = New Recordset
             sql = "SELECT * FROM CpTmRecord WHERE fecha='" & bfecha & "' and tarjeta='" & cTarjeta & "'"
             rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
             If rs.EOF Then
             Else
                rs!ping = p_ing
                rs!psal = p_sal
                rs!ping1 = p_ing1
                rs!psal1 = p_sal1
                rs!ming = m_ing
                rs!msal = m_sal
                rs!ming1 = m_ing1
                rs!msal1 = m_sal1
                rs!tardanza = ntardanza
                rs!Horario = cvalor
                rs!hrasprog = Val(Left(p_horas, 2)) + Val(Right(p_horas, 2)) / 100
                rs.Update
             End If
             rs.Close
          End If
          rt.MoveNext
       Loop
       Cmdprint.Caption = ""
    End If
    ''Cnx.CommitTrans
    rt.Close
End Sub

