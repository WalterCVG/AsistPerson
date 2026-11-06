VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AutorizaQ 
   Caption         =   "Consulta Autorizaciones"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "AutorizaQ.frx":0000
   ScaleHeight     =   5655
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxDH 
      Enabled         =   0   'False
      Height          =   195
      Left            =   6840
      MaxLength       =   1
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox Cbo_DH 
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Top             =   240
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
      Picture         =   "AutorizaQ.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir"
      Top             =   4080
      Width           =   645
   End
   Begin VB.CommandButton mnuFicImpSelec 
      Height          =   495
      Left            =   8520
      Picture         =   "AutorizaQ.frx":80B30
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Selec.Impresora"
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton CmdLIMPIA 
      Caption         =   "Limpia"
      Height          =   615
      Left            =   8520
      Picture         =   "AutorizaQ.frx":8130E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Limpia Plantilla"
      Top             =   2040
      Width           =   615
   End
   Begin VB.TextBox Txtrabajador 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      MaxLength       =   35
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin VB.TextBox Txtarjeta 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   8640
      Picture         =   "AutorizaQ.frx":81840
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cerrar"
      Top             =   4920
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   8055
      _ExtentX        =   14208
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
   Begin VB.CommandButton cmdBusca 
      Height          =   615
      Left            =   8520
      Picture         =   "AutorizaQ.frx":81DC2
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ubicar Registros"
      Top             =   600
      Width           =   735
   End
   Begin VB.ComboBox CboOcurre 
      Height          =   315
      Left            =   4320
      TabIndex        =   8
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox Txcodigo 
      Enabled         =   0   'False
      Height          =   195
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxfechaF 
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
      Left            =   2400
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox TxfechaI 
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
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Por :"
      Height          =   255
      Left            =   6840
      TabIndex        =   18
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivo"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "hasta"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "AutorizaQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DD As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim r_d As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim Criterio, ccadena As String
Dim nson, i, nai, naf As Integer
Private Sub Form_Load()
  If Autorizaciones.Tx_tarjeta <> "" Then
     Txtarjeta = Autorizaciones.Tx_tarjeta
     Txtrabajador = Autorizaciones.Tx_trabajador
  End If
  Cbo_DH.Clear
  Cbo_DH.AddItem "Dia :D"
  Cbo_DH.AddItem "Hora:H"
  Cbo_DH.Text = ""
  TxDH.Text = ""
  TxfechaI = DateValue("01/01/" + Trim(Str(genaño)))
  TxfechaF = Date
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpOcurrencias ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboOcurre.ListCount - 1
           If CboOcurre.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboOcurre.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  Set DD = New Recordset
  With DD
     .Fields.Append "tarjeta", adChar, 5
     .Fields.Append "trabajador", adChar, 35
     .Fields.Append "Codigo", adChar, 4
     .Fields.Append "DiaHora", adChar, 1
     .Fields.Append "Fechaini", adDate
     .Fields.Append "FechaFin", adDate
     .Fields.Append "horaIni", adChar, 4
     .Fields.Append "horaFin", adChar, 4
     .Fields.Append "Tipo", adChar, 1
     .Fields.Append "Documento", adChar, 30
     .Fields.Append "Turno", adChar, 2
     .Fields.Append "Usuario", adChar, 10     'registrado
     .Fields.Append "Fecha", adDate
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
          Case 0
          '.Caption = "tarjeta"
          .Width = Me.TextWidth(String(5, "S"))
          Case 1
          '.Caption = "trabajador"
          .Width = Me.TextWidth(String(35, "S"))
          Case 2
          '.Caption = "codigo"
          .Width = Me.TextWidth(String(7, "S"))
          Case 3
          '.Caption ="DiaHora"
          .Width = Me.TextWidth(String(6, "S"))
          Case 4
          '.Caption ="fechaini"
          .Width = Me.TextWidth(String(10, "S"))
          Case 5
          '.Caption ="fechafin"
          .Width = Me.TextWidth(String(10, "S"))
          Case 6
          '.Caption ="HoraIni"
          .Width = Me.TextWidth(String(7, "S"))
          Case 7
          '.Caption ="HoraFin"
          .Width = Me.TextWidth(String(7, "S"))
          Case 8
          '.Caption ="Tipo"
          .Width = Me.TextWidth(String(6, "S"))
          Case 9
          '.Caption ="Documento"
           .Width = Me.TextWidth(String(30, "S"))
            Case 11
          .Caption = "Registrado"
           .Width = Me.TextWidth(String(10, "S"))
           Case 12
          '.Caption ="Fecha"
           .Width = Me.TextWidth(String(10, "S"))
          End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
  Next
  DataGrid1.Enabled = False
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   With DD      ' Adodc1.Recordset
      If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
         .Sort = .Fields(ColIndex).[Name] & " Desc"
      Else
         .Sort = .Fields(ColIndex).[Name] & " Asc"
      End If
   End With
End Sub

Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Txtarjeta.Text) = "" Or (Trim(Txtarjeta.Text) <> "" And EsTexto(Txtarjeta.Text)) Then
         If Trim(Txtarjeta.Text) <> "" And EsTexto(Txtarjeta.Text) Then
            Ayuda1.TxUbiAP = Trim(Txtarjeta.Text)
         End If
         Load Ayuda1
         Ayuda1.Tag = 4.2
         Ayuda1.Show 1
      Else
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & Txtarjeta.Text & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "Tarjeta, No Ubicada", vbExclamation
            Txtarjeta.Text = ""
            Txtrabajador.Text = ""
         Else
            Txtrabajador.Text = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
         End If
         rs.Close
      End If
      Exit Sub
   End If
End Sub

Private Sub Cbo_DH_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(Cbo_DH.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbo_DH_LostFocus()
   If IsEmpty(Cbo_DH.Text) Then
   Else
      j = Cbo_DH.ListIndex
      If j <> -1 Then
         nposicion = InStr(Cbo_DH.List(j), ":")
         nlen = Len(Cbo_DH.List(j))
         TxDH.Text = Trim(Right(Cbo_DH.List(j), nlen - nposicion))
      End If
      Exit Sub
   End If
End Sub
Private Sub CboOcurre_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboOcurre.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboOcurre_LostFocus()
   If IsEmpty(CboOcurre.Text) Then
   Else
      j = CboOcurre.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboOcurre.List(j), ":")
         nlen = Len(CboOcurre.List(j))
         Txcodigo = Trim(Right(CboOcurre.List(j), nlen - nposicion))
      End If
      Exit Sub
   End If
End Sub

Private Sub cmdbusca_Click()
   If Not IsDate(TxfechaI) And Not IsDate(TxfechaF) Then Exit Sub
   If Txtarjeta <> "" Then
      Set rs = New Recordset
      sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
      rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "Servidor..No Ubicado", vbExclamation
         Txtarjeta = ""
         Txtrabajador = ""
         Exit Sub
      Else
         Txtrabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
      End If
      rs.Close
   End If
      
   fecha1 = Format(TxfechaI, "yyyymmdd")
   fecha2 = Format(TxfechaF, "yyyymmdd")
   nai = Year(TxfechaI)
   naf = Year(TxfechaF)
   
   If Trim(Txcodigo) = "" Then
      If Trim(TxDH) = "" Then
         Criterio = "Fechaini <= '" & fecha2 & "' AND FechaFin >= '" & fecha1 & "'"
         If Txtarjeta <> "" Then
            Criterio = "tarjeta='" & Txtarjeta & "' AND Fechaini <= '" & fecha2 & "' AND FechaFin >= '" & fecha1 & "'"
         End If
      Else
         Criterio = "DiaHora='" & TxDH & "' AND  FechaIni <= '" & fecha2 & "' AND FechaFin >= '" & fecha1 & "'"
         If Txtarjeta <> "" Then
            Criterio = "tarjeta='" & Txtarjeta & "' AND DiaHora='" & TxDH & "' AND  FechaIni <= '" & fecha2 & "' AND FechaFin >= '" & fecha1 & "'"
         End If
      End If
   Else
      If Trim(TxDH) = "" Then
         Criterio = "codigo='" & Txcodigo & "' AND FechaIni <= '" & fecha2 & "' AND FechaFin >= '" & fecha1 & "'"
         If Txtarjeta <> "" Then
            Criterio = "tarjeta='" & Txtarjeta & "' AND codigo='" & Txcodigo & "' AND FechaIni <= '" & fecha2 & "' AND FechaFin >= '" & fecha1 & "'"
         End If
      Else
         Criterio = "DiaHora='" & TxDH & "' AND  codigo='" & Txcodigo & "' AND FechaIni <= '" & fecha2 & "' AND FechaFin >= '" & fecha1 & "'"
         If Txtarjeta <> "" Then
            Criterio = "tarjeta='" & Txtarjeta & "' AND DiaHora='" & TxDH & "' AND codigo='" & Txcodigo & "'  AND FechaIni <= '" & fecha2 & "' AND FechaFin >= '" & fecha1 & "'"
         End If
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
     .Fields.Append "trabajador", adChar, 35
     .Fields.Append "Codigo", adChar, 4
     .Fields.Append "DiaHora", adChar, 1
     .Fields.Append "Fechaini", adDate
     .Fields.Append "FechaFin", adDate
     .Fields.Append "horaIni", adChar, 4
     .Fields.Append "horaFin", adChar, 4
     .Fields.Append "Tipo", adChar, 1
     .Fields.Append "Documento", adChar, 30
     .Fields.Append "Turno", adChar, 2
     .Fields.Append "Usuario", adChar, 10     'registrado
     .Fields.Append "Fecha", adDate
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
          Case 0
          '.Caption = "tarjeta"
          .Width = Me.TextWidth(String(5, "S"))
          Case 1
          '.Caption = "trabajador"
          .Width = Me.TextWidth(String(35, "S"))
          Case 2
          '.Caption = "codigo"
          .Width = Me.TextWidth(String(7, "S"))
          Case 3
          '.Caption ="DiaHora"
          .Width = Me.TextWidth(String(6, "S"))
          Case 4
          '.Caption ="fechaini"
          .Width = Me.TextWidth(String(10, "S"))
          Case 5
          '.Caption ="fechafin"
          .Width = Me.TextWidth(String(10, "S"))
          Case 6
          '.Caption ="HoraIni"
          .Width = Me.TextWidth(String(7, "S"))
          Case 7
          '.Caption ="HoraFin"
          .Width = Me.TextWidth(String(7, "S"))
          Case 8
          '.Caption ="Tipo"
          .Width = Me.TextWidth(String(6, "S"))
          Case 9
          '.Caption ="Documento"
           .Width = Me.TextWidth(String(30, "S"))
            Case 11
          .Caption = "Registrado"
           .Width = Me.TextWidth(String(10, "S"))
           Case 12
          '.Caption ="Fecha"
           .Width = Me.TextWidth(String(10, "S"))
          End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
  Next
   
   ccadena = ""
   For i = nai To naf
      anohoy = Trim(Str(i))
      sql = "SELECT * FROM CpAutori" & anohoy & " WHERE (anulado is NULL) and " & Criterio & " ORDER BY   tarjeta,fechaini"
      Set rs = New Recordset
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         Do While Not rs.EOF
            cTarjeta = rs!tarjeta
            cCodigo = rs!Codigo
            cDiahora = rs!diahora
            Fechai = Format(rs!fechaIni, "yyyymmdd")
            If InStr(ccadena, cTarjeta + cCodigo + "*" + cDiahora + Fechai) = 0 Then
               ccadena = ccadena + cTarjeta + cCodigo + "*" + cDiahora + Fechai + "_"
               ctrabajador = ""
               Set r_d = New Recordset
               sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & rs!tarjeta & "'"
               r_d.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If r_d.EOF Then
               Else
                  ctrabajador = Left(Trim(r_d!apelpat) + " " + Trim(r_d!apelmat) + " " + Trim(r_d!nombres), 35)
               End If
               r_d.Close
               
               DD.AddNew
               DD!Codigo = IIf(IsNull(rs!Codigo), "", rs!Codigo)
               DD!tarjeta = IIf(IsNull(rs!tarjeta), "", rs!tarjeta)
               DD!trabajador = ctrabajador
               DD!diahora = IIf(IsNull(rs!diahora), "", rs!diahora)
               If IsDate(rs!fechaIni) Then DD!fechaIni = rs!fechaIni
               If IsDate(rs!fechaFin) Then DD!fechaFin = rs!fechaFin
               DD!horaini = IIf(IsNull(rs!horaini), "", rs!horaini)
               DD!horafin = IIf(IsNull(rs!horafin), "", rs!horafin)
               DD!tipo = IIf(IsNull(rs!tipo), "", rs!tipo)
               DD!Documento = IIf(IsNull(rs!Documento), "", rs!Documento)
               DD!turno = IIf(IsNull(rs!turno), "", rs!turno)
               DD!usuario = IIf(IsNull(rs!usuario), "", rs!usuario)
               If IsDate(rs!fecha) Then DD!fecha = rs!fecha
               DD.Update
               nson = nson + 1
            End If
            rs.MoveNext
         Loop
      rs.Close
   Next i
   DataGrid1.Refresh
   DataGrid1.Visible = True
   
   'Call cmdLimpia_click()
   If nson = 0 Then
      MsgBox "No registra autorizaciones..", vbExclamation
   End If
End Sub
Private Sub CmdLimpia_Click()
   Txtarjeta = "": Txtrabajador = ""
   CboOcurre.Text = "": Txcodigo = ""
   Cbo_DH.Text = "": TxDH = ""
   TxfechaI = "": TxfechaF = ""
   Do While Not DD.EOF
      DD.Delete
      DD.MoveFirst
   Loop
   DataGrid1.Refresh
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
   If IsEmpty(Cbo_diaHra.Text) Then
   Else
      j = Cbo_diaHra.ListIndex
      If j <> -1 Then
         nposicion = InStr(Cbo_diaHra.List(j), ":")
         nlen = Len(Cbo_diaHra.List(j))
         Tx_diaHra.Text = Trim(Right(Cbo_diaHra.List(j), nlen - nposicion))
      End If
      Exit Sub
   End If
End Sub
Private Sub CmdPrint_Click()
   nson = 0
   Call CargaAutoriza
   If nson = 0 Then
      MsgBox "Seleccion, no registra autorizaciones..", vbExclamation
      Exit Sub
   End If
   W_lbldh = "Dia/Mint"
   If TxDH = "D" Then
      W_lbldh = "Dias"
   ElseIf TxDH = "H" Then
          W_lbldh = "Mints."
   End If
   If Txtarjeta = "" And Txcodigo <> "" Then
      nTipRpt = 3.02
   Else
      nTipRpt = 3
   End If
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Exit Sub
End Sub

Private Sub CargaAutoriza()
'***  CARGA  registros  ***   0:Mes   1:dia
    'Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmpTabla"
    Set rs_d = New Recordset
    sq_l = "SELECT * FROM CpTmptabla"
    rs_d.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    nson = 0
    fecha1 = Format(TxfechaI, "yyyymmdd")
    fecha2 = Format(TxfechaF, "yyyymmdd")
    nai = Year(TxfechaI)
    naf = Year(TxfechaF)
    ccadena = ""
    For i = nai To naf
      anohoy = Trim(Str(i))
      sql = "SELECT * FROM CpAutori" & anohoy & " WHERE (anulado is NULL) and " & Criterio & " ORDER BY FechaIni "
      Set r_d = New ADODB.Recordset
      r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      Do While Not r_d.EOF
         cTarjeta = r_d!tarjeta
         cCodigo = r_d!Codigo
         cDiahora = r_d!diahora
         Fechai = Format(r_d!fechaIni, "yyyymmdd")
         If InStr(ccadena, cTarjeta + cCodigo + "*" + cDiahora + Fechai) = 0 Then
            ccadena = ccadena + cTarjeta + cCodigo + "*" + cDiahora + Fechai + "_"
            nson = nson + 1
            rs_d.AddNew
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
            rs_d!Documento = IIf(IsEmpty(r_d!Documento), "", r_d!Documento)
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
            rs_d!usuario = r_d!usuario
            rs_d!fecha = r_d!fecha
            rs_d.Update
         End If
         r_d.MoveNext
      Loop
      ' Cnx.CommitTrans
      r_d.Close
    Next i
    rs_d.Close
End Sub

Private Sub cmdsalir_Click()
   Set DD = Nothing
   Set rs = Nothing
   Set r_d = Nothing
   Set rs_d = Nothing
   Me.Hide
   Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set DD = Nothing
   Set rs = Nothing
   Set r_d = Nothing
   Set rs_d = Nothing
   Me.Hide
   Unload Me
End Sub



