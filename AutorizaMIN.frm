VERSION 5.00
Begin VB.Form AutorizaMIN 
   Caption         =   "Generar Data de AUTORIZACIONES"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "AutorizaMIN.frx":0000
   ScaleHeight     =   3495
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   8760
      Picture         =   "AutorizaMIN.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Cerrar"
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton CmdExportaXLS 
      Height          =   615
      Left            =   7680
      Picture         =   "AutorizaMIN.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exporta a Xls"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox ChkXLS 
      Alignment       =   1  'Right Justify
      Caption         =   "XLS instalado"
      Height          =   195
      Left            =   6360
      TabIndex        =   20
      Top             =   2880
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "AutorizaMIN.frx":80FF2
      Left            =   6240
      List            =   "AutorizaMIN.frx":80FF4
      TabIndex        =   18
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   495
      Left            =   5160
      TabIndex        =   16
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Txtarjeta 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox TxDH 
      Enabled         =   0   'False
      Height          =   195
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox Cbo_DH 
      Height          =   315
      Left            =   6720
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton CmdLIMPIA 
      Caption         =   "Limpia"
      Height          =   615
      Left            =   8400
      Picture         =   "AutorizaMIN.frx":80FF6
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Limpia Plantilla"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Txtrabajador 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      MaxLength       =   35
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.CommandButton cmdBusca 
      Height          =   1095
      Left            =   8040
      Picture         =   "AutorizaMIN.frx":81528
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ubicar Registros"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox CboOcurre 
      Height          =   315
      Left            =   960
      TabIndex        =   8
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox Txcodigo 
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1800
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
      Left            =   2520
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
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   6360
      TabIndex        =   19
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivos elegidos (max. 3)"
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Por :"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Motivo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "hasta"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "AutorizaMIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DD As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim r_d As New ADODB.Recordset
Dim Criterio, ccadena, cUbica, cAutorizo, Sq0, canoHoy, canofue, cDOI, cCargo, cDscCargo, Ctipo As String
Dim nson, i, j, nai, naf, m_hini, m_hfin, m_tot As Integer

Private Sub Form_Load()
  Cbo_DH.Clear
  Cbo_DH.AddItem "Dia :D"
  Cbo_DH.AddItem "Hora:H"
  Cbo_DH.Text = ""
  TxDH = ""
  Label5.Caption = ""
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
  
End Sub


Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Txtarjeta) = "" Or (Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta)) Then
         If Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta) Then
            Ayuda1.TxUbiAP = Trim(Txtarjeta)
         End If
         Load Ayuda1
         Ayuda1.Tag = 4.3
         Ayuda1.Show 1
      Else
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "Tarjeta, No Ubicada", vbExclamation
            Txtarjeta = ""
            Txtrabajador = ""
         Else
            Txtrabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
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
         TxDH = Trim(Right(Cbo_DH.List(j), nlen - nposicion))
      End If
      Exit Sub
   End If
End Sub
Private Sub CboOcurre_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboOcurre.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboOcurre_Click()
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

Private Sub CboOcurre_LostFocus()
  Call CboOcurre_Click
End Sub

Private Sub Command1_Click()
   If List1.ListCount = 3 Then Exit Sub     ' Max. 3 codigos
   j = CboOcurre.ListIndex
   '-Caso Edicion -- Verifica NO duplicados
   nposicion = InStr(CboOcurre.List(j), ":")
   nlen = Len(CboOcurre.List(j))
   For X = 0 To List1.ListCount - 1
      If List1.List(X) = Txcodigo Then Exit Sub
   Next X
   List1.AddItem Txcodigo
   Label5.Caption = "Hay " & List1.ListCount & " Item(s)"
   List1.ListIndex = List1.ListCount - 1
   '-Libera seleccion
   'For X = 0 To cboOcurre.ListCount - 1
   '  If List1.Selected(X) = True Then
   '     CboOcurre.Selected(j) = False
   '  End If
   'Next X
   Txcodigo = ""
   CboOcurre.Text = ""
   CboOcurre.SetFocus
End Sub
Private Sub Command2_Click()
   List1.RemoveItem List1.ListIndex
   Label5.Caption = "Hay " & List2.ListCount & " Item(s)"
End Sub

Private Sub cmdbusca_Click()
   If Not IsDate(TxfechaI) And Not IsDate(TxfechaF) Then Exit Sub
   If Trim(TxDH) = "" Then Exit Sub
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
   Criterio = "(Fechaini <= '" & fecha2 & "' AND FechaFin >= '" & fecha1 & "') AND (anulado is NULL) "
   
   If Trim(TxDH) = "" Then
      If Txtarjeta <> "" Then
         Criterio = Criterio + " AND (tarjeta='" & Txtarjeta & "') "
      End If
   Else
      If Txtarjeta <> "" Then
         Criterio = Criterio + " AND (tarjeta='" & Txtarjeta & "') AND (DiaHora='" & TxDH & "') "
      End If
   End If
   
   '-Carga Autorizaciones
   ccadena = ""
   nson = List1.ListCount
   If nson > 0 Then
      ccadena = " AND ( Codigo IN('"         '-->IN('aa','bb','cc')
      For X = 0 To List1.ListCount - 1
          If Trim(List1.List(X)) <> "" Then
             If X < nson - 1 Then
                ccadena = ccadena + Left(Trim(List1.List(X)) + Space(4), 4) + "','"
             Else
                '-ultimo item
                ccadena = ccadena + Left(Trim(List1.List(X)) + Space(4), 4) + "') ) "
             End If
          End If
      Next X
      For X = 0 To List1.ListCount - 1
         If List1.Selected(X) = True Then          'Libera seleccion
            List1.Selected(X) = False
         End If
      Next X
      Criterio = Criterio + ccadena
   End If
   
   '--MsgBox Criterio
   
   '''On Error Resume Next
   Cnx.BeginTrans
   On Error GoTo CREA
   Cnx.Execute ("DELETE  FROM #tablaMI ")
   GoTo sigue
CREA:
   Cnx.Execute ("SELECT * Into #tablaMI FROM CpAutori" + genaño + "  WHERE 1 = 2")
   sql = "ALTER TABLE #tablaMI DROP COLUMN  FECHA, USUARIO, Hora, Estacion, anulado, UsuAnula, fecAnula, HorAnula , EstAnula"
   Cnx.Execute (sql)
   sql = "ALTER TABLE #tablaMi ADD DOI varchar(10) null,trabajador varchar(35) null, TipoPla varchar(3) null, cargo varchar(50) null, valor integer null "
   Cnx.Execute (sql)
   
sigue:
   Cnx.CommitTrans
       
   Sq0 = "SELECT * FROM #tablaMI "
   Set rt = New Recordset
   rt.Open Sq0, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   
   nson = 0
   cAutorizo = ""
   canoayer = ""
   
   Dim OpenForms
   j = 0
   
   For i = nai To naf
      anohoy = Trim(Str(i))
      sql = "SELECT * FROM CpAutori" & anohoy & " WHERE  " & Criterio & " ORDER BY  tarjeta,fechaini"
      Set rs = New Recordset
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      Do While Not rs.EOF
      
           '-Dar control a Windows y retomar hilo
           j = j + 1
           If j Mod 20 = 0 Then        'if loop has repeated
              OpenForms = DoEvents     'Yield to operating system
           End If
      
            cTarjeta = rs!tarjeta
            cCodigo = rs!Codigo
            cDiahora = rs!diahora
            Fechai = Format(rs!fechaIni, "yyyymmdd")
            cUbica = cTarjeta + cCodigo + ":" + cDiahora + Fechai
            
            If canoHoy <> canoayer Then
               MsgBox ("paso " + anohoy + " " + cUbica)
               canoayer = canoHoy
            End If
            
            If InStr(cAutorizo, cUbica) = 0 Then
               cAutorizo = cAutorizo + (cUbica + "_")
               
               '-MsgBox ("paso1 " + cUbica)
               
               cmdBusca.Caption = cTarjeta + ":" + Fechai
               nson = nson + 1
               ctrabajador = ""
               cDOI = "": cCargo = "": cDscCargo = "": ctipopla = ""
               Set r_d = New Recordset
               sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & rs!tarjeta & "'"
               r_d.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not r_d.EOF Then
                  cDOI = IIf(IsNull(r_d!documento), "", Trim(r_d!documento))
                  ctrabajador = Left(Trim(r_d!apelpat) + " " + Trim(r_d!apelmat) + " " + Trim(r_d!nombres), 35)
                  cCargo = r_d!Cargo
                  ctipopla = r_d!tipoPla
                  If Trim(cCargo) <> "" Then
                     Set DD = New Recordset
                     sql = "SELECT * FROM CpCargos WHERE codigo='" & Trim(cCargo) & "'"
                     DD.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                     If Not DD.EOF Then
                        cDscCargo = Trim(DD!descripcion)
                     End If
                     DD.Close
                  End If
               End If
               r_d.Close
               
               rt.AddNew
               rt!Codigo = IIf(IsNull(rs!Codigo), "", rs!Codigo)
               rt!tarjeta = IIf(IsNull(rs!tarjeta), "", rs!tarjeta)
               rt!DOI = cDOI
               rt!trabajador = ctrabajador
               rt!tipoPla = ctipopla
               rt!Cargo = cDscCargo
               rt!diahora = IIf(IsNull(rs!diahora), "", rs!diahora)
               If IsDate(rs!fechaIni) Then rt!fechaIni = rs!fechaIni
               If IsDate(rs!fechaFin) Then rt!fechaFin = rs!fechaFin
               If rs!diahora = "D" Then
                  rt!Valor = rs!fechaFin - rs!fechaIni + 1
                  rt!turno = IIf(IsEmpty(rs!turno), "", rs!turno)
               End If
               m_hini = "": m_hfin = "": m_tot = 0
               If rs!diahora = "H" Then
                  rt!horaini = IIf(IsEmpty(rs!horaini), "", rs!horaini)
                  rt!horafin = IIf(IsEmpty(rs!horafin), "", rs!horafin)
                  rt!tipo = IIf(IsNull(rs!tipo), "", rs!tipo)
                  m_hini = IIf(IsNull(rs!horaini), "", rs!horaini)
                  m_hfin = IIf(IsNull(rs!horafin), "", rs!horafin)
                  m_tot = (Val(Left(m_hfin, 2)) - Val(Left(m_hini, 2))) * 60 + Val(Right(m_hfin, 2)) - Val(Right(m_hini, 2))   'minutos
                  rt!Valor = m_tot    'mints
               End If
               rt!documento = IIf(IsNull(rs!documento), "", rs!documento)
               rt.Update
            Else
            '-   MsgBox ("X-[" + cUbica + "]")
            End If
            rs.MoveNext
      Loop
      rs.Close
   Next i
   cmdBusca.Caption = "Hallo " + Trim(Str(nson)) + " reg."
   Beep 500, 350
   If nson = 0 Then
      MsgBox "No registra autorizaciones..", vbExclamation
   Else
      Call CmdExportaXLS_Click
   End If
   cmdBusca.Caption = ""
End Sub
Private Sub CmdLimpia_Click()
   Txtarjeta = "": Txtrabajador = ""
   CboOcurre.Text = "": Txcodigo = ""
   Cbo_DH.Text = "": TxDH = ""
   TxfechaI = "": TxfechaF = ""
   CboOcurre.Text = ""
   List1.Clear
   ccadena = "": cCargo = "": cDscCargo = "": ctipopla = ""
End Sub

Private Sub CmdExportaXLS_Click()
    If ChkXLS.Value = 0 Then
       wTpoQuery = "10"
       Call ExportTableToXLS_click
    Else
       wTpoQuery = "10"
       wNameXLS = "C:\RPT\SUMAutori.xls"
       Call Exportar_ADO_Excel
    End If
    Beep 500, 350
    MsgBox "Se generó C:\RPT\SUMAutori.xls", vbInformation
    Exit Sub
End Sub

Private Sub cmdsalir_Click()
   Set rt = Nothing
   Set rs = Nothing
   Set r_d = Nothing
   Set DD = Nothing
   Me.Hide
   Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set rt = Nothing
   Set rs = Nothing
   Set r_d = Nothing
   Set DD = Nothing
   Me.Hide
   Unload Me
End Sub



