VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TblasNvoPeriodo 
   Caption         =   "Tablas Nuevo Periodo"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   LockControls    =   -1  'True
   Picture         =   "TblasNvoPeriodo.frx":0000
   ScaleHeight     =   2700
   ScaleWidth      =   5160
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Progreso1 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   0
      Picture         =   "TblasNvoPeriodo.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton CmdultDia 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Ult. Dia año anterior"
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "data 31/Dic Año  anterior"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   4320
      Picture         =   "TblasNvoPeriodo.frx":80EA8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cerrar"
      Top             =   1920
      Width           =   615
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton CmdProcede 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Procede"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Tablas y datos"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox TxPeriodo 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   0
      Top             =   480
      Width           =   675
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "TblasNvoPeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NvaTblaH, NvaTblaT, NvaTblaTM, NvaTblaA, NvaTblaM, NvaTblaMN, Sq0, sql As String
Dim nopcn_ano, n_mes As Integer
Dim unafecha As Date
Dim DfechaI, DfechaF As Date
Dim cTarjeta, cCodigo, cDiahora, cmMes, cTurno, choraini, chorafin As String
Dim lok As Boolean
Private Sub cmdsalir_Click()
  Unload Me
  Me.Hide
End Sub
Private Sub Command1_Click()
     c_grupo = "05": c_secuencia = "052"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub

Private Sub Form_Load()
  CmdultDia.Caption = "Ult. día año " + genaño
  '-periodo que inicia
  TxPeriodo = Val(genaño) + 1
  lok = True
End Sub

Private Sub UpDown1_DownClick()
    If Val(TxPeriodo) <= 1980 Then
       TxPeriodo = 2090
       Exit Sub
    End If
    TxPeriodo = Val(TxPeriodo) - 1
    genaño = Trim(TxPeriodo)
End Sub
Private Sub UpDown1_UpClick()
    If Val(TxPeriodo) >= 2090 Then
       TxPeriodo = 1980
       Exit Sub
    End If
    TxPeriodo = Val(TxPeriodo) + 1
    genaño = Trim(TxPeriodo)
End Sub
Private Sub CmdProcede_Click()
  
  n_ano = Year(Date)
  n_mes = Month(Date)
  If Val(TxPeriodo) - 1 > n_ano Then
     lok = False
  End If
  If Not lok Then
     MsgBox "Periodo incorrecto... ", vbCritical, "[IMPORTANTE]"
     Exit Sub
  End If
  If n_mes < 11 Then
     lok = False
  End If
  If Not lok Then
     If MsgBox("Ejecutará, antes del periodo recomendado (2meses) ?   ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
     End If
  End If

  NvaTblaM = "CpMarca" & TxPeriodo
  '--- hasta 2013 NvaTblaMN = "CpManual" & TxPeriodo
  NvaTblaA = "CpAutori" & TxPeriodo
  NvaTblaT = "CpChgTur" & TxPeriodo
  NvaTblaTM = "CpChgTurM" & TxPeriodo
  NvaTblaH = "CpHor" & TxPeriodo
  Dim rs As Recordset
  Set rs = New ADODB.Recordset
  Sq0 = "SELECT * FROM " & NvaTblaM
  On Error GoTo Proceso
  rs.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    MsgBox "Verifique... Periodo Ya Existe ", vbExclamation
    rs.Close
  GoTo Terminar
Proceso:
  CreaTablas
Terminar:
End Sub
Sub CreaTablas()
 'Crea tablas new periodo
 Cnx.CommandTimeout = 300
 sql = "SELECT * INTO " & NvaTblaM & " FROM CpMarca" & genaño & " WHERE 1=2"
 Set rD = Cnx.Execute(sql)
 '-hasta 2013 sql = "SELECT * INTO " & NvaTblaMN & " FROM CpManual" & genaño & " WHERE 1=2"
 'Set rD = Cnx.Execute(sql)
 sql = "SELECT * INTO " & NvaTblaA & " FROM CpAutori" & genaño & " WHERE 1=2"
 Set rD = Cnx.Execute(sql)
 sql = "SELECT * INTO " & NvaTblaT & " FROM CpChgTur" & genaño & " WHERE 1=2"
 Set rD = Cnx.Execute(sql)
 sql = "SELECT * INTO " & NvaTblaTM & " FROM CpChgTurM" & genaño & " WHERE 1=2"
 Set rD = Cnx.Execute(sql)
 sql = "SELECT * INTO " & NvaTblaH & " FROM CpHor" & genaño & " WHERE 1=2"
 Set rD = Cnx.Execute(sql)
  
 For nopc = 1 To 5
     Call procesadata
 Next
 Set Cnx = Nothing
 MsgBox "CONCLUIDO... se Creo tablas del " + TxPeriodo.Text, vbExclamation
 End
  
End Sub

Public Sub procesadata()
  Dim nreg As Integer
  '-Incluye mov ult dia periodo anterior
  unafecha = "31/12/" + Trim(Str(genaño))
  Call ObtenerPrimerUltimoDia(unafecha)
  fecha1 = Format(ultimo, "yyyymmdd")
  On Error GoTo error
  Select Case nopc
    Case 1
      Sq0 = "SELECT * FROM Cphor" & genaño & " WHERE  fecha='" & fecha1 & "'"
      sql = "SELECT * FROM " & NvaTblaH
    Case 2
      Sq0 = "SELECT * FROM CpChgTur" & genaño & " WHERE  fecha='" & fecha1 & "' OR Nfecha='" & fecha1 & "'"
      sql = "SELECT * FROM " & NvaTblaT
    Case 3
      Sq0 = "SELECT * FROM CpChgTurM" & genaño & " WHERE  fecha='" & fecha1 & "'"
      sql = "SELECT * FROM " & NvaTblaTM
    Case 4
      '-Registros que inician hasta 31/12  y terminen sgte año
      Sq0 = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and  fechaINI<='" & fecha1 & "' AND fechaFIN>='" & fecha1 & "'"
      sql = "SELECT * FROM " & NvaTblaA
    Case 5
      Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE  fecha='" & fecha1 & "'"
      sql = "SELECT * FROM " & NvaTblaM
    '-hasta 2013 Case 6
      'Sq0 = "SELECT * FROM CpManual" & genaño & " WHERE  fecha='" & fecha1 & "'"
      'sql = "SELECT * FROM " & NvaTblaMN
  End Select
  
  Set rs = New Recordset
  rs.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  If Not rs.EOF Then
     'nreg = rs.RecordCount
     'Progreso1.Value = 0
     'Progreso1.Max = nreg
     rs.MoveFirst
     Set Rst = New Recordset
     Rst.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
     Do While Not rs.EOF
        'With Progreso1
        '  If .Value < .Max Then
        '     .Value = .Value + 1
        '  End If
        'End With
        Rst.AddNew
        For i = 0 To rs.Fields.Count - 1
            On Error Resume Next
            'cCampo = rs.Fields(i).Name
            Select Case Rst.Fields(i).Type
              Case dbNumeric    'Cuando el campo de destino es un Autonumérico, no hace nada.
                Rst.Fields(i) = 0 + rs.Fields(i)
              Case dbText
                Rst.Fields(i) = "" & rs.Fields(i)
              Case Else
                Rst.Fields(i) = rs.Fields(i)
            End Select
        Next i
        Rst.UpdateBatch
        rs.MoveNext
     Loop
     Rst.Close
  End If
  rs.Close
  Exit Sub
error:
  MsgBox Err.Description & " Error: " & Err.Number & vbCr
End Sub

Private Sub CmdUltDia_Click()
  '-ultimo dia del año con datos a transferir
  unafecha = "31/12/" + Trim(Str(genaño))  '31/12/2020
  Call ObtenerPrimerUltimoDia(unafecha)
  fecha1 = Format(ultimo, "yyyymmdd")       '20201231
  If Val(TxPeriodo) <= genaño Then          '2021 <= 2020
     '-datos a transferir no pueden provenir de año >=
     Beep 500, 350
     MsgBox "<Acción Indebida>.. Nuevo Periodo(destino), No puede ser Menor/Igual al año de trabajo", vbCritical
     End
  End If
  
  If ultimo >= Date Then        '31/12/2020  >= fecha HOY
     Beep 500, 350
     If MsgBox("Transfiere datos del " + genaño + " al año " + TxPeriodo + " ... ", vbOKCancel + vbQuestion, "Confirme") = vbCancel Then
        End
     End If
  End If
  
  NvaTblaH = "CpHor" & TxPeriodo           'Nvo Periodo ==> genaño +1
  NvaTblaT = "CpChgTur" & TxPeriodo
  NvaTblaTM = "CpChgTurM" & TxPeriodo
  NvaTblaA = "CpAutori" & TxPeriodo
  NvaTblaM = "CpMarca" & TxPeriodo
  '- hasta 2013 NvaTblaMN = "CpManual" & TxPeriodo
  Dim rs As Recordset
  Set rs = New ADODB.Recordset
  Sq0 = "SELECT * FROM " & NvaTblaM
  On Error GoTo errorC
  rs.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  rs.Close
 Dim nreg As Integer
  '-Mov. Ultimo Dia AÑO ANTERIOR => periodo con que "Ingreso" [genaño] el cual transfiere datos al periodo "Destino" (Nvo Periodo) [TxPeriodo]
  For nopc = 1 To 5
   Select Case nopc
    Case 1
      '-Hor. Programado x Servidor, Retira del año NUEVO, DATOS cuya fecha <= 31/12/Ultimo dia del año Anterior
      SQX = "DELETE FROM " & NvaTblaH & " WHERE fecha<='" & fecha1 & "'"
      '-Adiciona DATOS del año Anterior, cuya fecha = ultimo Dia del año
      Sq0 = "SELECT * FROM Cphor" & genaño & " WHERE  fecha='" & fecha1 & "'"    ' adiciona datos = 31/12/año Anterior
      '-Abre tabla Año Actual
      sql = "SELECT * FROM " & NvaTblaH
    Case 2
      SQX = "DELETE FROM " & NvaTblaT & " WHERE fecha='" & fecha1 & "' OR Nfecha='" & fecha1 & "'"
      Sq0 = "SELECT * FROM CpChgTur" & genaño & " WHERE  fecha='" & fecha1 & "' OR Nfecha='" & fecha1 & "'"
      sql = "SELECT * FROM " & NvaTblaT
    Case 3
      SQX = "DELETE FROM " & NvaTblaTM & " WHERE fecha='" & fecha1 & "'"
      Sq0 = "SELECT * FROM CpChgTurM" & genaño & " WHERE  fecha='" & fecha1 & "'"
      sql = "SELECT * FROM " & NvaTblaTM
    Case 4
      '-registros Ult. Dia y todos aquellos que finalizan Nvo Año
      'SQX = "DELETE FROM " & NvaTblaA & " WHERE fechaINI <='" & Fecha1 & "' and fechaFIN>='" & Fecha1 & "'"
      SQX = ""
      Sq0 = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and  (fechaINI<='" & fecha1 & "' AND fechaFIN>='" & fecha1 & "') OR fechaFIN>='" & fecha1 & "'"
        ''  MsgBox "4datos-" + sq0
          
      sql = "SELECT * FROM " & NvaTblaA
    Case 5
      SQX = "DELETE FROM " & NvaTblaM & " WHERE fecha <='" & fecha1 & "'"
      Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE  fecha='" & fecha1 & "'"
      sql = "SELECT * FROM " & NvaTblaM
    '-Hasta 2013 Case 6
      'SQX = "DELETE FROM " & NvaTblaMN & " WHERE fecha<='" & fecha1 & "'"
      'Sq0 = "SELECT * FROM CpManual" & genaño & " WHERE  fecha='" & fecha1 & "'"
      'sql = "SELECT * FROM " & NvaTblaMN
   End Select
   If SQX <> "" Then
      '-Retira reg de año Anterior
      Cnx.Execute SQX
   End If
   Set rs = New Recordset
   rs.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      rs.MoveFirst
      If nopc = 4 Then
         Do While Not rs.EOF
            cTarjeta = rs!tarjeta
            cCodigo = rs!Codigo
            cDiahora = rs!diahora
            cTurno = IIf(IsNull(rs!turno), "", rs!turno)
            DfechaI = rs!fechaIni
            DfechaF = rs!fechaFin
            fecha2 = Format(rs!fechaIni, "yyyymmdd")
            fecha3 = Format(rs!fechaFin, "yyyymmdd")
            If cDiahora = "D" Then
               sql = "SELECT * FROM " & NvaTblaA & " WHERE  tarjeta='" & cTarjeta & "' AND codigo='" & cCodigo & "' AND diahora='" & cDiahora & "' AND fechaini='" & fecha2 & "'  AND fechaFin='" & fecha3 & "'"
               ''If cTarjeta = "59112" Then
               ''    MsgBox "4.1busca-" + sql
               ''End If
               
            Else
               choraini = rs!horaini
               chorafin = rs!horafin
               sql = "SELECT * FROM " & NvaTblaA & " WHERE  tarjeta='" & cTarjeta & "' AND codigo='" & cCodigo & "' AND diahora='" & cDiahora & "' AND fechaini='" & fecha2 & "'  AND fechaFin='" & fecha3 & "' AND Horaini='" & choraini & "' AND HoraFin='" & chorafin & "'"
            End If
            Set Rst = New Recordset
            Rst.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
              
            If Rst.EOF Then
               ''If cTarjeta = "59112" Then
               ''    MsgBox "4.2noUbico-" + cCodigo
               ''End If
            
               Rst.AddNew
               For i = 0 To rs.Fields.Count - 1
                  On Error Resume Next
                  'cCampo = rs.Fields(i).Name
                  Select Case Rst.Fields(i).Type
                    Case dbNumeric    'Cuando el campo de destino es un Autonumérico, no hace nada.
                      Rst.Fields(i) = 0 + rs.Fields(i)
                    Case dbText
                      Rst.Fields(i) = "" & rs.Fields(i)
                   Case Else
                     Rst.Fields(i) = rs.Fields(i)
                 End Select
              Next i
              Rst.UpdateBatch
              
              ''If cTarjeta = "59112" Then
              ''   MsgBox "4.3-GRABO "
              ''End If
              
            End If
            Rst.Close
            'Check Prog TURNO Directorio Vaca/lic
            'If cTurno <> "" And Trim(cDiaHora) = "D" Then    'Programa Turno
            '   Call ProgTURNO
            'End If
            rs.MoveNext
         Loop
         
      Else
         Set Rst = New Recordset
         Rst.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         Do While Not rs.EOF
            Rst.AddNew
            For i = 0 To rs.Fields.Count - 1
               On Error Resume Next
               'cCampo = rs.Fields(i).Name
               Select Case Rst.Fields(i).Type
                 Case dbNumeric    'Cuando el campo de destino es un Autonumérico, no hace nada.
                   Rst.Fields(i) = 0 + rs.Fields(i)
                 Case dbText
                   Rst.Fields(i) = "" & rs.Fields(i)
                 Case Else
                   Rst.Fields(i) = rs.Fields(i)
               End Select
            Next i
            Rst.UpdateBatch
            rs.MoveNext
         Loop
         Rst.Close
      End If
   End If
   rs.Close
 Next nopc
 MsgBox "Completo Carga Datos para el " + TxPeriodo.Text, vbExclamation
 End
 Exit Sub
errorC:
  MsgBox Err.Description & " Error: " & Err.Number & vbCr

End Sub

Private Sub ProgTURNO()
   fecha1 = Format(DfechaI.Text, "yyyymmdd")
   cmMes = Right("0" + Trim(Str(Month(DfechaI))), 2)
   fecha2 = Format(DfechaI.Text, "yyyymmdd")
   L_ok = True
   '-Programa Directorio,  ==> tarjeta,txfechaini,txfechafin
   If Year(DfechaI) = Val(TxPeriodo.Text) And Month(DfechaI) = Month(DfechaF) Then
      sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & cTarjeta & "' AND ano ='" & genaño & "' AND mes='" & cnmes + "'"
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      'Verifica si existe, Registros Programacion
      nson = 0
      If Not rs_d.EOF Then
         For i = Day(DfechaI) To Day(DfechaF)
            z = i - 1
            If Trim(rs_d.Fields(z + 4)) = "" Or IsNull(rs_d.Fields(z + 4)) Then
            Else
               ccadena = ccadena + Trim(Str(i)) + "_"
               nson = nson + 1
            End If
         Next i
      End If
      If nson > 0 Then
         MsgBox "Existe Directorio, Dias: " + ccadena, vbExclamation
         L_ok = False
      End If
      rs_d.Close
   ElseIf Month(DfechaI) <> Month(DfechaF) Then
          nmesI = Month(DfechaI)
          nmesF = Month(DfechaF)
          nmesA = 0
          ndifdias = DateValue(DfechaF) - DateValue(DfechaI) + 1
          For j = 0 To ndifdias - 1
             unafecha = DateValue(TxfechaIni.Text) + j
             nmes = Month(unafecha)
             nano = Year(unafecha)
             If nmesA <> nmes Then
                nmesA = nmes
                Call ObtenerPrimerUltimoDia(unafecha)
                If nmes = nmesI Then
                   primer = DateValue(DfechaI)
                End If
                If nmes = nmesF Then
                   ultimo = DateValue(DfechaF)
                End If
                cmes = Right("00" + Trim(Str(nmes)), 2)
                cano = Trim(Str(nano))
                sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & cTarjeta & "' AND ano ='" & cano & "' AND mes='" & cmes & "'"
                Set rs_d = New Recordset
                rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
                'Verifica si existen Registros previos
                nson = 0
                If Not rs_d.EOF Then
                   For i = Day(TxfechaIni) To Day(TxFechaFin)
                      z = i - 1
                      If Trim(rs_d.Fields(z + 4)) = "" Or IsNull(rs_d.Fields(z + 4)) Then
                      Else
                         ccadena = ccadena + Trim(Str(i)) + "_"
                         nson = nson + 1
                      End If
                   Next i
                End If
                If nson > 0 Then
                   MsgBox "Existe Directorio, Dias: " + ccadena, vbExclamation
                   L_ok = False
                End If
             End If
          Next j
   End If
   
   If L_ok Then
      '-Programa cambio Turno en Directorio  ==> tarjeta,txfechaini,txfechafin
      If Year(DfechaI) = Val(Txmes.Text) And Month(TxfechaIni.Text) = Month(TxFechaFin.Text) Then
         sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
         Set rs_d = New Recordset
         rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If rs_d.EOF Then
            Sw = True
            rs_d.AddNew
            rs_d!tarjeta = Txtarjeta.Text
            rs_d!ano = genaño
            rs_d!mes = Txmes.Text
            rs_d!usuario = wNick      'Firma Usuario
            rs_d!fecha = Date
         End If
         TxHoras.Text = "": TxD.Text = "": TxN.Text = "": TxM.Text = "": Txt.Text = ""
         nhoras = 0: nd = 0: nn = 0: nm = 0: nt = 0
         For i = Day(TxfechaIni) To Day(TxFechaFin)
             z = i - 1
             If Trim(rs_d.Fields(z + 4)) = "" Or IsNull(rs_d.Fields(z + 4)) Then
             Else
                fecha1 = Format(DateValue(TxfechaIni.Text) + z, "yyyymmdd")
                If Trim(rs_d.Fields(z + 4)) = "D" Or InStr("D*D1D2D3D4", Trim(rs_d.Fields(z + 4))) > 0 Then
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
                sql1 = "UPDATE CpHor" & genaño & " SET Horario='O' WHERE tarjeta='" + Txtarjeta & "' and fecha='" & fecha1 & "'"
                Cnx.Execute sql1
             End If
             rs_d.Fields(z + 4) = TxTurno.Text
         Next i
         If nd + nn + nm + nt > 0 Then
            TxD.Text = IIf(IsNull(rs_d!d), 0, Val(rs_d!d)) - nd
            TxN.Text = IIf(IsNull(rs_d!n), 0, Val(rs_d!n)) - nn
            TxM.Text = IIf(IsNull(rs_d!m), 0, Val(rs_d!m)) - nm
            Txt.Text = IIf(IsNull(rs_d!T), 0, Val(rs_d!T)) - nt
            TxHoras.Text = IIf(IsNull(rs_d!Horas), 0, rs_d!Horas) - ((nd + nn) * 12 + (nm + nt) * 6)
            rs_d!d = Trim(TxD.Text)
            rs_d!n = Trim(TxN.Text)
            rs_d!m = Trim(TxM.Text)
            rs_d!T = Trim(Txt.Text)
            rs_d!Horas = Val(TxHoras.Text)
         End If
         rs_d!carEst = IIf(IsNull(TxCarEst.Text), "", TxCarEst.Text)
         rs_d!CCosto = IIf(IsNull(Txccosto.Text), "", Txccosto.Text)
         rs_d!CodUpr = IIf(IsNull(TxcodUPR.Text), "", TxcodUPR.Text)
         rs_d!Cargo = IIf(IsNull(TxCargo.Text), "", TxCargo.Text)
         rs_d!PLH = IIf(IsNull(TxPlH.Text), "", TxPlH.Text)
         rs_d!Flag = IIf(Txtipopla.Text = "ADM", "1", IIf(Txtipopla.Text = "ASI", "0", "0"))   'administrativo/Asistencial
         rs_d!usuarioM = wNick      'Firma Usuario
         rs_d!fechaM = Date
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
                      rs_d.AddNew
                      rs_d!tarjeta = Txtarjeta.Text
                      rs_d!ano = cano
                      rs_d!mes = cmes
                      rs_d!usuario = wNick      'Firma Usuario
                      rs_d!fecha = Date
                   End If
                   TxHoras.Text = "": TxD.Text = "": TxN.Text = "": TxM.Text = "": Txt.Text = ""
                   nhoras = 0: nd = 0: nn = 0: nm = 0: nt = 0
                   nson = 0
                   For i = Day(primer) To Day(ultimo)
                       z = i - 1
                       nson = nson + 1
                       If Trim(rs_d.Fields(z + 4)) = "" Or IsNull(rs_d.Fields(z + 4)) Then
                       Else
                          fecha1 = Format(primer + z, "yyyymmdd")
                          If Trim(rs_d.Fields(z + 4)) = "D" Or InStr("D*D1D2D3D4", Trim(rs_d.Fields(z + 4))) > 0 Then
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
                          sql1 = "UPDATE CpHor" & genaño & " SET Horario='O' WHERE tarjeta='" + Txtarjeta & "' and fecha='" & fecha1 & "'"
                          Cnx.Execute sql1
                       End If
                       rs_d.Fields(z + 4) = TxTurno.Text
                   Next i
                   j = j + (nson - 1)
                   If nd + nn + nm + nt > 0 Then
                      TxD.Text = IIf(IsNull(rs_d!d), 0, Val(rs_d!d)) - nd
                      TxN.Text = IIf(IsNull(rs_d!n), 0, Val(rs_d!n)) - nn
                      TxM.Text = IIf(IsNull(rs_d!m), 0, Val(rs_d!m)) - nm
                      Txt.Text = IIf(IsNull(rs_d!T), 0, Val(rs_d!T)) - nt
                      TxHoras.Text = IIf(IsNull(rs_d!Horas), 0, rs_d!Horas) - ((nd + nn) * 12 + (nm + nt) * 6)
                      rs_d!d = Trim(TxD.Text)
                      rs_d!n = Trim(TxN.Text)
                      rs_d!m = Trim(TxM.Text)
                      rs_d!T = Trim(Txt.Text)
                      rs_d!Horas = Val(TxHoras.Text)
                   End If
                   rs_d!carEst = IIf(IsNull(TxCarEst.Text), "", TxCarEst.Text)
                   rs_d!CCosto = IIf(IsNull(Txccosto.Text), "", Txccosto.Text)
                   rs_d!CodUpr = IIf(IsNull(TxcodUPR.Text), "", TxcodUPR.Text)
                   rs_d!Cargo = IIf(IsNull(TxCargo.Text), "", TxCargo.Text)
                   rs_d!PLH = IIf(IsNull(TxPlH.Text), "", TxPlH.Text)
                   rs_d!Flag = IIf(Txtipopla.Text = "ADM", "1", IIf(Txtipopla.Text = "ASI", "0", "0"))   'administrativo/Asistencial
                   rs_d!usuario = wNick      'Firma Usuario
                   rs_d!fecha = Date
                   rs_d.Update
                   rs_d.Close
                End If
             Next j
      End If
   End If
   Cnx.CommitTrans
   
   L_ok = True
   If Trim(TxDiaHora.Text) = "D" And Trim(TxTurno.Text) <> "" Then
      'Valida Cambio Turno en Directorio  ==> tarjeta,txfechaini,txfechafin
      If Month(TxfechaIni.Text) = Val(Txmes.Text) And Month(TxfechaIni.Text) = Month(TxFechaFin.Text) Then
         sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & Txtarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'"
         Set rs_d = New Recordset
         rs_d.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         For i = Day(TxfechaIni) To Day(TxFechaFin)
            z = i - 1
            If Trim(rs_d.Fields(z + 4)) = TxTurno.Text Then
            Else
               L_ok = False
            End If
         Next
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
                   nson = 0
                   For i = Day(TxfechaIni) To Day(TxFechaFin)
                      z = i - 1
                      If Trim(rs_d.Fields(z + 4)) = TxTurno.Text Then
                      Else
                         L_ok = False
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
   
End Sub
