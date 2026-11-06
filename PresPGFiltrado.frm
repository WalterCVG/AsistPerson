VERSION 5.00
Begin VB.Form PresPGFiltrado 
   Caption         =   "Presupuesto Pago de Guardia FILTRADO"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "PresPGFiltrado.frx":0000
   ScaleHeight     =   2385
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   8760
      Picture         =   "PresPGFiltrado.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cerrar"
      Top             =   1560
      Width           =   615
   End
   Begin VB.CheckBox ChkXLS 
      Alignment       =   1  'Right Justify
      Caption         =   "XLS instalado"
      Height          =   195
      Left            =   5400
      TabIndex        =   11
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton CmdExportaXLS 
      Height          =   615
      Left            =   7080
      Picture         =   "PresPGFiltrado.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Exporta a Xls"
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox TxFFin 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox TxFIni 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Txtarjeta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TxFecFin 
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
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxFecIni 
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
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   1095
      Left            =   6720
      Picture         =   "PresPGFiltrado.frx":80FF2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ubicar Registros"
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   8760
      Picture         =   "PresPGFiltrado.frx":811CD
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Otra Consulta"
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Del :"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Al :"
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
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
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
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "PresPGFiltrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim r_d As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim rt0 As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim f_Fer As Boolean
Dim nimporte As Double
Dim nvalor, nnroguardias As Integer
Dim sql, SQX, cTarjeta, cDocumento, cTurno, cGrpGen, cnivGuardia, ccadena, c_a, cano, cPatron As String
Dim c_Costo, cDscCosto, cCodUpr, cDscUpr, ccarpla, cDsccarpla As String
Dim unafecha As Date

Private Sub Form_Load()
  '-Pago Guardias HEP/RES
  'Cnx.Execute ("DELETE  FROM CPPGfiltrado ")
End Sub



Private Sub txfIni_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If IsDate(TxFIni) Then
        TxFFin = TxFIni
        TxFFin.SetFocus
        Exit Sub
     Else
        MsgBox "verifique  fechas inicial", vbCritical
        TxFIni = ""
     End If
  End If
End Sub
Private Sub txfFin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If IsDate(TxFFin) Then
        If Month(DateValue(TxFIni)) > Month(DateValue(TxFFin)) Then
            MsgBox "verifique  fechas ", vbExclamation
            Exit Sub
        End If
     Else
        MsgBox "verifique  fechas final", vbCritical
        TxFFin = ""
     End If
  End If
End Sub
Private Sub TxFFin_LostFocus()
   Call txfFin_KeyPress(13)
End Sub

Private Sub cmdcarga_Click()
    If Trim(TxFIni) = "" Or Trim(TxFFin) = "" Then
       MsgBox "No inidicò periodo a procesar ", vbExclamation, "[IMPORTANTE]"
       Exit Sub
    End If
     
    If Year(TxFIni) <> Year(TxFFin) Then
        MsgBox "Fechas no corresponden al mismo periodo ", vbExclamation, "[IMPORTANTE]"
        Exit Sub
     End If
    
    '-Nro Anos - meses
    canoIni = Str(Year(TxFIni))
    canoFin = Str(Year(TxFFin))
    cmesIni = Right("0" + Trim(Str(Month(TxFIni))), 2)
    cmesFin = Right("0" + Trim(Str(Month(TxFFin))), 2)
    
    On Error GoTo borra
    Cnx.BeginTrans
    Cnx.Execute ("SELECT * Into #TabPGFilt  FROM CpPGFiltrado WHERE 1 = 2")
    GoTo sigue
borra:
    Cnx.Execute "DELETE  FROM  #TabPGFilt "
sigue:
    Cnx.CommitTrans
      
    Set rt = New Recordset
    SQX = "SELECT * FROM #TabPGFilt"
    'SQX = "SELECT * FROM CpPGfiltrado"
    rt.Open SQX, Cnx, adOpenStatic, adLockOptimistic, adCmdText
                 
    Dim OpenForms
    v_an = 0
    nvalor = 0
    For a = Val(canoIni) To Val(canoFin)
       cano = Trim(Str(a))
       mi = "01"
       mf = "12"
       If a = Val(canoIni) Then
          mi = cmesIni
       End If
       If a = Val(canoFin) Then
          mf = cmesFin
       End If
       'MsgBox "van año " + cano + " meses del " + mi + " al " + mf
       For m = Val(mi) To Val(mf)
          cmes = Right("0" + Trim(Str(m)), 2)
          'MsgBox "inicia mes " + cano + "/" + cmes
          
          unafecha = "01/" + Trim(cmes) + "/" + cano
          Call ObtenerPrimerUltimoDia(unafecha)    'Obtiene primer/ultimo
      
          TxFecIni = primer
          TxFecFin = ultimo
          If a = Val(canoIni) And m = Val(mi) Then
             TxFecIni = TxFIni
          End If
          If a = Val(canoFin) And m = Val(mf) Then
             TxFecFin = TxFFin
          End If
          '-dias ordinarios y feriados  del mes
          ccadena = ""
          For j = Day(DateValue(primer)) To Day(DateValue(TxFecFin))
             unafecha = primer + j - 1
             f_Fer = False     ' verifica Feriado/domingo
             If UCase(Left(WeekdayName(Weekday(primer + j - 1)), 3)) = "DOM" Then
                f_Fer = True
                '''MsgBox Format(unafecha, "dd/mm/yyyy") + ":" + Str(j) + "-" + WeekdayName(Weekday(unafecha)) + "-" + Left(WeekdayName(Weekday(unafecha)), 3)
             Else
                Set rs = New Recordset
                rs.Open "SELECT * FROM Feriados ORDER BY codigo", Cnx0, adOpenStatic, adLockReadOnly
                rs.MoveFirst
                While Not rs.EOF
                   If Len(Trim(rs!codigo)) = 6 Then       'codigo:mmddaa
                      c_a = Right(rs!codigo, 2)
                   ElseIf Len(Trim(rs!codigo)) = 4 Then   'Codigo:mmdd
                          c_a = Right(Year(unafecha), 2)
                   End If
                   If Val(Left(rs!codigo, 2)) = Val(cmes) And Val(Mid(rs!codigo, 3, 2)) = j And Right(Year(unafecha), 2) = c_a Then
                      f_Fer = True
                      '''  MsgBox rs!codigo + "-" + Format(unafecha, "dd/mm/yyyy") + "-" + c_a
                   End If
                   rs.MoveNext
                Wend
                rs.Close
             End If
             ccadena = ccadena + IIf(f_Fer, "F", "O")
          Next j
          '''  MsgBox ccadena
     
          '-carga  DIRECTORIO "TURNO ASISTENCIALES"
          sql = "SELECT  * FROM CpDirectorio  WHERE ano='" & cano & "' AND mes='" & Trim(cmes) & "' and flag='0' and (PLH='HEP' or PLH='RES') " & _
                "ORDER BY  Tarjeta"
                
          Set r_d = New Recordset
          r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          Do While Not r_d.EOF
                '-dar control a windows y permitir multihilo
                v_an = v_an + 1
                If v_an Mod 10 = 0 Then     'if loop has repeated
                   OpenForms = DoEvents 'Yield to operating system
                End If
                cTarjeta = r_d!TARJETA
                '''  MsgBox "trabajador " + cano + "/" + cmes + "->" + cTarjeta
                
                c_Costo = IIf(IsNull(r_d!CCosto), "", r_d!CCosto)
                cCodUpr = IIf(IsNull(r_d!CodUpr), "", r_d!CodUpr)
                Set rs = New Recordset
                sql = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & cTarjeta & "'"
                rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                cplh = "": cDocumento = "": caPat = "": caMat = "": cnomb = "": cnivGuardia = "": cGrpGen = ""
                ccarpla = "": cDsccarpla = ""
                If Not rs.EOF Then
                   cplh = IIf(IsNull(rs!PLH), "", rs!PLH)                  'guardias  HEP/RES
                   cDocumento = IIf(IsNull(rs!Documento), "", Trim(rs!Documento))
                   caPat = Trim(rs!apelpat)
                   caMat = Trim(rs!apelmat)
                   cnomb = Trim(rs!nombres)
                   cnivGuardia = IIf(IsNull(rs!NivGuardia), "", rs!NivGuardia)
                   ccarpla = IIf(IsNull(rs!carPla), "", rs!carPla)
                   cGrpGen = IIf(IsNull(rs!grpgen), "", rs!grpgen)
                End If
                rs.Close
                
                   If c_Costo <> "" Then
                      Set rs = New Recordset
                      rs.Open "SELECT * FROM CpcCosto WHERE codigo='" & c_Costo & "'", Cnx, adOpenStatic, adLockReadOnly
                      If Not rs.EOF Then
                         cDscCosto = Left(rs!descripcion, 50)
                      End If
                      rs.Close
                   End If
                   If cCodUpr <> "" Then
                      Set rs = New Recordset
                      rs.Open "SELECT * FROM CpUndPrgRol WHERE codUpr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
                      If Not rs.EOF Then
                         cDscUpr = Left(rs!descripcion, 50)
                      End If
                      rs.Close
                   End If
                   If ccarpla <> "" Then
                      Set rs = New Recordset
                      rs.Open "SELECT * FROM Cptablas WHERE tabla='28' and codigo='" & ccarpla & "'", Cnx, adOpenStatic, adLockReadOnly
                      If Not rs.EOF Then
                         cDsccarpla = Left(rs!descripcion, 50)
                      End If
                      rs.Close
                   End If
                   nnroguardias = 0: ndo = 0: nno = 0: ndf = 0: nnf = 0
                   nTarDO = 0#: nTarNO = 0#: nTarDF = 0#: nTarNF = 0#
                   If cnivGuardia <> "" Then
                      Set rs = New Recordset
                      rs.Open "SELECT * FROM CpTarifas  WHERE (baja is NULL)", Cnx, adOpenStatic, adLockReadOnly
                      Do While Not rs.EOF
                         If InStr(rs!Niveles, cnivGuardia) > 0 Then
                            nTarDF = rs!DiaFer
                            nTarDO = rs!DiaOrd
                            nTarNF = rs!NocheFer
                            nTarNO = rs!NocheOrd
                            Exit Do
                         End If
                         rs.MoveNext
                      Loop
                      rs.Close
                   End If
                   '-Procesa Dias del Directorio
                    '''   MsgBox "tarj " + cTarjeta + "->" + Format(TxFecIni, "dd/mm/yyyy") + " al " + Format(TxFecFin, "dd/mm/yyyy")
                   
                   For i = Day(DateValue(TxFecIni)) To Day(DateValue(TxFecFin))
                       z = i - 1
                       cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", Trim(r_d.Fields(z + 4)))
                       cmdcarga.Caption = cTarjeta + ": " + cano + cmes + "/" + Str(i)
                         
                       '-Evalua Guardia -- check hay aut. VAC/COMP
                       p_horas = ""
                       If InStr("DNHCP", Left(cTurno, 1)) > 0 Then
                          '''MsgBox cTarjeta + " " + Str(i) + "-<<" + cTurno + ">>"
                          fechaB = Format(primer + z, "yyyymmdd")
                          V_a = "S"
                          '-Verifica si aplicó AUTORIZACION VAC/COMP
                          sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta='" & cTarjeta & "' and FechaIni<='" + fechaB + "' and FechaFin>='" + fechaB + "'"
                          Set rs = New Recordset
                          rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                          Do While Not rs.EOF
                             If Left(rs!codigo, 2) = "VA" Or rs!codigo = "COMP" Then
                                '-pierde Guardia equivale  6hra trabajo
                                V_a = "N"
                                Exit Do
                             End If
                             rs.MoveNext
                          Loop
                          rs.Close
                          
                          If V_a = "S" Then
                            '-Verifica HORARIO-horas turno Prog. Guardia
                            cvalor = ""
                            sql = "SELECT * FROM CpHor" & cano & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
                            Set rs = New Recordset
                            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                            If rs.EOF Then
                                 If Trim(cGrpGen) <> "" Then
                                    rs.Close
                                    sql = "SELECT * FROM CpHor" & cano & " WHERE grupo='" & cGrpGen & "' AND fecha ='" & fechaB & "'"
                                    Set rs = New Recordset
                                    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                                End If
                            End If
                            'If cTarjeta = "00110" Then
                            '   MsgBox Str(i) + "-" + cTurno + "->" + sql
                            'End If
                            If Not rs.EOF Then
                                cvalor = IIf(IsNull(rs!Horario), "", rs!Horario)
                                If cvalor <> "" Then
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
                                        p_horas = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
                                    End If
                                    rs_d.Close
                                End If
                            End If
                            rs.Close
                            If Val(p_horas) > 0 Then
                                cPatron = Mid(ccadena, i, 1)     ' F/O
                                If cPatron = "F" Then
                                    If Left(cTurno, 1) = "D" Then
                                        ndf = ndf + 1
                                    ElseIf InStr("NHC", Left(cTurno, 1)) > 0 Then
                                            nnf = nnf + 1
                                    ElseIf Left(cTurno, 1) = "P" Then
                                       ndf = ndf + 1
                                       nnf = nnf + 1
                                    End If
                                ElseIf cPatron = "O" Then
                                    If Left(cTurno, 1) = "D" Then
                                       ndo = ndo + 1
                                    ElseIf InStr("NHC", Left(cTurno, 1)) > 0 Then
                                           nno = nno + 1
                                    ElseIf Left(cTurno, 1) = "P" Then
                                           ndo = ndo + 1
                                           nno = nno + 1
                                    End If
                                End If
                                '''MsgBox cTarjeta + "-> DIA:" + Trim(Str(i)) + "<<" + cTurno + ">> HRS:" + p_horas + " TIPO:" + cPatron + " VAN-> DO:" + Trim(Str(ndo)) + ", NO:" + Trim(Str(nno)) + ", DF:" + Trim(Str(ndf)) + ", NF:" + Trim(Str(nnf))
                            End If
                          End If
                          
                       End If    ' --> fin dia guardia programada
              
                   Next i
                   nimporte = 0#
                   nnroguardias = ndo + nno + ndf + nnf
                   If nnroguardias > 0 Then
                      nimporte = ndo * nTarDO + nno * nTarNO + ndf * nTarDF + nnf * nTarNF
                      
                      rt.AddNew
                      rt!ano = cano
                      rt!mes = cmes
                      rt!CCosto = c_Costo
                      rt!Oficina = cDscCosto
                      rt!CodUpr = cCodUpr
                      rt!Servicio = cDscUpr
                      rt!TARJETA = cTarjeta
                      rt!Documento = cDocumento
                      rt!servidor = Left(caPat + " " + caMat + " " + cnomb, 45)
                      rt!cargo = ccarpla
                      rt!DscCargo = cDsccarpla
                      rt!Nivel = cnivGuardia
                      '-rt!TarifaDO = nTarDO
                      '-rt!TarifaNO = nTarNO
                      '-rt!TarifaDF = nTarDF
                      '-rt!TarifaNF = nTarNF
                      rt!guardiaDO = ndo
                      rt!guardiaNO = nno
                      rt!guardiaDF = ndf
                      rt!guardiaNF = nnf
                      rt!Monto = nimporte
                      rt!guardias = nnroguardias
                      rt.Update
                      
                      '''  MsgBox "Grabo: " + cTarjeta + " Periodo " + cano + "/" + cmes
                      nvalor = nvalor + 1
                   End If
                   
                r_d.MoveNext
          Loop
          r_d.Close
          'MsgBox "paso  " + cano + " / " + cmes
       Next m
       'MsgBox "completo   " + cano + " / " + cmes
    Next a
   
    cmdcarga.Caption = ""
      MsgBox "total registros procesados " + Str(nvalor)
    rt.Close
    Exit Sub
    '---Call Limpiacampos
End Sub

Private Sub CmdExportaXLS_Click()
   wTpoQuery = "1.5"
    If ChkXLS.Value = 0 Then
       Call ExportTableToXLS_click
    Else
       wNameXLS = "C:\RPT\PGFiltrado.xls"
       Call Exportar_ADO_Excel
    End If
       
    Beep 500, 350
    MsgBox "Generò C:\RPT\PGFiltrado.xls", vbInformation, "[IMPORTANTE]"
    Exit Sub
End Sub

Private Sub Limpiacampos()
   TxFIni = "": TxFFin = ""
   TxFecIni = "": TxFecFin = ""
   'Txtarjeta.Text = "": Txtarjeta.Enabled = False
   '-  Cnx.Execute ("DELETE  FROM CPPGfiltrado ")
   
End Sub
Private Sub CmdClear_Click()
   Call Limpiacampos
End Sub
Private Sub cmdsalir_Click()
   Set r_d = Nothing
   Set rs = Nothing
   Set rs_d = Nothing
   Set rt = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If UCase(Forms(cont).Name) <> "CPMENU" And UCase(Forms(cont).Name) <> "PRESPGFILTRADO" Then
           Unload Forms(cont)
        End If
      End If
   Next
   
   Me.Hide
   Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set r_d = Nothing
   Set rs = Nothing
   Set rs_d = Nothing
   Set rt = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "CpMenu" And UCase(Forms(cont).Name) <> "PRESPGFILTRADO" Then
           Unload Forms(cont)
        End If
      End If
   Next
   Me.Hide
   Unload Me
End Sub


