VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ResPgoGuarAbon 
   Caption         =   "Resolucion - Dif. pago de Guardia - ABONO"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "ResPgoGuarAbon.frx":0000
   ScaleHeight     =   2850
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   8640
      Picture         =   "ResPgoGuarAbon.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cerrar"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox ChkXLS 
      Alignment       =   1  'Right Justify
      Caption         =   "XLS instalado"
      Height          =   195
      Left            =   6720
      TabIndex        =   24
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1335
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
      Left            =   2640
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
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
      Left            =   1800
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
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
      Left            =   2640
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
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
      Left            =   1800
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxFechaR0 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6000
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TxFecReso0 
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
      Left            =   5040
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxNroReso0 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      MaxLength       =   15
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TxNroResolu 
      Height          =   285
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxMesAno 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdExportaXLS 
      Height          =   615
      Left            =   6960
      Picture         =   "ResPgoGuarAbon.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exporta a Xls"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      MaxLength       =   2
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxMesdeAno 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxFecActual 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxMesActual 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   4200
      TabIndex        =   17
      Top             =   1800
      Width           =   1815
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
         Left            =   960
         Picture         =   "ResPgoGuarAbon.frx":80FF2
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprimir"
         Top             =   240
         Width           =   765
      End
      Begin VB.CommandButton mnuFicImpSelec 
         Height          =   675
         Left            =   120
         Picture         =   "ResPgoGuarAbon.frx":818BC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Selec.Impresora"
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   975
      Left            =   6120
      Picture         =   "ResPgoGuarAbon.frx":8209A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Ubicar Registros"
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   8760
      Picture         =   "ResPgoGuarAbon.frx":82275
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Otra Consulta"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "ResPgoGuarAbon.frx":82B3F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. Resolucion"
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "ResPgoGuarAbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim r_d As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim f_Fer As Boolean
Dim nLacta, nSonLacta, nposicion, nvalor As Integer
Dim csexo, cAplica, cTarjeta, cDni, sql, SQ1, c_a, cNroReso0, cplh, cTurno, cGrpGen, c_Costo, cDscCosto, cCodUpr, cDscUpr, ccadhep, cDcadhep, ccarpla, cDcarpla, cnivGuardia, ccadena, cPatron As String
Dim unafecha As Date
Dim fecReso0 As Date
Private Sub Command1_Click()
    c_grupo = "04": c_secuencia = "046"
    Load AydAcerca
    AydAcerca.Left = 0: AydAcerca.Top = 0
    AydAcerca.Show
End Sub
Private Sub Form_Load()
  'Aplica asistenciales HEP / RES
  TxMesActual = Trim(Mid(wMeses, (Month(Date) - 1) * 9 + 1, 9))
  ChkXLS.Value = 1
  TxFecActual = "La Victoria, " + Right$("00" + Trim(Str(Day(Date))), 2) + " de " + TxMesActual + " de " & Year(Date)

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
         TxNombMes.Text = Left(CboMes.List(j), nposicion - 1)
         TxMesdeAno.Text = TxNombMes.Text + " del " & genaño
         TxMesAno.Text = TxNombMes.Text + " " & genaño
         unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
         Call ObtenerPrimerUltimoDia(unafecha)
         Txfecini = primer
         TxFecFin = ultimo
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
            Txmes.Text = Right(CboMes.List(j), nlen - nposicion)
            TxNombMes.Text = Left(CboMes.List(j), nposicion - 1)
            TxMesdeAno.Text = TxNombMes.Text + " del " & genaño
            TxMesAno.Text = TxNombMes.Text + " " & genaño
            unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)
            Txfecini = primer
            TxFecFin = ultimo
         End If
      End If
      Exit Sub
   End If
End Sub
Private Sub cmdcarga_Click()
   If Txmes.Text = "" Then
      MsgBox "No Selecciono mes ", vbExclamation
      CboMes.SetFocus
      Exit Sub
   End If
   If TxNroResolu.Text = "" Then
      MsgBox "No Indico Nro. Resolución.. ", vbExclamation
      TxNroResolu.SetFocus
      Exit Sub
   End If
   ' -*
   sql = "SELECT count(*) AS Tot FROM CpDirectorio WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes) & "'and Nroresolu <>'' and flag='0' "
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   nson = rs!tot
   rs.Close
   If nson = 0 Then
      MsgBox "No Existe, Resolución  Pago de Planillas ...", vbExclamation
      Exit Sub
   End If
   
   unafecha = "01/" + Txmes.Text + "/" + Trim(Str(genaño))
   Call ObtenerPrimerUltimoDia(unafecha)
   Txfecini = primer
   TxFecFin = ultimo
   
   Cnx.Execute "DELETE  FROM CpTmDifPgoGuar"
   Set rt = New Recordset
   SQ1 = "SELECT * FROM CpTmDifPgoGuar"
   rt.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   ccadena = ""
   For i = Day(primer) To Day(ultimo)
      f_Fer = False     ' verifica Feriado/domingo
      '''If Weekday(unafecha + i - 1) = 1 Then
      If UCase(Left(WeekdayName(Weekday(primer + i - 1)), 3)) = "DOM" Then
         f_Fer = True
      Else
         Set rs = New Recordset
         rs.Open "SELECT * FROM Feriados ORDER BY codigo", Cnx0, adOpenStatic, adLockReadOnly
         rs.MoveFirst
         While Not rs.EOF
            If Len(Trim(rs!Codigo)) = 6 Then       'codigo:mmddaa
               c_a = Right(rs!Codigo, 2)
            ElseIf Len(Trim(rs!Codigo)) = 4 Then   'Codigo:mmdd
                   c_a = Right(Year(unafecha), 2)
            End If
            If Val(Left(rs!Codigo, 2)) = Val(Txmes.Text) And Val(Mid(rs!Codigo, 3, 2)) = i And Right(Year(unafecha), 2) = c_a Then
               f_Fer = True
            End If
            rs.MoveNext
         Wend
         rs.Close
      End If
      ccadena = ccadena + IIf(f_Fer, "F", "O")
   Next i
   
   '-Carga  DIRECTORIO "TURNO"
   sql = "SELECT  * FROM CpDirectorio  WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes) & "' and flag='0' " & _
        "ORDER BY Tarjeta, ano, mes "
   Set r_d = New Recordset
   r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not r_d.EOF Then
     nvalor = 0
     cNroReso0 = Trim(r_d!nroresolu)
     fecReso0 = r_d!Fecresolu
     Do While Not r_d.EOF
        cDni = "": cplh = "": caPat = "": caMat = "": cnomb = "": cnivGuardia = "": cGrpGen = "": c_Costo = "": cDscCosto = "": cCodUpr = "": cDscUpr = ""
        ccadhep = "": cDcadhep = "": ccarpla = "": cDcarpla = "":  cDni = ""
        cTarjeta = r_d!tarjeta
        c_Costo = IIf(IsNull(r_d!CCosto), "", r_d!CCosto)
        cCodUpr = IIf(IsNull(r_d!CodUpr), "", r_d!CodUpr)
           
           '-SABER si una columna existe en una tabla
           SQ1 = "SELECT * FROM  INFORMATION_SCHEMA.COLUMNS AS c1  WHERE  c1.column_name = 'cadhep'  AND c1.table_name = 'cpdirectorio'"
           Set rs_2 = New Recordset
           rs_2.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
           If Not rs_2.EOF Then
              ccadhep = IIf(IsNull(r_d!CadHep), "", Trim(r_d!CadHep))
           End If
           rs_2.Close
        
        Set rs = New Recordset
        sql = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & cTarjeta & "'"
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If rs.EOF Then
        Else
           cplh = IIf(IsNull(rs!PLH), "", rs!PLH)
           cDni = IIf(IsNull(rs!documento), "", rs!documento)
           caPat = Trim(rs!apelpat)
           caMat = Trim(rs!apelmat)
           cnomb = Trim(rs!nombres)
           csexo = IIf(IsNull(rs!sexo), "", rs!sexo)
           cnivGuardia = IIf(IsNull(rs!NivGuardia), "", rs!NivGuardia)
           ccarpla = IIf(IsNull(rs!carPla), "", rs!carPla)
           cGrpGen = IIf(IsNull(rs!grpgen), "", rs!grpgen)
           If ccadhep = "" Then
              ccadhep = IIf(IsNull(rs!CadHep), "", rs!CadHep)
           End If
        End If
        rs.Close
                
        '-Afecto X LACTANCIA
        '=============================
        cAplica = "N": nSonLacta = 0
        TxIniLac = "": TxFinLac = ""
      
        If cplh = "HEP" Or cplh = "RES" Then
           If c_Costo <> "" Then
               Set rs_d = New Recordset
               rs_d.Open "SELECT * FROM Cpccosto WHERE  codigo='" & c_Costo & "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rs_d.EOF Then
                  cDscCosto = Trim(rs_d!descripcion)
               End If
               rs_d.Close
           End If
           If cCodUpr <> "" Then
               Set rs_d = New Recordset
               rs_d.Open "SELECT * FROM CpUndPrgRol WHERE  codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rs_d.EOF Then
                  cDscUpr = Trim(rs_d!descripcion)
               End If
               rs_d.Close
           End If
           If ccadhep <> "" Then
               Set rs_d = New Recordset
               rs_d.Open "SELECT * FROM Cpcadena WHERE ano='" & genaño & "' and cadena='" & ccadhep & "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rs_d.EOF Then
                  cDcadhep = Trim(rs_d!descripcion)
               End If
               rs_d.Close
           End If
           If ccarpla <> "" Then
               Set rs_d = New Recordset
               rs_d.Open "SELECT * FROM Cptablas WHERE tabla='28' and codigo='" & ccarpla & "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rs_d.EOF Then
                  cDcarpla = Trim(rs_d!descripcion)
               End If
               rs_d.Close
           End If
           nTarDO = 0#: nTarNO = 0#: nTarDF = 0#: nTarNF = 0#
           If cnivGuardia <> "" Then
              Set rs = New Recordset
              rs.Open "SELECT * FROM CpTarifas WHERE (baja is NULL) ", Cnx, adOpenStatic, adLockReadOnly
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
           ndo = 0: nno = 0: ndf = 0: nnf = 0: nmonto = 0
           cdiaant = "N"
           For i = Day(primer) To Day(ultimo)
              z = i - 1
              cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", Trim(r_d.Fields(z + 4)))
              cmdcarga.Caption = cTarjeta + " " + Str(i)
              '<--Evalua Turno
              If (InStr("1DN4HCP", Left(cTurno, 1)) > 0 And InStr("MT_MX_NA_PA", cTurno) = 0) And cTurno <> "*" And Trim(cTurno) <> "" Then
                 F_falta = False
                 
                 '-Directorio REAL---Verifica Asistencia
                 fechaB = Format(primer + z, "yyyymmdd")
                 dfentra = primer + z
                 dfsale = primer + z
                 p0_ing = "": p0_sal = "": p0_ing1 = "": p0_sal1 = ""   'HORARIO ORIGINAL
                 p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": c2dias = "N"
                 m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
                 nValorH = 0: nValorM = 0: mdscto = 0: nFaltaH = 0: nFaltaM = 0:  F_falta = False: nmotpersonal = 0: nComision = 0
                 cvalor = "": cAuPrg = "": cDiahora = "": Ctipo = "": choraini = "": chorafin = "": cDescuenta = "": c_a = ""
                 cDiahora = "X"
                 cMantGuardia = "N"
                 
                 cAplica = "N": nSonLacta = 0
                 TxIniLac = "": TxFinLac = ""
                 If csexo = "F" Then
                    '-Ubica lactancia al periodo
                    Set rs = New Recordset
                    sql = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "' order by Fecnac"
                    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                    Do While Not rs.EOF
                        dftope = DateAdd("YYYY", 1, rs!fecnac)   'Tope Lactancia
                        'MsgBox Format(dftope, "dd/mm/yyyy")
                        If fechaB >= rs!fecnac And fechaB <= Format(dftope, "yyyymmdd") Then
                            nSonLacta = nSonLacta + 1
                            '-Carga Parametros de Ajuste horario por Lactancia
                            TxIniLac = rs!fecnac
                            TxFinLac = rs!FecFIN
                            cAplica = rs!Aplica
                        End If
                        rs.MoveNext
                    Loop
                    rs.Close
                 End If
                 
                 '-Tolerancia x LACTANCIA (I)ngreso/(S)alida
                 '=======================
                 nLacta = 0
                 If cAplica <> "N" Then
                    nLacta = IIf(nSonLacta = 1, 1, 2)
                 End If
                 '' MsgBox "Aplica:" + cAplica + " [" + Trim(Str(nLacta)) + "] Dia:" + Format(primer + z, "dd/mm/yyyy") + " en " + Format(TxIniLac, "dd/mm/yyyy") + " - " + Format(TxFinLac, "dd/mm/yyyy")
       
                 '-AUTORIZACIONES: Motivos Personales, Comisiones Serv   - excluye LVR
                 sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta='" & cTarjeta & "' and FechaIni<='" + fechaB + "' and FechaFin>='" + fechaB + "' and turno=''"
                 Set rs = New Recordset
                 rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                 Do While Not rs.EOF
                    If Not rs.EOF Then
                        cAuPrg = Trim(rs!Codigo)
                        cDiahora = rs!diahora
                        Ctipo = IIf(IsNull(rs!tipo), "", rs!tipo)            'I/S/O   caso horas
                        choraini = IIf(IsNull(rs!horaini), "", rs!horaini)
                        chorafin = IIf(IsNull(rs!horafin), "", rs!horafin)
                    End If
                    If cAuPrg <> "" Then
                        sql = "SELECT * FROM CpOcurrencias WHERE codigo='" & cAuPrg & "'"
                        Set rs_2 = New Recordset
                        rs_2.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                        If Not rs_2.EOF Then
                            cDescuenta = IIf(rs_2!DESCUENTO = "", "N", Trim(rs_2!DESCUENTO))
                            cMantGuardia = IIf(IsNull(rs_2!MantGuardia) Or Trim(rs_2!MantGuardia) = "", "N", rs_2!MantGuardia)
                        End If
                        rs_2.Close
                        '-Aplica Autorización Entrada(I)/(S)alida
                        If Ctipo = "I" Then
                            m_ing = choraini
                        ElseIf Ctipo = "S" Then
                                m_sal = chorafin
                        End If
                    End If
                    rs.MoveNext
                 Loop
                 If (cDiahora = "H" Or cDiahora = "D") And cMantGuardia = "N" Then
                    'Pierde Guardia
                    F_falta = True
                 Else
                    '-HORARIO prog en Dia '-**
                    sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
                    Set rs = New Recordset
                    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                    If rs.EOF Then
                       If Trim(cGrpGen) <> "" Then
                          rs.Close
                          sql = "SELECT * FROM CpHor" & genaño & " WHERE grupo='" & cGrpGen & "' AND fecha ='" & fechaB & "'"
                          Set rs = New Recordset
                          rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                       End If
                    End If
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
                             p_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
                             If rs_d!Horas < "1800" Then                               '========>
                                p_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                                p_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
                             End If
                             p_sal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
                             
                             '-horario ORIGINAL
                             '=================
                             p0_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
                             ''If rs_d!Horas <> "2400" And rs_d!Horas <> "1800" Then                               '========>
                                p0_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                                p0_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
                             ''End If
                             p0_sal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
                             '-Ajuste horario x LACTANCIA
                             '===========================
                             If cAplica <> "N" And nLacta >= 1 Then
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
                             ''MsgBox (cAplica + " => " + p0_ing + "-" + p_ing + " " + p0_sal + "-" + p_sal + " " + p0_sal1 + "-" + p_sal1)
             
                             p_horas = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
                             nValorH = Val(Left(p_horas, 2))
                             nValorM = Val(Right(p_horas, 2))
                             nFaltaH = Val(Left(p_horas, 2))
                             nFaltaM = Val(Right(p_horas, 2))
                             If (p0_ing > p0_sal And Val(p0_ing1) = 0) Then         '======>Or p_horas = "2400"
                                c2dias = "S"
                             ElseIf Val(p0_sal1) > 0 And p0_ing < p0_ing1 Then
                                    c2dias = "S"
                             End If
                          End If
                          rs_d.Close
                       End If
                    End If
                    rs.Close
                    
                    If Val(p_horas) > 0 Then
                       '-Marcas de Turno Programado
                       Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
                       '-Marcas del dia, solo si Justifica "D"
                       If cDiahora <> "D" Then
                          If Val(p_sal1) > 0 Or Val(p_sal) < Val(p_ing) Then
                             '-Ubica marcas 2dias Actual-siguiente caso: turno t/n, n/d
                             fecha1 = Format(primer + z + 1, "yyyymmdd")
                             Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha >='" & fechaB & "' AND fecha <='" & fecha1 & "' ORDER BY fecha,hora"
                             cdiaant = "S"
                          End If
                       End If
                       Set rs_d = New Recordset
                       rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                       If Not rs_d.EOF Then
                          '-ASIGNA marcas basado en Horario
                          d_fentra = rs_d!fecha      '========> 1era marca del servidor para turno programado
                          Do While Not rs_d.EOF
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
                                
                             ElseIf Val(m_ing) > 0 And (Val(rs_d!hora) - Val(m_ing) < 60) And c2dias = "S" And i = Day(dfsale) Then
                                   'Doble Ingreso¿2dias?     ============>
                                   m_ing1 = rs_d!hora
                             ElseIf i = Day(dfsale) And Val(m_ing) > 0 And (Val(rs_d!hora) - Val(m_ing) < 60) And ((Val(rs_d!hora) < Val(p_sal) And c2dias = "N")) Then
                                    'Doble Ingreso¿?
                                    m_ing1 = rs_d!hora
                             ElseIf i = Day(dfsale) And Val(p_ing1) > 0 And Val(m_ing1) = 0 And (Val(rs_d!hora) - Val(p_ing1) < 10) Then
                                     '-ing Turno partido
                                     m_ing1 = rs_d!hora
                             ElseIf i = Day(dfsale) And Val(p_ing1) = 0 And Val(m_ing1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then
                                    m_ing1 = rs_d!hora
                   
                             ElseIf (Val(m_sal) = 0 And (Val(rs_d!hora) - Val(p_sal) >= 0)) Or (Val(m_sal) = 0 And Val(p_sal) > Val(rs_d!hora)) Then
                                    '07001900-19000700   0701 0738x 0740
                                    m_sal = rs_d!hora
                             ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) > 0 And Val(m_sal1) = 0 And (Val(rs_d!hora) - Val(p0_sal1) >= 0) Then
                                    '-sal Turno partido
                                    m_sal1 = rs_d!hora
                                    
                             '-Check carga marcaciones sin valor para turno
                             'ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 And Val(m_ing1) > 0 Then
                             ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then     'And Val(m_ing1) > 0
                                    m_sal1 = rs_d!hora
                                    
                             End If
                             rs_d.MoveNext
                          Loop
                       End If
                       rs_d.Close
          
                    End If
                    '-Evalua MARCAS Ingreso
                    If Val(p_horas) > 0 Then
                        If (cDiahora = "D" Or cDiahora = "H") And cMantGuardia = "N" Then
                           F_falta = True
                           
                        'ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) Then      '======>
                         ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And Val(p_ing1) > 0 And p_sal <> p_ing1 Then  '======>
                               'Caso Horario:2000-0800   marca 0805-2010   O marca 0749-1404 1930-0820 ó  0803-2015/-0824
                               F_falta = True
                           
                        ElseIf (Val(m_ing) + Val(m_ing1) = 0) Then
                           '-No marco Ing,   valida vs autoriza
                           If cDiahora = "H" And Ctipo = "I" Then
                              cInicio = p_ing
                              cFin = choraini
                              If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                                 F_falta = True
                                 cDescuenta = "S"
                              End If
                           ElseIf cDiahora = "D" Then   'valida autoriza Dia Comp/ pierde
                              If cDescuenta = "S" Then
                                 F_falta = True
                              End If
                           Else
                              F_falta = True
                           End If
                        ElseIf (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                           'Marco tarde  Autoriza H al ingreso
                           If cDiahora = "H" And Ctipo = "I" Then
                              cInicio = p_ing
                              cFin = choraini
                              If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                                 F_falta = True
                              End If
                           End If
                           
                        ElseIf Val(m_ing) > 0 Then
                               '--- INSERTADO Excepciones ----09.08.17
                                If Val(m_ing1) = 0 Then
                                   '  PI   PS            MI   MS     MI1  MS1
                                   '- 2000-0800      :   0740-1401   FALTA
                                   '                 :   2130-0835   FALTA no marco salida  ¿?
                                   '-                :   1750-0820   ok
                                   '--               :   0741-2044             marcas fuera Turno (antes/despues)-> FALTA
                                    '---  0700-1900 1900-0700  24Hras  :marcas   0715- / -0748
                                   If (m_sal > m_ing And m_sal < p_ing) And p_horas <> 2400 Then    '¿2400?
                                      '-trabajo Antes de Turno
                                      F_falta = True
                                   Else
                                      If c2dias = "S" And ((m_sal > p_sal And d_fentra = dfsale) Or (m_sal < p_sal And d_fentra < dfsale)) Then
                                         '-No completo el turno
                                         F_falta = True
                                      End If
                                   End If
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
                                      End If
                                   End If
                                End If
                                                        
                        End If
                        '-Evalua MARCAS Salida
                        If F_ALTA Then
                        ElseIf (Val(m_sal) + Val(m_sal1) = 0) Then
                           '-No marco  Sal   valida vs autoriza
                           If cDiahora = "H" And Ctipo = "S" Then
                              If (Val(p_sal) - Val(chorafin) > 0) Or (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                                 F_falta = True
                                 cDescuenta = "S"
                              End If
                           ElseIf cDiahora = "D" Then     'valida autoriza Dia Comp/ pierde
                              If cDescuenta = "S" Then
                                 F_falta = True
                              End If
                           Else
                              F_falta = True
                           End If
                           
                        ElseIf Val(p_sal1) > 0 And p0_sal <> p0_ing1 Then    'p_horas<>"2400"
                                If Val(m_sal1) = 0 Then
                                    F_falta = True
                                Else
                                    If Val(p_sal1) > Val(m_sal1) Then
                                        F_falta = True
                                    End If
                                End If
                        
                        ElseIf Val(m_sal) > 0 And Val(p0_sal) > Val(m_sal) Then      'p_sal
                           'Marco antes de, Autoriza H
                            If cDiahora = "H" And Ctipo = "S" Then    '====> And cTipo = "S"  ** TO CHECK **
                               cInicio = p_sal
                               cFin = chorafin
                               If Val(p_sal) - Val(chorafin) > 0 Then
                                  F_falta = True
                               End If
                            ElseIf cDiahora = "X" Then
                               cInicio = p_sal
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
                                  
                               ElseIf Val(p_sal1) - Val(cFin) > 0 Then
                                      F_falta = True
                               End If
                            End If
                        End If
                        
                        If F_falta Then
                        Else
                           If cDiahora = "H" And Ctipo = "I" Then
                           ElseIf cDiahora <> "D" Then      '<===No Autorizaciones
                              If (Val(m_ing) + Val(m_ing1) > 0) Then
                                 '-Hay marcas, aplica tolerancia, calcula tardanza
                                 If cDiahora = "H" And Ctipo = "I" Then
                                    cInicio = p_ing
                                    cFin = choraini
                                    nmintarde = CalcMints(cInicio, cFin)
                                 Else
                                    nmintarde = CalcMTard(p_ing, m_ing)
                                 End If
                                 If nmintarde >= 6 And nmintarde <= 10 Then
                                    mdscto = 10
                                 End If
                                 If nmintarde >= 11 And nmintarde <= 20 Then
                                    mdscto = 20
                                 End If
                                 If nmintarde >= 21 And nmintarde < 31 Then
                                    mdscto = 30
                                 End If
                                 If nmintarde > 30 Then
                                    F_falta = True
                                 End If
                              End If
                           End If
                        End If
                    End If
                 End If
                    
                 If F_falta Then
                 Else
                    If Val(p_horas) > 0 Then
                       If cDiahora = "D" And (Left(cAuPrg, 2) = "VA" Or cAuPrg = "COMP") Then
                       Else
                       
                          If cAuPrg = "" Or (cAuPrg <> "" And cMantGuardia = "S") Then
                            cPatron = Mid(ccadena, i, 1)     ' F/O
                            If cPatron = "F" Then
                               If InStr("D", Left(Trim(cTurno), 1)) > 0 Then
                                  ndf = ndf + 1
                               ElseIf InStr("N_H_C", Left(Trim(cTurno), 1)) > 0 Then
                                      nnf = nnf + 1
                               ElseIf Left(cTurno, 1) = "P" Then
                                      ndf = ndf + 1
                                      nnf = nnf + 1
                               End If
                            ElseIf cPatron = "O" Then
                                   If InStr("D", Left(Trim(cTurno), 1)) > 0 Then
                                      ndo = ndo + 1
                                   ElseIf InStr("N_H_C", Left(Trim(cTurno), 1)) > 0 Then
                                          nno = nno + 1
                                   ElseIf Left(cTurno, 1) = "P" Then
                                          ndo = ndo + 1
                                          nno = nno + 1
                                   End If
                            End If
                          End If       'c/Autoriza y MantGuardia
                          
                       End If
                    End If
                 End If
              
              End If        '<-- evaluo turno
           Next i
           nimporte = 0#
           If ndo + nno + ndf + nnf > 0 Then
              nimporte = ndo * nTarDO + nno * nTarNO + ndf * nTarDF + nnf * nTarNF
              SQ1 = "SELECT  * FROM CpDirectorio  WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes) & "' and tarjeta='" & cTarjeta & "' and flag='0' "
              Set rs_d = New Recordset
              rs_d.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
              ndo0 = IIf(IsNull(rs_d!D_O), 0, rs_d!D_O)
              nno0 = IIf(IsNull(rs_d!N_O), 0, rs_d!N_O)
              ndf0 = IIf(IsNull(rs_d!D_F), 0, rs_d!D_F)
              nnf0 = IIf(IsNull(rs_d!N_F), 0, rs_d!N_F)
              nImporte0 = rs_d!importe
              rs_d.Close
              If ndo <> ndo0 Or nno <> nno0 Or ndf <> ndf0 Or nnf <> nnf0 Then
                 If nimporte > nImporte0 Then
                    rt.AddNew
                     rt!CCosto = c_Costo
                     rt!Dsccosto = cDscCosto
                     rt!CodUpr = cCodUpr
                     rt!dscUPR = cDscUpr
                    rt!CadHep = ccadhep
                    rt!Dsccadhep = cDcadhep
                    rt!tarjeta = cTarjeta
                    rt!apelpat = caPat
                    rt!apelmat = caMat
                    rt!nombres = cnomb
                    rt!Dni = cDni
                    rt!NivGuardia = cnivGuardia
                    rt!DsccarPla = cDcarpla
                    rt!TarifaDO = nTarDO
                    rt!D_O = ndo
                    rt!TarifaNO = nTarNO
                    rt!N_O = nno
                    rt!TarifaDF = nTarDF
                    rt!D_F = ndf
                    rt!TarifaNF = nTarNF
                    rt!N_F = nnf
                    rt!Monto = nimporte
                    rt!Monto0 = nImporte0
                    rt!NroReso0 = cNroReso0
                    rt!ABONO = nimporte - nImporte0
                    rt.Update
                    nvalor = nvalor + 1
                 End If
              End If
           End If
        End If
        r_d.MoveNext
     Loop
   End If
   r_d.Close
   rt.Close
   cmdcarga.Caption = ""
   If nvalor > 0 Then
     MsgBox "Total Servidores " + Trim(Str(nvalor)), vbExclamation
     TxNroReso0.Text = cNroReso0
     TxFecReso0.Text = fecReso0
     
     c_a = Right("0" + Trim(Str(Month(TxFecReso0.Text))), 2)
     Set rs = New Recordset
     rs.Open "SELECT * FROM cptablas WHERE tabla='32' and codigo ='" & c_a & "'", Cnx, adOpenStatic, adLockReadOnly
     cPatron = Trim(rs!descripcion)
     rs.Close
     TxFechaR0.Text = Trim(Str(Day(TxFecReso0.Text))) + " de " + cPatron + " del " & Trim(Str(Year(TxFecReso0.Text)))
   End If
   CmdPrint.SetFocus
   Exit Sub
   'Call Limpiacampos
End Sub

Private Sub TxNroResolu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     TxNroResolu.Text = UCase(TxNroResolu.Text)
     Exit Sub
   End If
End Sub

Private Sub CmdExportaXLS_Click()
    If ChkXLS.Value = 0 Then
       wTpoQuery = "4"
       Call ExportTableToXLS_click
    Else
       wTpoQuery = "4.0"
       wNameXLS = "C:\RPT\ResPGAbo.xls"
       Call Exportar_ADO_Excel
    End If
    Beep 500, 350
    MsgBox "Se generó C:\RPT\ResPGAbo.xls", vbInformation
    Exit Sub
End Sub

Private Sub CmdPrint_Click()
   If Txmes.Text = "" Then
      MsgBox "Defina..  Periodo... ", vbExclamation
      Exit Sub
   End If
   If nvalor = 0 Then
      MsgBox "No existe registros... ", vbExclamation
      Exit Sub
   End If
   If Trim(TxNroResolu) = "" Then
      MsgBox "Indique Nº Resolución... ", vbExclamation
      Exit Sub
   End If
   
   TxNroResolu = UCase(TxNroResolu)
   nTipRpt = 16
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Call Limpiacampos
   CmdPrint.Caption = ""
   Exit Sub
End Sub
Private Sub Limpiacampos()
   CboMes.Text = "":   Txmes = "": TxNombMes = "": TxNroResolu = "": TxNroReso0 = "": TxFecReso0 = "": TxFechaR0 = ""
   Txfecini = "": TxFecFin = "": TxIniLac = "": TxFinLac = ""
   ChkXLS.Value = 1
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
        If UCase(Forms(cont).Name) <> "CPMENU" And UCase(Forms(cont).Name) <> "RESPGOGUARABON" Then
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
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "ResPgoGuarAbon" Then
           Unload Forms(cont)
        End If
      End If
   Next
   
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


