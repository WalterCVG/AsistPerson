VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ResMP 
   Caption         =   "Resolución -  MOTIVOS PERSONALES"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "ResMP.frx":0000
   ScaleHeight     =   3300
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   9000
      Picture         =   "ResMP.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cerrar"
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox ChkXLS 
      Alignment       =   1  'Right Justify
      Caption         =   "XLS instalado"
      Height          =   195
      Left            =   6960
      TabIndex        =   24
      Top             =   1680
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
      Left            =   3000
      TabIndex        =   23
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
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
      Left            =   2160
      TabIndex        =   22
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxRegPlh 
      Enabled         =   0   'False
      Height          =   225
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboRegPlh 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox TxNroResolu 
      Height          =   285
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8880
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Txmesano 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdExportaXLS 
      Height          =   615
      Left            =   7200
      Picture         =   "ResMP.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exporta a Xls"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox TxMesActual 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TxMesdeAno 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
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
      Left            =   1440
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox TxMES 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      MaxLength       =   2
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "ResMP.frx":80FF2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton CmdClear 
      Height          =   855
      Left            =   8880
      Picture         =   "ResMP.frx":81C34
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Otra Consulta"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   975
      Left            =   6120
      Picture         =   "ResMP.frx":824FE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ubicar Registros"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Txfecini 
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
      Left            =   2160
      TabIndex        =   13
      Top             =   1800
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
      Left            =   3000
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   4440
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
      Begin VB.CommandButton mnuFicImpSelec 
         Height          =   675
         Left            =   120
         Picture         =   "ResMP.frx":826D9
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selec.Impresora"
         Top             =   240
         Width           =   765
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
         Height          =   675
         Left            =   960
         Picture         =   "ResMP.frx":82EB7
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir"
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg. Laboral"
      Height          =   255
      Left            =   1920
      TabIndex        =   21
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. Resolucion"
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "ResMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim rt1 As New ADODB.Recordset
Dim Sw, f_Adm, F_add As Boolean
Dim nLacta, nSonLacta, nposicion, nson, nsondias As Integer
Dim csexo, cAplica, cDni, sql, SQ1, cTarjetaOLD, cTarjeta, cano, cmes As String
Dim unafecha As Date
Dim nRemTot, nRemHor, ntTardanza, nTDscto, nTMotPer, nMontoD, nMontoH, nMontoPer As Double
Private Sub Command1_Click()
     c_grupo = "02": c_secuencia = "041"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub
Private Sub Form_Load()
  ResMP.Caption = "Resolución - MOTIVOS PERSONALES " & genaño
  TxMesActual = Trim(Mid(wMeses, (Month(Date) - 1) * 9 + 1, 9))

  TxFecActual = "La Victoria, " + Right$("00" + Trim(Str(Day(Date))), 2) + " de " + TxMesActual + " de " & Year(Date)
  ChkXLS.Value = 1
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
  unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
  Call ObtenerPrimerUltimoDia(unafecha)
  Txfecini = primer
  TxFecFin = ultimo
  '-Regimen Laboral
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
  rs.Close
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
         TxMesdeAno.Text = TxNombMes.Text + " de " & genaño
         TxMesAno.Text = TxNombMes.Text + " " & genaño
         unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
         Call ObtenerPrimerUltimoDia(unafecha)    'primer/ultimo
         Txfecini.Text = primer
         TxFecFin.Text = ultimo
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
             TxMesdeAno.Text = TxNombMes.Text + " de " & genaño
             TxMesAno.Text = TxNombMes.Text + " " & genaño
            unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)            'primer/ultimo
            Txfecini.Text = primer
            TxFecFin.Text = ultimo
         End If
      End If
      Exit Sub
   End If
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
     nposicion = InStr(CboRegPLH.List(j), ":")
     nlen = Len(CboRegPLH.List(j))
     TxRegPLH.Text = Right(CboRegPLH.List(j), nlen - nposicion)
  End If
End Sub
Private Sub cmdcarga_Click()
   If IsDate(Txfecini.Text) And IsDate(TxFecFin.Text) Then
      TxMesdeAno.Text = TxNombMes.Text + " de " & genaño
      TxMesAno.Text = TxNombMes.Text + " " & genaño
   Else
      Exit Sub
   End If
   
   '-Limpia TEMPORALES   Detalle/Cabecera
   Cnx.Execute "DELETE  FROM CpTmResMP"
   sql = "SELECT * FROM CpTmResMP"
   Set rt = New Recordset
   rt.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   dfi = Day(Txfecini.Text)
   dff = Day(TxFecFin.Text)
   cmes = Right("00" + Trim(Str(Month(Txfecini.Text))), 2)
   unafecha = Txfecini.Text
   Call ObtenerPrimerUltimoDia(unafecha)
   
   fecha1 = Format(primer, "yyyymmdd")
   fecha2 = Format(ultimo, "yyyymmdd")
   
   sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and  fechaini>='" & fecha1 & "' AND fechafin<='" & fecha2 & "' And Codigo='MP'  ORDER BY tarjeta,fechaini"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   nvan = 0
   nson = 0
   Do While Not rs.EOF
      nvan = nvan + 1
      f_Adm = False
      cApelP = "": cApelM = "":    cnomb = "": cDni = "": cGrpGen = ""
      cplh = "plh":  c_Costo = "": cDscCosto = "": cCodUpr = "": cDscUpr = "": ccadhep = "": cDsccadhep = "": ccarpla = "": cDsccarpla = "": cnivGuardia = ""
      nRemTot = 0: nRemHor = 0
      cTarjeta = rs!tarjeta     '-*
      sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & cTarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs_d.EOF Then
         '-SABER si una columna existe en una tabla
         SQ1 = "SELECT * FROM  INFORMATION_SCHEMA.COLUMNS AS c1  WHERE  c1.column_name = 'cadhep'  AND c1.table_name = 'cpdirectorio'"
         Set rs_2 = New Recordset
         rs_2.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If Not rs_2.EOF Then
            ccadhep = IIf(IsNull(rs_d!CadHep), "", Trim(rs_d!CadHep))
         End If
         rs_2.Close
      End If
      rs_d.Close
      cmdcarga.Caption = cTarjeta + ":" + Str(nvan)
      cDni = ""
      SQ1 = "SELECT * FROM Cpgeneral WHERE (baja is NULL) and  tarjeta='" & Trim(cTarjeta) & "'"
      Set rs_d = New Recordset
      rs_d.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs_d.EOF Then
         cApelP = IIf(IsNull(rs_d!apelpat), "", rs_d!apelpat)
         cApelM = IIf(IsNull(rs_d!apelmat), "", rs_d!apelmat)
         cnomb = IIf(IsNull(rs_d!nombres), "", rs_d!nombres)
         cDni = IIf(IsNull(rs_d!documento), "", rs_d!documento)
         csexo = IIf(IsNull(rs_d!sexo), "", rs_d!sexo)
         cGrpGen = "" + Trim(rs_d!grpgen)
         c_Costo = IIf(IsNull(rs_d!CCosto), "", rs_d!CCosto)
         cCodUpr = IIf(IsNull(rs_d!CodUpr), "", rs_d!CodUpr)
         If ccadhep = "" Then
            ccadhep = IIf(IsNull(rs_d!CadHep), "", rs_d!CadHep)
         End If
         ccarpla = IIf(IsNull(rs_d!carPla), "", rs_d!carPla)
         cnivGuardia = IIf(IsNull(rs_d!NivGuardia), "", rs_d!NivGuardia)
         nRemTot = IIf(IsNull(rs_d!RemTot), 0, rs_d!RemTot)
         '-PLH del servidor
         cplh = IIf(IsNull(rs_d!PLH), "plh", rs_d!PLH)
         f_Adm = IIf(rs_d!tipoPla = "ASI", False, True)
      End If
      rs_d.Close
            
      '- Afecto X LACTANCIA
      '======================
      cAplica = "N": nSonLacta = 0
      TxIniLac = "": TxFinLac = ""
            
      ntdiasfalt = 0: ntmontoH = 0: ntmontoD = 0: ntTardanza = 0: nTMotPer = 0
      Do While rs!tarjeta = cTarjeta
        nmonto = 0:  nvalor = 0:  mdscto = 0: nmotpersonal = 0
        F_falta = False
        nValorH = 0: nValorM = 0:  nFaltaH = 0: nFaltaM = 0
        cvalor = "": cAuPrg = "": cDiahora = "": Ctipo = "": choraini = "": chorafin = "": cDescuenta = ""
        cMantGuardia = ""
        '' MsgBox "plh_serv-" + capelp + " " + cnomb + ":" + cplh + "[" + TxPlh + "]"
        If (TxRegPLH <> "" And TxRegPLH = cplh) Or TxRegPLH = "" Then    '.....A
           cAuPrg = Trim(rs!Codigo)
           cDiahora = rs!diahora             'D/H
           Ctipo = IIf(IsNull(rs!tipo), "", rs!tipo)            'I/S/O   caso horas
           choraini = IIf(IsNull(rs!horaini), "", rs!horaini)
           chorafin = IIf(IsNull(rs!horafin), "", rs!horafin)
           DfechaI = rs!fechaIni
           DfechaF = rs!fechaFin
           SQ1 = "SELECT * FROM CpOcurrencias WHERE codigo='" & cAuPrg & "'"
           Set rs_d = New Recordset
           rs_d.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
           If Not rs_d.EOF Then
              cDescuenta = IIf(rs_d!DESCUENTO = "", "N", Trim(rs_d!DESCUENTO))
              cMantGuardia = IIf(IsNull(rs_d!MantGuardia), "", rs_d!MantGuardia)
           End If
           rs_d.Close
           If f_Adm Then
              nRemHor = Round(nRemTot / (30 * 7.25), 2)
           Else
              nRemHor = Round(nRemTot / (30 * 6), 2)
           End If
           '-Evalua Registro
           dfi = Day(DfechaI)
           dff = Day(DfechaF)
           If cDiahora = "H" Then
              cInicio = choraini
              cFin = chorafin
              nmotpersonal = CalcMints(cInicio, cFin)
              If nmotpersonal Mod 60 > 30 Then
                 mdscto = Int(nmotpersonal / 60) + 1
              Else
                 mdscto = Int(nmotpersonal / 60)
              End If
              '-Verifica  DIRECTORIO "TURNO"
              sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & cTarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
              Set rs_d = New Recordset
              rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
              If Not rs_d.EOF Then
                 cdiaant = "N"
                 For i = dfi To dff
                   z = i - 1
                   p0_ing = "": p0_sal = "": p0_ing1 = "": p0_sal1 = ""   'HORARIO ORIGINAL
                   cTurno = "": p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": c2dias = "N"
                   m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
                   cvalor = "": nValorH = 0: nValorM = 0: nmintarde = 0: ndscto = 0: F_falta = False
                   
                   cTurno = IIf(IsNull(rs_d.Fields(z + 4)), "", rs_d.Fields(z + 4))
                   fechaB = Format(primer + z, "yyyymmdd")
                   Fechant = Format(primer + z, "yyyymmdd")
                   dfentra = primer + z
                   dfsale = primer + z
                   
                   cAplica = "N": nSonLacta = 0
                   TxIniLac = "": TxFinLac = ""
                   If csexo = "F" Then
                      '-Ubica lactancia al periodo
                      Set rs_1 = New Recordset
                      SQ1 = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "'  order by fecnac"
                      rs_1.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                      Do While Not rs_1.EOF
                        dftope = DateAdd("YYYY", 1, rs_1!fecnac)   'Tope Lactancia
                        'MsgBox Format(dftope, "dd/mm/yyyy")
                        If fechaB >= rs_1!fecnac And fechaB <= Format(dftope, "yyyymmdd") Then
                            nSonLacta = nSonLacta + 1
                            '-Carga Parametros de Ajuste horario por Lactancia
                            TxIniLac = rs_1!fecnac
                            TxFinLac = rs_1!FecFIN
                            cAplica = rs_1!Aplica
                        End If
                        rs_1.MoveNext
                      Loop
                      rs_1.Close
                   End If
                   '-Tolerancia x LACTANCIA (I)ngreso/(S)alida
                   '=======================
                   nLacta = 0
                   If cAplica <> "N" Then
                      nLacta = IIf(nSonLacta = 1, 1, 2)
                   End If
                   '' MsgBox "Aplica:" + cAplica + " [" + Trim(Str(nLacta)) + "] Dia:" + Format(primer + z, "dd/mm/yyyy") + " en " + Format(TxIniLac, "dd/mm/yyyy") + " - " + Format(TxFinLac, "dd/mm/yyyy")
       
                   '-HORARIO Prog en dia -**
                   SQ1 = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
                   Set rs_1 = New Recordset
                   rs_1.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                   If rs_1.EOF Then
                      rs_1.Close
                      Fechant = Format(primer + z - 1, "yyyymmdd")
                      SQ1 = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "' and fecha ='" & Fechant & " '"
                      Set rs_1 = New Recordset
                      rs_1.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                   End If
                   If Not rs_1.EOF Then
                      cvalor = IIf(IsNull(rs_1!Horario), "", rs_1!Horario)
                   End If
                   rs_1.Close
                   '-Horario de-a
                   SQ1 = "SELECT * FROM CpHorarioNew  WHERE Codigo='" & cvalor & "'"
                   Set rs_1 = New Recordset
                   rs_1.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                   If Not rs_1.EOF Then
                      p_ing = IIf(IsNull(rs_1!ingreso), "", rs_1!ingreso)
                      If rs_d!Horas < "1800" Then                               '========>
                         p_sal = IIf(IsNull(rs_1!salida), "", rs_1!salida)
                         p_ing1 = IIf(IsNull(rs_1!Ingreso1), "", rs_1!Ingreso1)
                      End If
                      p_sal1 = IIf(IsNull(rs_1!salida1), "", rs_1!salida1)
                                            
                      '-horario ORIGINAL
                      '=================
                      p0_ing = IIf(IsNull(rs_1!ingreso), "", rs_1!ingreso)
                      ''If rs_d!Horas <> "2400" And rs_d!Horas <> "1800" Then                               '========>
                         p0_sal = IIf(IsNull(rs_1!salida), "", rs_1!salida)
                         p0_ing1 = IIf(IsNull(rs_1!Ingreso1), "", rs_1!Ingreso1)
                      ''End If
                      p0_sal1 = IIf(IsNull(rs_1!salida1), "", rs_1!salida1)
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
                          
                      p_horas = IIf(IsNull(rs_1!Horas), "0000", rs_1!Horas)
                      nValorH = Val(Left(p_horas, 2))
                      nValorM = Val(Right(p_horas, 2))
                      
                      If (p0_ing > p0_sal And Val(p0_ing1) = 0) Then          '======>Or p_horas = "2400"
                         c2dias = "S"
                      ElseIf Val(p0_sal1) > 0 And p0_ing < p0_ing1 Then
                             c2dias = "S"
                      End If
                      
                   End If
                   rs_1.Close
                   '-MARCAS del dia       fecha  hora =>  hhmm
                   If Val(p_horas) > 0 Then
                      '-Marcas en DIA Programado
                      If fechaB = Fechant Then
                         SQ1 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
                         'If Val(p_sal1) > 0 Or Val(p_sal) < Val(p_ing) Then
                      Else
                         '-Ubica marcas 2dias Actual-siguiente caso: turno t/n, n/d
                         'Fecha1 = Format(primer + z + 1, "yyyymmdd")
                         SQ1 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha >='" & Fechant & "' AND fecha <='" & fechaB & "' ORDER BY fecha,hora"
                         cdiaant = "S"
                      End If
                 
                      Set rs_1 = New Recordset
                      rs_1.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                      If Not rs_1.EOF Then
                         '-ASIGNA marcas basado en Horario
                         d_fentra = rs_d!fecha       '========> 1era marca del servidor para turno programado
                         Do While Not rs_1.EOF
                            cInicio = p_ing
                            cFin = rs_1!hora
                            dfsale = rs_1!fecha
                            If Val(m_ing) = 0 Then
                               'If (nTotMints(cInicio, cFin) >= -60 And nTotMints(cInicio, cFin) <= 60) And Day(dfsale) = i Then
                               If Day(dfsale) = i Then
                                  m_ing = rs_1!hora
                               Else
                                   m_sal = rs_1!hora
                               End If
                               
                            ElseIf Val(m_ing) > 0 And (Val(rs_1!hora) - Val(m_ing) < 60) And c2dias = "S" And i = Day(dfsale) Then
                                   'Doble Ingreso¿2dias?     ============>
                                   m_ing1 = rs_1!hora
                            ElseIf i = Day(dfsale) And Val(m_ing) > 0 And (Val(rs_1!hora) - Val(m_ing) < 60) And ((Val(rs_1!hora) < Val(p_sal) And c2dias = "N")) Then
                                   'Doble Ingreso¿?
                                   m_ing1 = rs_1!hora
                            ElseIf i = Day(dfsale) And Val(p_ing1) > 0 And Val(m_ing1) = 0 And (Val(rs_1!hora) - Val(p_ing1) < 10) Then
                                   '-ing Turno partido
                                   m_ing1 = rs_1!hora
                            ElseIf i = Day(dfsale) And Val(p_ing1) = 0 And Val(m_ing1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then
                                   m_ing1 = rs_1!hora
                          
                            ElseIf (Val(m_sal) = 0 And (Val(rs_1!hora) - Val(p_sal) >= 0)) Or (Val(m_sal) = 0 And Val(p_sal) > Val(rs_1!hora)) Then
                                   '07001900-19000700   0701 0738x 0740
                                   m_sal = rs_1!hora
                            ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) > 0 And Val(m_sal1) = 0 And (Val(rs_1!hora) - Val(p0_sal1) >= 0) Then
                                   '-sal Turno partido  1300-1900/1900-0700
                                   m_sal1 = rs_1!hora
                                   
                            '-Check carga marcaciones sin valor para turno
                            'ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 And Val(m_ing1) > 0 Then
                             ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then     'And Val(m_ing1) > 0
                                   m_sal1 = rs_1!hora
                            End If
                                   
                            rs_1.MoveNext
                         Loop
                      End If
                      rs_1.Close
          
                      '-Evalua MARCAS Ingreso
                      If cDiahora = "D" And cDescuenta = "S" Then
                         F_falta = True
                      'ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And p_sal <> p_ing1 Then      '======>
                       ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And Val(p_ing1) > 0 And p_sal <> p_ing1 Then  '======>
                            'Caso Horario:2000-0800   marca 0805-2010   O marca 0749-1404 1930-0820 ó  0803-2015/-0824
                            F_falta = True
                      
                      ElseIf (Val(m_ing) + Val(m_ing1) = 0) Then
                         '-NO marco Ing,   valida vs autoriza
                         If cDiahora = "H" And Ctipo = "I" Then
                            cInicio = p_ing
                            cFin = choraini
                            If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                               F_falta = True
                            End If
                         ElseIf cDiahora = "D" Then     'valida autoriza Dia Comp/ pierde
                            'If cDescuenta = "S" Then
                            '   F_falta = True
                            'End If
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
                            If (m_sal > m_ing And m_sal < p_ing) And p_horas <> 2400 Then       '¿2400?
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
                         '-NO marco  Sal   valida vs autoriza
                         If cDiahora = "H" And Ctipo = "S" Then
                            If (Val(p_sal) - Val(chorafin) > 0) Or (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                               F_falta = True
                            End If
                         ElseIf cDiahora = "D" Then     'valida autoriza Dia Comp/ pierde
                            'If cDescuenta = "S" Then
                            '   F_falta = True
                            'End If
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
                      
                      ElseIf Val(m_sal) > 0 And Val(p0_sal) > Val(m_sal) Then     'p_sal
                         'Marco antes de, Autoriza H
                         If cDiahora = "H" And Ctipo = "S" Then    '====> And cTipo = "S"  ** TO CHECK **
                            cInicio = p_sal
                            cFin = chorafin
                            If Val(p_sal) - Val(chorafin) > 0 Then
                               F_falta = True
                            End If
                         ElseIf cDiahora = "" Then
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
                      
                      If falta Then
                      Else
                         ntTardanza = ntTardanza + nmotpersonal   'mints
                         nTMotPer = nTMotPer + mdscto             'hras
                         ntmontoH = ntmontoH + (mdscto * nRemHor)
                      End If
                   End If
                 Next i
              End If
           Else
              '-Carga programa DIRECTORIO "TURNO"
              sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & cTarjeta & "' AND ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "'"
              Set rs_d = New Recordset
              rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
              If Not rs_d.EOF Then
                 For i = dfi To dff
                    z = i - 1
                    cTurno = "": p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = ""
                    m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
                    cvalor = "": nValorH = 0: nValorM = 0: nmintarde = 0: ndscto = 0: F_falta = False
                    cTurno = IIf(IsNull(rs_d.Fields(z + 4)), "", rs_d.Fields(z + 4))
                    fechaB = Format(primer + z, "yyyymmdd")
                    
                    cAplica = "N": nSonLacta = 0
                    TxIniLac = "": TxFinLac = ""
                    If csexo = "F" Then
                      '-Ubica lactancia al periodo
                      Set rs_1 = New Recordset
                      SQ1 = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "'  order by fecnac"
                      rs_1.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                      Do While Not rs_1.EOF
                        dftope = DateAdd("YYYY", 1, rs_1!fecnac)   'Tope Lactancia
                        'MsgBox Format(dftope, "dd/mm/yyyy")
                        If fechaB >= rs_1!fecnac And fechaB <= Format(dftope, "yyyymmdd") Then
                            nSonLacta = nSonLacta + 1
                            '-Carga Parametros de Ajuste horario por Lactancia
                            TxIniLac = rs_1!fecnac
                            TxFinLac = rs_1!FecFIN
                            cAplica = rs_1!Aplica
                        End If
                        rs_1.MoveNext
                      Loop
                      rs_1.Close
                    End If
                    '-Tolerancia x LACTANCIA (I)ngreso/(S)alida
                    '=======================
                    nLacta = 0
                    If cAplica <> "N" Then
                       nLacta = IIf(nSonLacta = 1, 1, 2)
                    End If
                    
                    
                    '-HORARIO Prog en dia
                    SQ1 = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
                    Set rs_1 = New Recordset
                    rs_1.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                    If rs_1.EOF Then
                       If Trim(cGrpGen) <> "" Then
                          rs_1.Close
                          SQ1 = "SELECT * FROM CpHor" & genaño & " WHERE grupo='" & cGrpGen & "' AND fecha ='" & fechaB & "'"
                          Set rs_1 = New Recordset
                          rs_1.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                       End If
                    End If
                    If Not rs_1.EOF Then
                       cvalor = IIf(IsNull(rs_1!Horario), "", rs_1!Horario)
                    End If
                    rs_1.Close
                    SQ1 = "SELECT * FROM CpHorarioNew  WHERE Codigo='" & cvalor & "'"
                    Set rs_1 = New Recordset
                    rs_1.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                    If Not rs_1.EOF Then
                       p_ing = IIf(IsNull(rs_1!ingreso), "", rs_1!ingreso)
                       p_sal = IIf(IsNull(rs_1!salida), "", rs_1!salida)
                       p_ing1 = IIf(IsNull(rs_1!Ingreso1), "", rs_1!Ingreso1)
                       p_sal1 = IIf(IsNull(rs_1!salida1), "", rs_1!salida1)
                       '-horario ORIGINAL
                       '=================
                       p0_ing = IIf(IsNull(rs_1!ingreso), "", rs_1!ingreso)
                       p0_sal = IIf(IsNull(rs_1!salida), "", rs_1!salida)
                       p0_ing1 = IIf(IsNull(rs_1!Ingreso1), "", rs_1!Ingreso1)
                       p0_sal1 = IIf(IsNull(rs_1!salida1), "", rs_1!salida1)
                       
                       '-Ajuste horario x LACTANCIA
                       '===========================
                       If cAplica <> "N" And nLacta >= 1 Then
                          If cAplica = "I" Then   'al Ingreso
                             If Val(p_ing) > 0 And Val(p_ing1) = 0 Then
                                phora = Right("00" + Trim(Str(Val(Left(p_ing, 2)) + 1)), 2)
                                pmints = Right(p_ing, 2)
                                p_ing = phora + pmints
                             End If
                          ElseIf cAplica = "S" Then   'a Salida
                                 If Val(p_sal) > 0 And Val(p_sal1) = 0 Then
                                    phora = Right("00" + Trim(Str(Val(Left(p_sal, 2)) - 1)), 2)
                                    pmints = Right(p_sal, 2)
                                    p_sal = phora + pmints
                                 ElseIf Val(p_sal1) > 0 Then
                                        phora = Right("00" + Trim(Str(Val(Left(p_sal1, 2)) - 1)), 2)
                                        pmints = Right(p_sal1, 2)
                                        p_sal1 = phora + pmints
                                 End If
                          End If
                       End If
                       ''MsgBox (cAplica + " => " + p0_ing + "-" + p_ing + " " + p0_sal + "-" + p_sal + " " + p0_sal1 + "-" + p_sal1)
                       
                       p_horas = IIf(IsNull(rs_1!Horas), "0000", rs_1!Horas)
                       nValorH = Val(Left(p_horas, 2))
                       nValorM = Val(Right(p_horas, 2))
                    End If
                    rs_1.Close
                    '-MARCAS del dia       fecha  hora =>  hhmm
                    If Val(p_horas) > 0 Then
                       '-Turnos default "M", "T", "A"
                       nsondias = 1
                       If InStr("N_D", Left(cTurno, 1)) > 0 Or InStr(cTurno, "MT") > 0 Or InStr(cTurno, "MX") > 0 Then
                          nsondias = nsondias + 1
                       ElseIf InStr("H_C", Left(cTurno, 1)) > 0 Then
                              nsondias = nsondias + 2
                       ElseIf InStr("P", Left(cTurno, 1)) > 0 Then
                              nsondias = nsondias + 3
                       End If
                       ntdiasfalt = ntdiasfalt + nsondias
                       If f_Adm Then
                          ntmontoD = ntmontoD + Round(nRemTot / 30, 2)
                       Else
                          ntmontoD = ntmontoD + Round((nsondias * nRemTot / 30), 2)
                       End If
                    End If
                 Next i
              End If
           End If
           'If F_falta Or (nmotpersonal > 1) Then
              'rt1.AddNew
              'rt1!Tarjeta = cTarjeta
              'rt1!dscactividad = cDActividad
              'rt1!nombdia = Left(WeekdayName(Weekday(primer + z)), 3)
              'rt1!fecha = primer + z
              'rt1!turno = c_turno
              'rt1!horario = cvalor
              '- Horario ORIGINAL
              'rt1!ping = "" + p0_ing
              'rt1!psal = "" + p0_sal
              'rt1!ping1 = "" + p0_ing1
              'rt1!psal1 = "" + p0_sal1
              'rt1!hrasprog = 0 + (nvalorH + nvalorM / 60)  'p_horas
              'rt1!ming = "" + m_ing
              'rt1!msal = "" + m_sal
              'rt1!ming1 = "" + m_ing1
              'rt1!msal1 = "" + m_sal1
              'rt1!tardanza = 0 + IIf(F_falta, 0, IIf((Val(m_ing) - Val(p_ing)) < 0, 0, Val(m_ing) - Val(p_ing)))
              'rt1!descuento = 0 + IIf(F_falta, 0, mdscto)
              'rt1!autorizac = "" + cAuPrg
              'rt1!descuenta = "" + cDescuenta
              'rt1!diahora = "" + cDiahora
              'rt1!tipo = "" + cTipo
              'rt1!horaini = "" + choraini
              'rt1!horafin = "" + chorafin
              'rt1!falta = IIf(F_falta, "S", "N")
              'rt1!Comision = nComision
              'rt1!motpersonal = nmotpersonal
              'rt1!HrasTrab = IIf(F_falta, IIf(cDescuenta = "N", nvalor, 0), nvalor - nmotpersonal)
              'rt1!hrasfalt = IIf(F_falta, IIf(cDescuenta = "N", 0, nvalor), 0)
              'rt1.Update
           'End If
        End If            '......A
        rs.MoveNext
        If rs.EOF Then
           Exit Do
        End If
      Loop
      If nTMotPer + ntdiasfalt > 0 Then
         If c_Costo <> "" Then
            Set rs_d = New Recordset
            rs_d.Open "SELECT * FROM Cpccosto WHERE codigo='" & c_Costo & "'", Cnx, adOpenStatic, adLockReadOnly
            If Not rs_d.EOF Then
               cDscCosto = Trim(rs_d!descripcion)
            End If
            rs_d.Close
         End If
         If cCodUpr <> "" Then
            Set rs_d = New Recordset
            rs_d.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
            If Not rs_d.EOF Then
               cDscUpr = Trim(rs_d!descripcion)
            End If
            rs_d.Close
         End If
         If ccadhep <> "" Then
            Set rs_d = New Recordset
            rs_d.Open "SELECT * FROM Cpcadena WHERE ano='" & genaño & "' and cadena='" & ccadhep & "'", Cnx, adOpenStatic, adLockReadOnly
            If Not rs_d.EOF Then
               cDsccadhep = Trim(rs_d!descripcion)
            End If
            rs_d.Close
         End If
         If ccarpla <> "" Then
            Set rs_d = New Recordset
            rs_d.Open "SELECT * FROM Cptablas WHERE tabla='28' and codigo='" & ccarpla & "'", Cnx, adOpenStatic, adLockReadOnly
            If Not rs_d.EOF Then
               cDsccarpla = Trim(rs_d!descripcion)
            End If
            rs_d.Close
         End If
         nson = nson + 1
         rt.AddNew
         rt!tarjeta = cTarjeta
         rt!apelpat = cApelP
         rt!apelmat = cApelM
         rt!nombres = cnomb
         rt!Dni = cDni
          rt!CCosto = c_Costo
          rt!Dsccosto = cDscCosto
          rt!CodUpr = cCodUpr
          rt!dscUPR = cDscUpr
         rt!CadHep = ccadhep
         rt!Dsccadhep = cDsccadhep
         rt!NivGuardia = cnivGuardia
         rt!carPla = ccarpla
         rt!DsccarPla = cDsccarpla
         rt!RemTot = nRemTot
         rt!MintTarda = ntTardanza
         rt!Hrsmp = nTMotPer
         rt!Diasfalt = ntdiasfalt
         rt!RemHor = nRemHor
         rt!montoH = ntmontoH       'importe tardanzas
         rt!montoD = ntmontoD       'importe faltas
         rt!montoPer = ntmontoH + ntmontoD      'Total  importe tardanzas + faltas
         rt!Monto = ntmontoH + ntmontoD
         rt.Update
      End If
      If c2dias = "N" And cdiaant = "S" Then
         '-No turno 2dias y marca corresponde turno Dia Anterior
         cdiaant = "N"
      End If
   Loop
   rs.Close
   rt.Close
   cmdcarga.Caption = ""
   If nson > 0 Then
      MsgBox "Total Registros " + Trim(Str(nson)), vbExclamation
      CmdPrint.SetFocus
  End If
  Exit Sub
End Sub

Private Sub TxNroResolu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     TxNroResolu.Text = UCase(TxNroResolu.Text)
     Exit Sub
   End If
End Sub

Private Sub CmdExportaXLS_Click()
    If ChkXLS.Value = 0 Then
       wTpoQuery = "1"
       Call ExportTableToXLS_click
    Else
       wTpoQuery = "1.0"
       wNameXLS = "C:\RPT\ResMP.xls"
       Call Exportar_ADO_Excel
    End If
    Beep 500, 350
    MsgBox "Se generó C:\RPT\ResMP.xls", vbInformation
    Exit Sub
End Sub

Private Sub CmdPrint_Click()
   If nson = 0 Then
      MsgBox "No registra resultados...", vbExclamation
      Exit Sub
   End If
   If Trim(TxNroResolu.Text) = "" Then
      MsgBox "Indique Nro Resolucion...", vbExclamation
      Exit Sub
   End If
   TxNroResolu.Text = UCase(TxNroResolu.Text)
   nTipRpt = 12.1
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Exit Sub
End Sub
Private Sub CmdClear_Click()
  CboMes.Text = "": Txmes.Text = "": TxNombMes.Text = "": TxNroResolu.Text = ""
  CboRegPLH.Text = "": TxRegPLH.Text = ""
  Txfecini.Text = "": TxFecFin.Text = "": TxIniLac.Text = "": TxFinLac.Text = ""
  TxMesdeAno.Text = "": ChkXLS.Value = 1
  CboMes.SetFocus
End Sub
Private Sub cmdsalir_Click()
   Set rs_d = Nothing
   Set rs = Nothing
   Set rt = Nothing
   Set rt1 = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If UCase(Forms(cont).Name) <> "CPMENU" And UCase(Forms(cont).Name) <> "RESMP" Then
           Unload Forms(cont)
        End If
      End If
   Next
   Me.Hide
   Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set rs_d = Nothing
   Set rs = Nothing
   Set rt = Nothing
   Set rt1 = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "ResMP" Then
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

