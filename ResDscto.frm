VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ResDscto 
   Caption         =   "Resolución - Descuentos Inasistencias y Tardanzas"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10695
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "ResDscto.frx":0000
   ScaleHeight     =   2685
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   9840
      Picture         =   "ResDscto.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Cerrar"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox ChkXLS 
      Alignment       =   1  'Right Justify
      Caption         =   "XLS instalado"
      Height          =   195
      Left            =   7080
      TabIndex        =   26
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
      Left            =   3960
      TabIndex        =   25
      Top             =   2040
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
      Left            =   3120
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox TxMsgError 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5040
      TabIndex        =   23
      Top             =   1200
      Width           =   5655
   End
   Begin VB.ComboBox CboRegPlh 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox TxRegPlh 
      Enabled         =   0   'False
      Height          =   225
      Left            =   4560
      MaxLength       =   3
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxNroResolu 
      Height          =   285
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Txmesano 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton CmdExportaXLS 
      Height          =   615
      Left            =   8400
      Picture         =   "ResDscto.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exporta a Xls"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox TxMesActual 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxMesdeAno 
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
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
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TxMES 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "ResDscto.frx":80FF2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   9360
      Picture         =   "ResDscto.frx":81C34
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Otra Consulta"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   975
      Left            =   6240
      Picture         =   "ResDscto.frx":824FE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ubicar Registros"
      Top             =   120
      Width           =   975
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
      Left            =   3120
      TabIndex        =   13
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
      Left            =   3960
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   4920
      TabIndex        =   11
      Top             =   1560
      Width           =   1935
      Begin VB.CommandButton mnuFicImpSelec 
         Height          =   675
         Left            =   120
         Picture         =   "ResDscto.frx":826D9
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
         Picture         =   "ResDscto.frx":82EB7
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
      Left            =   1200
      TabIndex        =   22
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. Resolucion"
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "ResDscto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim rt1 As New ADODB.Recordset
Dim Sw, f_Adm, F_add As Boolean
Dim nLacta, nSonLacta, nposicion, nson, nsondias, nvan As Integer
Dim csexo, cAplica, cTarjetaOLD, cTarjeta, cDni, sql, SQ1, c_tipo, c_Costo, cCodUpr, ccadhep, ccarpla, cnivGuardia, cano, cmes, cDiasfalta As String
Dim unafecha As Date
Dim nRemTot, nRemHor, nHrasProg, ntTardanza, nTDscto, nMontoD, nMontoH As Double
Private Sub Command1_Click()
     c_grupo = "02": c_secuencia = "041"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub
Private Sub Form_Load()
  ResDscto.Caption = "Resolución - Inasistencia y Tardanzas  " & genaño
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
  unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
  Call ObtenerPrimerUltimoDia(unafecha)
  Txfecini.Text = primer
  TxFecFin.Text = ultimo
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
         TxMesdeAno.Text = TxNombMes.Text + " del " & genaño
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
   If IsDate(Txfecini) And IsDate(TxFecFin) Then
      TxMesdeAno = TxNombMes + " de " & genaño
      TxMesAno = TxNombMes + " " & genaño
   Else
      Exit Sub
   End If
   '---On Error Resume Next
   Cnx.Execute "DELETE  FROM CpTmResDsct1"
   Cnx.Execute "DELETE  FROM CpTmResDscto"
   
   SQ1 = "SELECT * FROM CpTmResDsct1"
   Set rt1 = New Recordset
   rt1.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   
   sql = "SELECT * FROM CpTmResDscto"
   Set rt = New Recordset
   rt.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   
   dfi = Day(Txfecini)
   dff = Day(TxFecFin)
   cmes = Right("00" + Trim(Str(Month(Txfecini))), 2)
   unafecha = Txfecini.Text
   Call ObtenerPrimerUltimoDia(unafecha)
   '-carga  DIRECTORIO "TURNO -*
   sql = "SELECT * FROM CpDirectorio  WHERE ano ='" & genaño & "' AND mes='" & cmes & "' ORDER BY tarjeta"
   '--sql = "SELECT * FROM CpDirectorio  WHERE ano ='" & genaño & "' AND mes='" & cmes & "' ANd tarjeta='04793'"
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   nson = 0: nvan = 0
   
     Dim OpenForms
     v_an = 0
   
   Do While Not rs.EOF
   
      '-dar control a windows y permitir multihilo
      v_an = v_an + 1
      If v_an Mod 10 = 0 Then     'if loop has repeated
         OpenForms = DoEvents 'Yield to operating system
      End If
   
      nvan = nvan + 1
      ntHrasProg = 0: ntTardanza = 0: nTDscto = 0: ntdiasfalt = 0: ntHrasFalt = 0: ntmontoH = 0: ntmontoD = 0
      F_add = False
      cDni = "": cApelP = "": cApelM = "": cnomb = "": c_Costo = "": cCodUpr = "": ccadhep = "": cplh = "plh": ctipopla = "": cEstado = ""
      cnivGuardia = "": ccarpla = "": cGrpGen = ""
      nRemTot = 0: nRemHor = 0: cDiasfalta = ""
      cTarjeta = rs!tarjeta
      c_Costo = IIf(IsNull(rs!CCosto), "", rs!CCosto)
      cCodUpr = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)
         
         '-SABER si una columna existe en una tabla
         SQ1 = "SELECT * FROM  INFORMATION_SCHEMA.COLUMNS AS c1  WHERE  c1.column_name = 'cadhep'  AND c1.table_name = 'cpdirectorio'"
         Set rs_2 = New Recordset
         rs_2.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If Not rs_2.EOF Then
            ccadhep = IIf(IsNull(rs!CadHep), "", Trim(rs!CadHep))
         End If
         rs_2.Close
      
      cmdcarga.Caption = "  " + Trim(Str(nson)) & "/" & Trim(Str(nvan))
      sql = "SELECT * FROM Cpgeneral WHERE  (baja is NULL) and tarjeta='" & Trim(cTarjeta) & "'"
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs_d.EOF Then
         '-Carga generales
         cDni = IIf(IsNull(rs_d!documento), "", Trim(rs_d!documento))
         cApelP = IIf(IsNull(rs_d!apelpat), "", rs_d!apelpat)
         cApelM = IIf(IsNull(rs_d!apelmat), "", rs_d!apelmat)
         cnomb = IIf(IsNull(rs_d!nombres), "", rs_d!nombres)
         csexo = IIf(IsNull(rs_d!sexo), "", rs_d!sexo)
         If ccadhep = "" Then
            ccadhep = IIf(IsNull(rs_d!CadHep), "", Trim(rs_d!CadHep))
         End If
         cnivGuardia = IIf(IsNull(rs_d!NivGuardia), "", rs_d!NivGuardia)
         ccarpla = IIf(IsNull(rs_d!carPla), "", rs_d!carPla)
         cGrpGen = IIf(IsNull(rs_d!grpgen), "", rs_d!grpgen)
         nRemTot = IIf(IsNull(rs_d!RemTot), 0, rs_d!RemTot)
         cplh = IIf(IsNull(rs_d!PLH), "plh", rs_d!PLH)
         ctipopla = IIf(IsNull(rs_d!tipoPla), "", rs_d!tipoPla)
         cEstado = IIf(IsNull(rs_d!estado), "", rs_d!estado)
          c_Costo = rs_d!CCosto
          cCodUpr = rs_d!CodUpr
      End If
      rs_d.Close
      
      TxMsgError = ""
      If cEstado <> "A" Then
         TxMsgError = (cTarjeta + " Servidor, No Activo..!")
      ElseIf TxRegPLH <> "" And cplh <> TxRegPLH Then
          TxMsgError = (cTarjeta + " Reg. Laboral [" + cplh + "]")
      '' --MUESTRA servidores con falta informacion caso cadenas, c_costo
      ''ElseIf (cplh = "plh" And Trim(ctipopla) = "") Or Trim(ccadhep) = "" Or (Trim(c_Costo) = "" And Trim(cGrpGen) = "") Then
      ''    TxMsgError.Text = (cTarjeta + " T.Pla:" + ctipopla + " /Cadena:" + cccadhep + " /C.Costo:" + c_Costo + " / Hor.Gen=" + cGrpGen + " del Servidor.")
      Else
          
         '-Afecto X LACTANCIA
         '=============================
         cAplica = "N": nSonLacta = 0  ' caso parto multiple
         TxIniLac = "": TxFinLac = ""
       
         cdiaant = "N"
         For i = dfi To dff
            nmonto = 0:  nvalor = 0
            z = i - 1
            c_turno = Trim(IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4)))
            If IsEmpty(c_turno) Or InStr("*_V_L_RX", Trim(c_turno)) > 0 Then
            Else
               f_Adm = IIf(Trim(c_turno) = "A", True, f_Adm)
               fechaB = Format(primer + z, "yyyymmdd")
               dfentra = primer + z
               dfsale = primer + z
               c_tipo = ""
               p0_ing = "": p0_sal = "": p0_ing1 = "": p0_sal1 = ""   'HORARIO ORIGINAL
               p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": c2dias = "N"
               m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
               nValorH = 0: nValorM = 0: ndscto = 0: nmintarde = 0:  F_falta = False:
               cvalor = "": cAuPrg = "": cDiahora = "": Ctipo = "": choraini = "": chorafin = "": cDescuenta = ""
               cMantGuardia = ""
               
               '-Establece Afecto X LACTANCIA
               '=============================
               cAplica = "N": nSonLacta = 0  ' caso parto multiple
               TxIniLac = "": TxFinLac = ""
               If csexo = "F" Then
                    '-Ubica lactancia al periodo
                    Set rs_d = New Recordset
                    sql = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "'  order by Fecnac"
                    rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                    Do While Not rs_d.EOF
                      dftope = DateAdd("YYYY", 1, rs_d!fecnac)   'Tope Lactancia
                      'MsgBox Format(dftope, "dd/mm/yyyy")
                       If fechaB >= rs_d!fecnac And fechaB <= Format(dftope, "yyyymmdd") Then
                          nSonLacta = nSonLacta + 1
                          '-Carga Parametros de Ajuste horario por Lactancia
                          TxIniLac = rs_d!fecnac
                          TxFinLac = rs_d!FecFIN
                          cAplica = rs_d!Aplica
                       End If
                       rs_d.MoveNext
                    Loop
                    rs_d.Close
               End If
               
               '-Tolerancia x LACTANCIA (I)ngreso/(S)alida
               '=======================
               nLacta = 0
               If cAplica <> "N" Then
                  nLacta = IIf(nSonLacta = 1, 1, 2)
               End If
               '' MsgBox "Aplica:" + cAplica + " [" + Trim(Str(nLacta)) + "] Dia:" + Format(primer + z, "dd/mm/yyyy") + " en " + Format(TxIniLac, "dd/mm/yyyy") + " - " + Format(TxFinLac, "dd/mm/yyyy")
               
               '-HORARIO prog en dia -**
               sql = "SELECT * FROM CpHor" & genaño & " WHERE Tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
               Set rs_d = New Recordset
               rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If rs_d.EOF Then
                  If Trim(cGrpGen) <> "" Then
                     rs_d.Close
                     sql = "SELECT * FROM CpHor" & genaño & " WHERE grupo='" & cGrpGen & "' AND fecha ='" & fechaB & "'"
                     Set rs_d = New Recordset
                     rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  End If
               End If
               If Not rs_d.EOF Then
                  cvalor = IIf(IsNull(rs_d!Horario), "", rs_d!Horario)
                  rs_d.Close
                  Sq0 = "SELECT * FROM CpHorarioNew  WHERE Codigo='" & cvalor & "'"
                  Set rs_d = New Recordset
                  rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  If rs_d.EOF Then
                     rs_d.Close
                     Set rs_d = New Recordset
                     Sq0 = "SELECT * FROM CpHorarios  WHERE Codigo='" & cvalor & "'"
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
                     ntHrasProg = ntHrasProg + (nValorH + nValorM / 60)    'acumula
                     If (p0_ing > p0_sal And Val(p0_ing1) = 0) Then           '======>Or p_horas = "2400"
                        c2dias = "S"
                     ElseIf Val(p0_sal1) > 0 And p0_ing < p0_ing1 Then
                            c2dias = "S"
                     End If
                  End If
               End If
               rs_d.Close
              
               '-AUTORIZACIONES: Excepto Motivos Personales <> I/S
               cDiahora = "X"
               cMantGuardia = "N"
               sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and tarjeta='" & Trim(cTarjeta) & "' and FechaIni<='" + fechaB + "' and FechaFin>='" + fechaB + "'"
               Set rs_d = New Recordset
               rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               Do While Not rs_d.EOF
                  If Not rs_d.EOF Then
                     cAuPrg = Trim(rs_d!Codigo)
                     cDiahora = rs_d!diahora
                     Ctipo = IIf(IsNull(rs_d!tipo), "", rs_d!tipo)            'I/S/O   caso horas
                     choraini = IIf(IsNull(rs_d!horaini), "", rs_d!horaini)
                     chorafin = IIf(IsNull(rs_d!horafin), "", rs_d!horafin)
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
                  rs_d.MoveNext
               Loop
               rs_d.Close
               If Val(p_horas) > 0 Then               '.....X
                  Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
                  '-Marcas del dia, solo si Justifica "D"
                   p_Sale = ""
                  If cDiahora <> "D" Then
                     If Val(p_sal1) > 0 Or Val(p_sal) < Val(p_ing) Then
                        'Ubica marcas 2dias actual-siguiente  turno t/n, n/d
                        fecha1 = Format(primer + z + 1, "yyyymmdd")
                        Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha >='" & fechaB & "' AND fecha <='" & fecha1 & "' ORDER BY fecha,hora"
                        cdiaant = "S"
                        p_Sale = p_sal
                        If Val(p_sal1) > 0 Then
                          p_Sale = p_sal1         'Corrige fin turno
                        End If
                     End If
                  End If
                  Set rs_d = New Recordset
                  rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  If Not rs_d.EOF Then
                     '-ASIGNA marcas basado en Horario
                     d_fentra = rs_d!fecha             '========> 1era marca del servidor para turno programado
                     Do While Not rs_d.EOF
                        cInicio = p_ing
                        cFin = rs_d!hora
                        dfsale = rs_d!fecha
                        If Trim(m_ing) = "" Then
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
                               '-sal Turno partido  1300-1900/1900-0700
                               m_sal1 = rs_d!hora
                               
                        '-Check carga marcaciones sin valor para turno
                        ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then       'And Val(m_ing1) > 0
                               m_sal1 = rs_d!hora
                               
                        End If
                        rs_d.MoveNext
                     Loop
                  End If
                  rs_d.Close
                              
                 '-Muestra  marcas realizadas
                 'If z = 7 And Txtarjeta = "04793" Then    'And ( z = 21 Or z = 22 Or z = 23)
                 '   MsgBox "Horas: " + Trim(Str(Day(primer + z))) + " -> " + p_ing + "-" + p_sal + ": " + m_ing + " , " + m_sal + " , " + m_ing1 + " , " + m_sal1
                 'End If
               
                  '-Evalua MARCAS Ingreso
                  If cDiahora = "D" And cDescuenta = "S" Then
                     F_falta = True
                     
                  'ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And p_sal <> p_ing1 Then       '======>
                  ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And Val(p_ing1) > 0 And p_sal <> p_ing1 Then  '======>
                       ' Caso Horario:2000-0800   marca 0805-2010   O marca 0749-1404 1930-0820
                         F_falta = True
                     
                  ElseIf (Val(m_ing) + Val(m_ing1) = 0) Then
                     'No marco Ing,   valida vs autoriza
                     If cDiahora = "H" And Ctipo = "I" Then
                        cInicio = p_ing
                        cFin = choraini
                        If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                           F_falta = True
                           c_tipo = "e"
                           cDescuenta = "S"
                        End If
                     ElseIf cDiahora = "D" Then    'valida autoriza Dia ComP / Pierde
                        If cDescuenta = "S" Then
                           F_falta = True
                           c_tipo = "D"
                        End If
                     Else
                        F_falta = True
                        c_tipo = "i"
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
                            If (m_sal > m_ing And m_sal < p_ing) And p_horas <> 2400 Then      '¿2400?
                               '-salida Antes de Turno
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
                  
                 '-Muestra  marcas Falta ¿? Txtarjeta = "61961" And
                 'If (z = 7) And Txtarjeta = "04793" Then      'Or z = 15
                 '   MsgBox Trim(Str(Day(primer + z))) + " " + cDiahora + " ¿ING. Falta? " + IIf(F_falta, "S", "N") + " [" + c_tipo + "] -> " + p_ing + "-" + p_sal + ": " + m_ing + " , " + m_sal + " , " + m_ing1 + " , " + m_sal1
                 'End If
                  
                  '-Evalua MARCAS Salida
                  If F_ALTA Then
                  ElseIf (Val(m_sal) + Val(m_sal1) = 0) Then
                     '-No marco  Sal   valida vs autoriza
                     If cDiahora = "H" And Ctipo = "S" Then
                        If (Val(p_sal) - Val(chorafin) > 0) Or (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                           F_falta = True
                           c_tipo = "f"
                           cDescuenta = "S"
                        End If
                     ElseIf cDiahora = "D" Then     'valida autoriza Dia Comp/ pierde
                        If cDescuenta = "S" Then
                           F_falta = True
                           c_tipo = "2"
                        End If
                     Else
                        F_falta = True
                        c_tipo = "s"
                     End If
                  ElseIf Val(p_sal1) > 0 And p0_sal <> p0_ing1 Then    'p_horas<>"2400"
                         If Val(m_sal1) = 0 Then
                             F_falta = True
                         Else
                             If Val(p_sal1) > Val(m_sal1) Then
                                F_falta = True
                             End If
                        End If
                     
                  ElseIf Val(m_sal) > 0 And Val(p0_sal) > Val(m_sal) Then         'p_sal
                     'Marco antes de, Autoriza H
                     If cDiahora = "H" And Ctipo = "S" Then     '====> And cTipo = "S"   ** TO CHECK **
                        cInicio = p_sal
                        cFin = chorafin
                        If Val(p_sal) - Val(chorafin) > 0 Then
                           F_falta = True
                        End If
                     ElseIf cDiahora = "X" Then
                        '-----cInicio = p_sal
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
                           c_tipo = "x"
                           
                        ElseIf Val(p_sal1) - Val(cFin) > 0 Then
                               F_falta = True
                               c_tipo = "x"
                               
                        End If
                     End If
                  End If
                  
                 '-Muestra  marcas Falta ¿? Txtarjeta = "61961" And
                 'If (z = 7) And Txtarjeta = "04793" Then    'Or z = 15
                 '   MsgBox Trim(Str(Day(primer + z))) + " " + cDiahora + " ¿SAL. Falta? " + IIf(F_falta, "S", "N") + " [" + c_tipo + "] -> " + p_ing + "-" + p_sal + ": " + m_ing + " , " + m_sal + " , " + m_ing1 + " , " + m_sal1
                 'End If
                  
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
                           If nmintarde < 6 Then
                              nmintarde = 0
                           End If
                           If nmintarde >= 6 And nmintarde <= 10 Then
                              ndscto = 10
                           End If
                           If nmintarde >= 11 And nmintarde <= 20 Then
                              ndscto = 20
                           End If
                           If nmintarde >= 21 And nmintarde < 31 Then
                              ndscto = 30
                           End If
                           If nmintarde > 30 Then
                              F_falta = True
                              c_tipo = "F"
                           End If
                        End If
                     
                     End If
                  End If
                  
                  If f_Adm Then
                     nRemHor = Round(nRemTot / (30 * 7.25), 2)
                  Else
                     nRemHor = Round(nRemTot / (30 * 6), 2)
                  End If
                  If F_falta Then
                     If cAuPrg = "" Or cDiahora = "H" Or (cAuPrg <> "" And cAuPrg <> "MP" And cDiahora = "D" And cDescuenta = "S") Then
                        'No incluye faltas "MP"
                        nsondias = 1    'Turnos default "M", "T" "A"
                        If InStr("DN", Left(c_turno, 1)) > 0 Or InStr(c_turno, "MT") > 0 Or InStr(c_turno, "MX") > 0 Then
                           nsondias = nsondias + 1
                        ElseIf InStr("HC", Left(c_turno, 1)) > 0 Then
                            nsondias = nsondias + 2
                        ElseIf InStr("P", Left(c_turno, 1)) > 0 Then
                            nsondias = nsondias + 3
                        End If
                        ntdiasfalt = ntdiasfalt + nsondias
                        ntmontoD = ntmontoD + Round((nsondias * IIf(f_Adm, 7.25, 6) * nRemHor), 2)
                     End If
                  Else
                     ntTardanza = ntTardanza + nmintarde
                     nTDscto = nTDscto + ndscto
                     ntmontoH = ntmontoH + (ndscto / 60) * nRemHor
                  End If
                  If F_falta Or ndscto >= 10 Then
                     F_add = True
                     rt1.AddNew
                     rt1!tarjeta = cTarjeta
                     rt1!nombdia = Left(WeekdayName(Weekday(primer + z)), 3)
                     rt1!fecha = primer + z
                     rt1!turno = c_turno
                     rt1!Horario = cvalor
                     rt1!CCosto = c_Costo
                     rt1!CodUpr = cCodUpr
                     '-Horario ORIGINAL
                     '=================
                     rt1!ping = "" + p0_ing
                     rt1!psal = "" + p0_sal
                     rt1!ping1 = "" + p0_ing1
                     rt1!psal1 = "" + p0_sal1
                     rt1!hrasprog = 0 + (nValorH + nValorM / 60)
                     rt1!ming = "" + m_ing
                     rt1!msal = "" + m_sal
                     rt1!ming1 = "" + m_ing1
                     rt1!msal1 = "" + m_sal1
                     rt1!tardanza = 0 + IIf(F_falta, 0, nmintarde)
                     rt1!DESCUENTO = 0 + IIf(F_falta, 0, ndscto)
                     rt1!autorizac = "" + cAuPrg
                     rt1!descuenta = "" + cDescuenta
                     rt1!diahora = "" + cDiahora
                     rt1!tipo = "" + c_tipo
                     rt1!horaini = "" + choraini
                     rt1!horafin = "" + chorafin
                     rt1!falta = IIf(F_falta, "S", "N")
                     rt1.Update
                     If F_falta Then
                        cDiaF = Right("0" + Trim(Str(Day(primer + z))), 2)       ' d/falto
                        If InStr(cDiasfalta, cDiaF) = 0 Then                     ' no existe => Agrega a Lista
                           cDiasfalta = cDiasfalta + cDiaF + ","
                        End If
                     End If
                  End If
                  If c2dias = "N" And cdiaant = "S" Then
                     '-No turno 2dias y marca corresponde turno Dia Anterior
                     cdiaant = "N"
                  End If
               End If      '.....X
            End If
         Next
         
      End If
      If F_add Then
         If ntHrasProg > 0 And (ntTardanza > 60 Or ntdiasfalt > 0) Then
            nson = nson + 1
            If ntTardanza > 60 Then
               nvalor = Int(nTDscto / 60)
               If nTDscto Mod 60 > 30 Then
                   nvalor = nvalor + 1
               End If
               nTDscto = nvalor
            Else
               nTDscto = 0
               ntmontoH = 0
            End If
            cDscCosto = ""
            If c_Costo <> "" Then
               Set rs_d = New Recordset
               rs_d.Open "SELECT * FROM Cpccosto WHERE codigo='" & c_Costo & "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rs_d.EOF Then
                  cDscCosto = Trim(rs_d!descripcion)
               End If
               rs_d.Close
            End If
            cDscUpr = ""
            If cCodUpr <> "" Then
               Set rs_d = New Recordset
               rs_d.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rs_d.EOF Then
                  cDscUpr = Trim(rs_d!descripcion)
               End If
               rs_d.Close
            End If
            cDsccadhep = ""
            If ccadhep <> "" Then
               Set rs_d = New Recordset
               rs_d.Open "SELECT * FROM Cpcadena WHERE ano='" & genaño & "' and cadena='" & ccadhep & "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rs_d.EOF Then
                  cDsccadhep = Trim(rs_d!descripcion)
               End If
               rs_d.Close
            End If
            cDsccarpla = ""
            If ccarpla <> "" Then
               Set rs_d = New Recordset
               rs_d.Open "SELECT * FROM Cptablas WHERE tabla='28' and codigo='" & ccarpla & "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rs_d.EOF Then
                  cDsccarpla = Trim(rs_d!descripcion)
               End If
               rs_d.Close
            End If
            
            rt.AddNew
            rt!tarjeta = cTarjeta
            rt!Dni = cDni
            rt!apelpat = cApelP
            rt!apelmat = cApelM
            rt!nombres = cnomb
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
            rt!RemHor = nRemHor
            rt!hrasprog = ntHrasProg
            rt!tardanza = ntTardanza
            rt!DESCUENTO = nTDscto
            rt!Diasfalt = ntdiasfalt
            rt!montoH = ntmontoH        'importe tardanzas
            rt!montoD = ntmontoD        'importe Faltas
            rt!Monto = ntmontoH + ntmontoD      'total importe
            rt!DiasFalta = cDiasfalta
            rt.Update
         End If
         
      End If
      rs.MoveNext
   Loop
   rs.Close
   rt1.Close
   rt.Close
   cmdcarga.Caption = ""
   If nson > 0 Then
      MsgBox "Total Registros " + Trim(Str(nson)) & " de:" & Trim(Str(nvan)), vbExclamation
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
       wTpoQuery = "0"
       Call ExportTableToXLS_click
    Else
       wTpoQuery = "0.1"
       wNameXLS = "C:\RPT\ResDscts.xls"
       Call Exportar_ADO_Excel
    End If
    Beep 500, 350
    MsgBox "Se generó C:\RPT\ResDscts.xls", vbInformation
    Exit Sub
End Sub

Private Sub CmdPrint_Click()
   If nson = 0 Then
      MsgBox "No registra resultados...", vbExclamation
      Exit Sub
   End If
   If TxNroResolu.Text = "" Then
      MsgBox "Indique Nº resolución...", vbExclamation
      Exit Sub
   End If
   TxNroResolu.Text = UCase(TxNroResolu.Text)
   nTipRpt = 12
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Exit Sub
End Sub
Private Sub CmdClear_Click()
  CboMes.Text = "": Txmes = "": TxNombMes = "":: TxNroResolu = ""
  CboRegPLH.Text = "": TxRegPLH = ""
  Txfecini = "": TxFecFin = "": TxIniLac = "": TxFinLac = ""
  TxMesdeAno = "": TxMesAno = "": ChkXLS.Value = 1
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
        If UCase(Forms(cont).Name) <> "CPMENU" And UCase(Forms(cont).Name) <> "RESDSCTO" Then
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
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "ResDscto" Then
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

