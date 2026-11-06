VERSION 5.00
Begin VB.Form DirMesNutricion 
   Caption         =   "Nurtición - Directorio Mensual de ASISTENCIALES"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txano 
      Height          =   285
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   195
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   735
      Left            =   5160
      Picture         =   "DirMesNutricion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cerrar"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton CmdProcDatos 
      Caption         =   "Se enviará horarios de guardia Asistencial del perìodo indicado."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Lbldestino 
      BackStyle       =   0  'Transparent
      Caption         =   "Lbldestino"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label LblRec 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LblRec"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AÑO :"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MES :"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "DirMesNutricion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cano, cmes, Criterio, cDni As String
Dim i, wdato, nson, nvan, nmes, nano As Integer
Dim rs As New ADODB.Recordset
Dim r_d As New ADODB.Recordset
Dim rsr As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim cDSource2, cICatal2, CString2, cUsr2, cPwd2 As String
Dim nLacta, nSonLacta, nposicion, nvalor As Integer
Dim csexo, cAplica, sql, sqr, sqg, cTarjeta, cNombre, cTurno, chhmm As String
Dim unafecha As Date
Dim lgrabo As Boolean

Private Sub Form_Load()

   cDSource2 = "192.168.1.198"    ' "192.168.1.198"     '
   cICatal2 = "Heptrabajadores"
   cUsr2 = "sa"
   cPwd2 = "P4sscomedor"       '  "P4ssw0rd"    '
  LblRec.Caption = ""
  Lbldestino.Caption = cDSource2
  Txano = genaño
  cano = genaño
  Set rs = New Recordset
  rs.Open "SELECT * FROM cptablas WHERE tabla='32' AND codigo<>'..' ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboMes.ListCount - 1
           If CboMes.List(i) = rs.Fields("Descripcion") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboMes.AddItem rs.Fields("Descripcion")
       End If
       rs.MoveNext
  Wend
 
  Txmes = Right("0" + Trim(Str(genmes)), 2)
  CboMes.ListIndex = genmes - 1
  Call Cbomes_KeyPress(13)
 
   Me.Caption = "Nutrición: Guardias Asistenciales "
   
End Sub
Private Sub Cbomes_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboMes.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboMes_Click()
  Cbomes_KeyPress (13)
End Sub
Private Sub Cbomes_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboMes.Text) Then
         CboMes.SetFocus
      Else
         j = CboMes.ListIndex
         If j <> -1 Then
            cmes = Right("0" + Trim(Str(j + 1)), 2)
            Txmes = cmes
            unafecha = "01/" + cmes + "/" + cano
            Call ObtenerPrimerUltimoDia(unafecha)            'primer/ultimo
         End If
      End If
      Exit Sub
   End If
End Sub

Private Sub CmdProcDatos_Click()
   LblRec.Caption = ""
   wdato = 0
   nson = 0: nvan = 0
   
   Criterio = "  ano='" + cano + "' and mes='" + cmes + "' and FLAG='0' "
   Set r_d = New ADODB.Recordset
   sqg = "SELECT *  FROM CpDirectorio WHERE " & Criterio & " ORDER BY tarjeta"
   r_d.Open sqg, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If r_d.EOF Then
      MsgBox "Periodo " + cmes + "/" + cano + " sin información..", vbCritical, "* IMPORTANTE *"
      r_d.Close
      Exit Sub
   End If
   
   Dim cnx2  As ADODB.Connection
   CString2 = "Provider=" + cProvider + ";Data Source='" + cDSource2 + "';" & _
       "initial catalog=" + cICatal2 + "; user id='" + cUsr2 + "'; password='" + cPwd2 + "';"
   Set cnx2 = New ADODB.Connection
   With cnx2
      .ConnectionString = CString2
      .Open
   End With
   cnx2.CommandTimeout = 300
   cnx2.BeginTrans
   cnx2.Execute "DELETE CPDirTrabajadores WHERE ano='" + cano + "' and mes='" + cmes + "' and servicio<>'FAMILIAR'"
   cnx2.CommitTrans

   Set rsr = New Recordset
   sqr = "SELECT * FROM CPDirTrabajadores"       'Directorio mensual
   rsr.Open sqr, cnx2, adOpenStatic, adLockOptimistic, adCmdText
   rsr.Close
   
   '-MsgBox "SELECT * FROM CpDirectorio WHERE " & Criterio & " ORDER BY tarjeta "
   
   cnx2.BeginTrans
   Dim OpenForms
   v_an = 0
   diai = Day(primer)     'Dias de inicio y cierre del periodo consultado
   diaf = Day(ultimo)     'Day(ultimo)
   Do While Not r_d.EOF         '=====> Servidores
      '-dar control a windows y permitir multihilo
      v_an = v_an + 1
      If v_an Mod 10 = 0 Then     'if loop has repeated
         OpenForms = DoEvents     'Yield to operating system
      End If
   
      capelPat = "": capelMat = "": cnombres = "": csexo = "":  wdscUpr = ""
      cdescEspec = "SIN ESPECIALIDAD":  cCadenaDias = ""
      cTarjeta = Trim(r_d!tarjeta)     '- servidor
      cCodUpr = IIf(IsNull(r_d!CodUpr), "", Trim(r_d!CodUpr))
      cCargo = IIf(IsNull(r_d!Cargo), "", Trim(r_d!Cargo))
         
      Set rs_d = New Recordset
      sql = "SELECT * FROM CpGeneral WHERE tarjeta='" + cTarjeta + "'"
      rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs_d.EOF Then
         capelPat = IIf(IsNull(rs_d!apelpat), "", Trim(rs_d!apelpat))
         capelMat = IIf(IsNull(rs_d!apelmat), "", Trim(rs_d!apelmat))
         cnombres = IIf(IsNull(rs_d!nombres), "", Trim(rs_d!nombres))
         cDni = Trim(rs_d!documento)
         csexo = IIf(IsNull(rs_d!sexo), "", Trim(rs_d!sexo))
      End If
      rs_d.Close
      
    If cDni <> "" Then  '<====  servidor identificado para comedor
      CmdProcDatos.Caption = Left(capelPat + " " + cnombres, 20)
      If cCodUpr <> "" Then
          Set rs_d = New Recordset
          rs_d.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs_d.EOF Then
             wdscUpr = Trim(rs_d!descripcion)   ' Trim(rs_d!servicioSusalud)
          End If
          rs_d.Close
       End If
       For i = diai To diaf      '<<===dias calendario mes analizar, max 31
            z = i - 1
            FechaB = Format(primer + z, "yyyymmdd")
            dfentra = primer + z
            cTurno = "": cHoraI1 = "": cHoraF1 = "":: cHoraI2 = "": cHoraF2 = "": p_horas = ""
            p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = ""
            
            '-Establece Afecto X LACTANCIA
            '=============================
            cAplica = "N": nSonLacta = 0
            TxIniLac = "": TxFinLac = ""
            If csexo = "F" Then
               '-Ubica lactancia al periodo
               Set rs_d = New Recordset
               sql = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "'  order by fecnac"
               rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               Do While Not rs_d.EOF
                  dftope = DateAdd("YYYY", 1, rs_d!fecnac)   'Tope Lactancia
                  'MsgBox Format(dftope, "dd/mm/yyyy")
                  If FechaB >= rs_d!fecnac And FechaB <= Format(dftope, "yyyymmdd") Then
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
                             
            cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
            '-HORARIO prog en dia.
            sql = "SELECT * FROM CpHor" + cano + "  WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & FechaB & "'"
            Set rs = New Recordset
            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If rs.EOF Then
               If Trim(cGrpGen) <> "" Then
                  rs.Close
                  sql = "SELECT * FROM CpHor" + cano + "  WHERE grupo='" & cGrpGen & "' AND fecha ='" & FechaB & "'"
                  Set rs = New Recordset
                  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               End If
            End If
            If Not rs.EOF Then
               cvalor = IIf(IsNull(rs!Horario), "", rs!Horario)
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
                  p_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                  p_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
                  p_sal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
                                                
                  '-horario ORIGINAL
                  '=================
                  p0_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
                  p0_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                  p0_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
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
                  nHrProg = (nValorH + nValorM / 60)
               End If
               rs_d.Close
            End If
            rs.Close
            
            If Val(p_horas) > 0 Then
               cCadenaDias = cCadenaDias + Right("0" + Trim(Str(i)), 2)
               nvez = 1
               cHoraI = Left(p_ing, 2) + ":" + Right(p_ing, 2)   'p0_ing
               cHoraF = Left(p_sal, 2) + ":" + Right(p_sal, 2)
               If Val(p_ing1) > 0 Then
                  cHoraI1 = Left(p_ing1, 2) + ":" + Right(p_ing1, 2) 'p0_ing1
                  cHoraF1 = Left(p_sal1, 2) + ":" + Right(p_sal1, 2)
                  nvez = 2
               End If
               'nson = 7 + (z * 2)
               'nson1 = nson + 1
               '--ccIng = rRS.Fields(nson).Name
               '--CCsal = rRS.Fields(nson + 1).Name
               For k = 1 To nvez
                  lgrabo = False
                  If k = 1 Then
                     If Val(Left(cHoraI, 2)) < Val(Left(cHoraF, 2)) Then
                        nvalor = Val(Left(cHoraF, 2)) - Val(Left(cHoraI, 2))
                        If nvalor >= 12 Then
                           Set rsr = New Recordset
                           sqr = "SELECT * FROM CpDirTrabajadores  where DNI='" + cDni + "'"      'Directorio mensual
                           rsr.Open sqr, cnx2, adOpenStatic, adLockOptimistic, adCmdText
                           'If rsr.EOF() Then         '1er dia programado
                              rsr.AddNew
                              rsr!ano = cano
                              rsr!mes = cmes
                              rsr!Dia = Right("0" + Trim(Str(i)), 2)
                              rsr!Dni = cDni
                              rsr!apellidos = Trim(capelPat) + " " + Trim(capelMat)
                              rsr!nombres = Trim(cnombres)
                              rsr!Servicio = Trim(wdscUpr)
                              rsr!tarjeta = cTarjeta
                              rsr!fechareg = Date
                              rsr!horareg = Format(Time, "HH:mm:ss")
                              rsr!usuarioReg = "Admin"
                              rsr!estacionREg = w_clienName
                           'End If
                           rsr!entrada = cHoraI
                           rsr!salida = cHoraF
                           lgrabo = True
                        End If
                     Else
                        nvalor = Val(Left(cHoraF, 2)) - Val(Left(cHoraI, 2)) + 24
                        If nvalor >= 12 Then
                           Set rsr = New Recordset
                           sqr = "SELECT * FROM CpDirTrabajadores  where DNI='" + cDni + "'"      'Directorio mensual
                           rsr.Open sqr, cnx2, adOpenStatic, adLockOptimistic, adCmdText
                           'If rsr.EOF() Then         '1er dia programado
                              rsr.AddNew
                              rsr!ano = cano
                              rsr!mes = cmes
                              rsr!Dia = Right("0" + Trim(Str(i)), 2)
                              rsr!Dni = cDni
                              rsr!apellidos = Trim(capelPat) + " " + Trim(capelMat)
                              rsr!nombres = Trim(cnombres)
                              rsr!Servicio = Trim(wdscUpr)
                              rsr!tarjeta = cTarjeta
                              rsr!fechareg = Date
                              rsr!horareg = Format(Time, "HH:mm:ss")
                              rsr!usuarioReg = "Admin"
                              rsr!estacionREg = w_clienName
                           'End If
                           rsr!entrada = cHoraI
                           rsr!salida = cHoraF
                           lgrabo = True
                        End If
                     End If
                  Else
                     Set rsr = New Recordset
                     sqr = "SELECT * FROM CpDirTrabajadores   where DNI='" + cDni + "'"
                     rsr.Open sqr, cnx2, adOpenStatic, adLockOptimistic, adCmdText
                     nvalor = Val(Left(cHoraF, 2)) - Val(Left(cHoraI, 2)) + 24
                     If nvalor >= 12 Then
                        rsr.AddNew
                        rsr!ano = cano
                        rsr!mes = cmes
                        rsr!Dia = Right("0" + Trim(Str(i)), 2)
                        rsr!Dni = cDni
                        rsr!apellidos = Trim(capelPat) + " " + Trim(capelMat)
                        rsr!nombres = Trim(cnombres)
                        rsr!Servicio = Trim(wdscUpr)
                        rsr!tarjeta = cTarjeta
                        rsr!fechareg = Date
                        rsr!horareg = Format(Time, "HH:mm:ss")
                        rsr!usuarioReg = "Admin"
                        rsr!estacionREg = w_clienName
                        rsr!entrada = cHoraI1
                        rsr!salida = cHoraF1
                        lgrabo = True
                     End If
                  End If
                  If lgrabo Then
                     rsr.Update
                     rsr.Close
                  End If
               Next k
            End If
            
       Next i            '<<=== sgt Dia calendario mes analizar, max 31
       
       If Len(cCadenaDias) > 0 Then
          wdato = wdato + 1
          LblRec.Caption = "van..." + Trim(Str(wdato)) + " servidores."
       End If
       
     End If   '====>  fin servidor identificado
     r_d.MoveNext     'Sgt Servidor
   Loop
   r_d.Close
   
   If wdato > 1 Then
     '-GRABA REGSUMARIO
     Set rs = New Recordset
     sql = "SELECT * FROM registroEnvios"
     rs.Open sql, cnx2, adOpenStatic, adLockOptimistic, adCmdText
     rs.AddNew
     rs!fecha = Date
     rs!hora = Format(Time, "HH:mm:ss")
     rs!ano = cano
     rs!mes = cmes
     rs!Servidores = wdato
     rs!usuario = "Admin"
     rs!estacion = w_clienName
     rs.Update
     rs.Close
   End If
   
   cnx2.CommitTrans
   
   Beep 500, 350
   Beep 500, 350
   CmdProcDatos.Caption = "Se Enviará guardia de servidores Asistenciales para el período indicado."
   If wdato > 0 Then
      LblRec.Caption = "Son " & Trim(Str(wdato)) + " servidores actualizados"
   End If
   cmdsalir.SetFocus
End Sub
   
  
Public Function TimeMsg()
'Sustituye 4000 por el nº de milisegundos que quieres que el mensaje aparezca
'1000 milisegundos=1 segundo
   hi = Format(Time, "hh:mm:ss")
   nseg = 0
   Do While nseg < 4
      hf = Format(Time, "hh:mm:ss")
      nseg = DateDiff("S", hi, hf)
   Loop
End Function

Private Sub cmdsalir_Click()
   Unload Me
   Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   Me.Hide
End Sub


