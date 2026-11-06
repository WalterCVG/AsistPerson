VERSION 5.00
Begin VB.Form ChkVarios 
   Caption         =   "Matching Varios Procesos"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   2040
      TabIndex        =   9
      Top             =   1080
      Width           =   3975
      Begin VB.ComboBox CboMes 
         Height          =   315
         Left            =   720
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox TxPeriodo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         MaxLength       =   2
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CmdChkDir 
         Caption         =   "Check Drirectorio"
         Height          =   855
         Left            =   2760
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label_1 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.TextBox Txmes 
      Height          =   285
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Cmdfecha 
      Caption         =   "ChkFecha"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdRemuHep 
      Caption         =   "ACTUALIZA   HEP->RemTot"
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton ChkCAS 
      Caption         =   "A c t u a l i z a  CAS"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   5160
      Picture         =   "ChkVarios.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cerrar"
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Cmdprocesa 
      Caption         =   "P r o c e s a Tarjeta->CPgrupos"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Lblfecha 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Mes"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Lbldato 
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "ChkVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim cTarjeta, cCodniv, clibele, cgrupo, cPatron, capelPat, capelMat, cnombres As String
Dim nson, npos, nlen As Integer
Dim nremunera As Double
Dim dfi, dff, z As Integer
Dim unafecha As Date
Private Sub Form_Load()
  Set rs = New Recordset
  rs.Open "SELECT * FROM cptablas WHERE tabla='32' and codigo <>'..' ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboMes.ListCount - 1
           If CboMes.List(i) = rs.Fields("Descripcion") & " :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboMes.AddItem rs.Fields("Descripcion") & " :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
End Sub

Private Sub CmdProcesa_Click()
   Set rs = New Recordset
   rs.Open "SELECT * FROM CpGrupos ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
   nson = 0: nvan = 0
   Do While Not rs.EOF
      cgrupo = rs!grupo
      If InStr("G01_G02_G03_G04_G05", cgrupo) Then
      Else
      nvan = nvan + 1
      cPatron = Trim(rs!descripcion)
      nlen = Len(cPatron)
      npos = InStr(cPatron, " ")
      capelPat = "": capelMat = ""
      If npos = 0 Then
      Else
        capelPat = Left(cPatron, npos - 1)
        cPatron = Right(cPatron, nlen - npos)
        npos = InStr(cPatron, " ")
        If npos = 0 Then
           capelMat = cPatron
        Else
           capelMat = Left(cPatron, npos - 1)
        End If
      End If
      CmdProcesa.Caption = capelPat + " " + capelMat
      Set rs_d = New Recordset
      sql = "SELECT * FROM  CpGeneral where (baja is NULL) and apelpat='" & capelPat & "'"
      rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      While Not rs_d.EOF
        If Trim(rs_d!apelmat) = capelMat Then
           cTarjeta = rs_d!tarjeta
           sql = "UPDATE CPgrupos  SET tarjeta='" & cTarjeta & "' WHERE  (baja is NULL) and grupo='" & cgrupo & "'"
           Cnx.Execute sql
           sql = "UPDATE CPgeneral SET grpper='" & cgrupo & "' WHERE  (baja is NULL) and tarjeta='" & cTarjeta & "'"
           Cnx.Execute sql
           nson = nson + 1
           CmdProcesa.Caption = "Van : " + CStr(nson)
        End If
        rs_d.MoveNext
      Wend
      rs_d.Close
      End If
      rs.MoveNext
   Loop
   rs.Close
   CmdProcesa.Caption = "P r o c e s a   Tarjeta->CpGrupos"
   MsgBox ("Se actualizo " & CStr(nson) & " de " & CStr(nvan))
   Call cmdsalir_Click
End Sub


Private Sub ChkCAS_Click()
   Set rs = New Recordset
   rs.Open "SELECT * FROM cpGeneral WHERE  (baja is NULL) and estado= 'A' ORDER BY apelpat,apelmat", Cnx, adOpenStatic, adLockReadOnly
   nhay = 0: nson = 0: nvan = 0
   Do While Not rs.EOF
      nvan = nvan + 1
      capelPat = Trim(rs!apelpat)
      capelMat = Trim(rs!apelmat)
      cnombres = Trim(rs!nombres)
      cPatron = capelPat + " " + capelMat
      cTarjeta = rs!tarjeta
      ChkCAS.Caption = cPatron
      If UCase(rs!apelmat) = "D'BROT" Then
      Else
        Set rs_d = New Recordset
        sql = "SELECT * FROM dat_cas WHERE Nombre LIKE '" & cPatron & "%'  ORDER BY  nombre"
        rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not rs_d.EOF
           If InStr(rs_d!nombre, cPatron) > 0 Then
              If InStr(rs_d!nombre, cnombres) > 0 Then
                 nremunera = 0: clibele = "": cCond = "": ctpo = "": cCodniv = "": Chor = "": dFecnac = "": dFecalt = ""
                 For i = 0 To rs_d.Fields.Count - 1
                    On Error Resume Next
                    ccampo = UCase(rs_d.Fields(i).Name)
                    If ccampo = "LIBELE" Then
                       clibele = rs_d.Fields(i)
                    ElseIf ccampo = "REMTOT" Then
                           nremunera = IIf(IsNull(rs_d.Fields(i)), 0, rs_d.Fields(i))
                    ElseIf ccampo = "COND" Then
                          cCond = rs_d.Fields(i)
                    ElseIf ccampo = "TPO" Then
                           ctpo = rs_d.Fields(i)
                    ElseIf ccampo = "HOR" Then
                           Chor = rs_d.Fields(i)
                    ElseIf ccampo = "FECNAC" Then
                           dFecnac = rs_d.Fields(i)
                    ElseIf ccampo = "FECALT" Then
                           dFecalt = rs_d.Fields(i)
                    End If
                 Next i
                 sql = "UPDATE CpGeneral  SET documento='" & clibele & "', plh='CAS', tipoPla='" & ctpo & "', grpGen='" & Chor & "', RemTot=" & nremunera & " WHERE  (baja is NULL) and tarjeta='" & cTarjeta & "'"
                 Cnx.Execute sql
                 nson = nson + 1
                 ChkCAS.Caption = "Van : " + CStr(nson)
                 Exit Do
              End If
           End If
           rs_d.MoveNext
        Loop
        rs_d.Close
      End If
      rs.MoveNext
   Loop
   rs.Close
   ChkCAS.Caption = "A C T U A L I Z A  CAS"
   MsgBox ("Se actualizo " & CStr(nson) & " de " & CStr(nvan))
   Call cmdsalir_Click
End Sub


Private Sub CmdRemuHep_Click()
   Set rs = New Recordset
   rs.Open "SELECT * FROM cpGeneral WHERE (baja is NULL) and  ESTADO='A' ORDER BY apelpat,apelmat", Cnx, adOpenStatic, adLockReadOnly
   nhay = 0: nson = 0: nvan = 0
   Do While Not rs.EOF
      nvan = nvan + 1
      capelPat = Trim(rs!apelpat)
      capelMat = Trim(rs!apelmat)
      cnombres = Trim(rs!nombres)
      cPatron = capelPat + " " + capelMat
      cTarjeta = rs!tarjeta
      CmdRemuHep.Caption = cPatron
      If UCase(rs!apelmat) = "D'BROT" Then
      Else
        Set rs_d = New Recordset
        sql = "SELECT * FROM dat_nomb WHERE Nombre LIKE '" & cPatron & "%'  ORDER BY  nombre"
        rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not rs_d.EOF
           If InStr(rs_d!nombre, cPatron) > 0 Then
              If InStr(rs_d!nombre, cnombres) > 0 Then
                 nremunera = 0: clibele = "": cCond = "": ctpo = "": cCodniv = "": Chor = "": dFecnac = "": dFecalt = ""
                 For i = 0 To rs_d.Fields.Count - 1
                   On Error Resume Next
                   ccampo = UCase(rs_d.Fields(i).Name)
                   If Left(ccampo, 2) = "C1" Then
                      If ccampo = "C1047" Then
                         'ndato = IIf(IsNull(rs.Fields(i)), 0, rs.Fields(i))
                         'nguardia = nguardia + ndato
                      ElseIf InStr("C1049_C1053_C1054_C1056_C1064_C1160", ccampo) > 0 Then
                             'ndato = IIf(IsNull(rs.Fields(i)), 0, rs.Fields(i))
                             'nbonocasion = nbonocasion + ndato
                      Else
                         ndato = IIf(IsNull(rs_d.Fields(i)), 0, rs_d.Fields(i))
                         nremunera = nremunera + ndato
                      End If
                   Else
                      If ccampo = "LIBELE" Then
                         clibele = rs_d.Fields(i)
                      ElseIf ccampo = "COND" Then
                             cCond = rs_d.Fields(i)
                      ElseIf ccampo = "TPO" Then
                             ctpo = rs_d.Fields(i)
                      ElseIf ccampo = "HOR" Then
                             Chor = rs_d.Fields(i)
                      ElseIf ccampo = "CODNIV" Then
                             cCodniv = rs_d.Fields(i)
                      ElseIf ccampo = "FECNAC" Then
                             dFecnac = rs_d.Fields(i)
                      ElseIf ccampo = "FECALT" Then
                             dFecalt = rs_d.Fields(i)
                      End If
                   End If
                 Next i
                 sql = "UPDATE CpGeneral  SET documento='" & clibele & "', plh='" & cCond & "', tipoPla='" & ctpo & "', codniv='" & cCodniv & "', grpGen='" & Chor & "', RemTot=" & nremunera & " WHERE  (baja is NULL) and tarjeta='" & cTarjeta & "'"
                 Cnx.Execute sql
                 nson = nson + 1
                 CmdRemuHep.Caption = "Van : " + CStr(nson)
                 Exit Do
              End If
           End If
           rs_d.MoveNext
        Loop
        rs_d.Close
      End If
      rs.MoveNext
   Loop
   rs.Close
   CmdRemuHep.Caption = "A C T U A L I Z A  HEP->RemTot->General"
   MsgBox ("Se actualizo " & CStr(nson) & " de " & CStr(nvan))
   Call cmdsalir_Click
End Sub


Private Sub Cmdfecha_Click()
   If Val(Txmes.Text) = 0 Or Val(Txmes.Text) > 12 Then Exit Sub
   unafecha = "01/" + Txmes.Text + "/" + Trim(Str(genaño))
   Call ObtenerPrimerUltimoDia(unafecha)
   dfi = Day(primer)
   dff = Day(ultimo)
   For i = dfi To dff
       z = i - 1
       Lblfecha.Caption = unafecha + i - 1
       Lbldato.Caption = Trim(WeekdayName(Weekday(primer + z))) + " " & i
       If z > 5 Then
          Exit For
       Else
          MsgBox ("press any key")
       End If
   Next
   Lblfecha.Caption = ""
   Lbldato.Caption = ""
   Txmes.Text = ""
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
         TxPeriodo.Text = Right(CboMes.List(j), nlen - nposicion)
      End If
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
            TxPeriodo.Text = Right(CboMes.List(j), nlen - nposicion)
         End If
      End If
   End If
End Sub
Private Sub CmdChkDir_Click()
   Set rs = New Recordset
   rs.Open "SELECT * FROM cpDirectorio WHERE ano='" & genaño & "' And Mes='" & TxPeriodo.Text & "' And flag='0'", Cnx, adOpenStatic, adLockReadOnly
   nhay = 0: nson = 0: nvan = 0
   Do While Not rs.EOF
      cTarjeta = rs!tarjeta
      nvan = nvan + 1
      unafecha = "01/" + TxPeriodo.Text + "/" + Trim(Str(genaño))
      Call ObtenerPrimerUltimoDia(unafecha)
      nhoras = 0: nd = 0: nn = 0: nm = 0: nt = 0
      For i = Day(primer) To Day(ultimo)
         z = i - 1
         cTurno = Trim(IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4)))
         If cTurno = "D" Or cTurno = "D*" Or cTurno = "1" Or cTurno = "D1" Then
            nhoras = nhoras + 12
            nd = nd + 1
         ElseIf cTurno = "N" Or cTurno = "N*" Or cTurno = "4" Then
                nhoras = nhoras + 12
                nn = nn + 1
         ElseIf cTurno = "M" Or cTurno = "M1" Then
                nhoras = nhoras + 6
                nm = nm + 1
         ElseIf cTurno = "T" Or cTurno = "T1" Then
                nhoras = nhoras + 6
                nt = nt + 1
         End If
      Next
      n_horas = IIf(IsNull(rs!Horas), 0, rs!Horas)
      c_d = IIf(IsNull(rs!d), "", rs!d)
      c_n = IIf(IsNull(rs!n), "", rs!n)
      c_m = IIf(IsNull(rs!m), "", rs!m)
      c_t = IIf(IsNull(rs!T), "", rs!T)
      If nhoras <> n_horas Or nd <> Val(c_d) Or nn <> Val(c_n) Or nm <> Val(c_m) Or nt <> Val(c_t) Then
         If nd <> Val(c_d) Then
            sql = "UPDATE CpDirectorio  SET horas=" & nhoras & ", d='" & Trim(Str(nd)) & "' WHERE tarjeta='" & cTarjeta & "' and mes='" & TxPeriodo.Text & "' and ano='" & genaño & "'"
            Cnx.Execute sql
         End If
         If nn <> Val(c_n) Then
            sql = "UPDATE CpDirectorio  SET horas=" & nhoras & ", n='" & Trim(Str(nn)) & "' WHERE tarjeta='" & cTarjeta & "' and mes='" & TxPeriodo.Text & "' and ano='" & genaño & "'"
            Cnx.Execute sql
         End If
         If nm <> Val(c_m) Then
            sql = "UPDATE CpDirectorio  SET horas=" & nhoras & ", m='" & Trim(Str(nm)) & "' WHERE tarjeta='" & cTarjeta & "' and mes='" & TxPeriodo.Text & "' and ano='" & genaño & "'"
            Cnx.Execute sql
         End If
         If nt <> Val(c_t) Then
            sql = "UPDATE CpDirectorio  SET horas=" & nhoras & ", t='" & Trim(Str(nt)) & "' WHERE tarjeta='" & cTarjeta & "' and mes='" & TxPeriodo.Text & "' and ano='" & genaño & "'"
            Cnx.Execute sql
         End If
         CmdChkDir.Caption = cTarjeta + ":" + Str(nd) + "/" + Str(nn) + "/" + Str(nm) + "/" + Str(nt)
         nson = nson + 1
      End If
      rs.MoveNext
   Loop
   rs.Close
   CmdChkDir.Caption = "Check  Directorio"
   MsgBox ("Se Revizado " & CStr(nson) & " de " & CStr(nvan))
   Call cmdsalir_Click
   
End Sub

Private Sub cmdsalir_Click()
  Set rs_d = Nothing
  Set rs = Nothing
  Me.Hide
  Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set rs_d = Nothing
  Set rs = Nothing
  Me.Hide
  Unload Me
End Sub



     
