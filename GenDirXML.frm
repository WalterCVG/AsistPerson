VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form GenDirXML 
   Caption         =   "Directorio Mensual XML"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "GenDirXML.frx":0000
   ScaleHeight     =   2850
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txmesdeano 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   2160
      TabIndex        =   12
      Top             =   1200
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
      Left            =   720
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   -960
      Picture         =   "GenDirXML.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   195
      Left            =   840
      MaxLength       =   2
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   7680
      Picture         =   "GenDirXML.frx":80EA8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cerrar"
      Top             =   2040
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton CmdExporXML 
         Caption         =   "Exporta XML"
         Height          =   735
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdcarga 
         Height          =   1095
         Left            =   120
         Picture         =   "GenDirXML.frx":812EA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ubicar Registros"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdClear 
         Height          =   735
         Left            =   2400
         Picture         =   "GenDirXML.frx":814C5
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Otra Consulta"
         Top             =   360
         Width           =   735
      End
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
      Left            =   6120
      Picture         =   "GenDirXML.frx":81D8F
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir"
      Top             =   1800
      Width           =   765
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton mnuFicImpSelec 
      Height          =   495
      Left            =   5040
      Picture         =   "GenDirXML.frx":82659
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Selec.Impresora"
      Top             =   1800
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "GenDirXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim r_d As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim Sw As Boolean
Dim nposicion, nvalor As Integer
Dim sql, SQ1, cTarjeta, cDni, cvalor, cTurno, cdia As String
Dim unafecha As Date
Private Sub Command1_Click()
     c_grupo = "05": c_secuencia = "057"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub
Private Sub Form_Load()
  TxMesActual = Trim(Mid(wMeses, (Month(Date) - 1) * 9 + 1, 9))
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
            unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)            'primer/ultimo
            Txfecini.Text = primer
            TxFecFin.Text = ultimo
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
   On Error Resume Next
   Cnx.Execute "DELETE  FROM CpTmDirXML"
   Set rt = New Recordset
   SQ1 = "SELECT * FROM CpTmDIRXML "
   rt.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   '-carga  DIRECTORIO
   sql = "SELECT  * FROM CpDirectorio  WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "' and Flag='0' and ((ccosto >='30' and ccosto<='32') or ccosto='24')" & _
        "ORDER BY ccosto, codupr, Tarjeta"
   Set r_d = New Recordset
   r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   nvalor = 0
   Do While Not r_d.EOF
      cDni = "": caPat = "": caMat = "": cnomb = ""
      cCodUpr = "": cDscUpr = "": cCarest = "": cdscCarEst = ""
      cTarjeta = r_d!tarjeta
      Sw = False
      '-Coge Generales del trabajador
      sql = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & cTarjeta & "'"
      Set rs = New Recordset
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then
         caPat = Trim(rs!apelpat)
         caMat = Trim(rs!apelmat)
         cnomb = Trim(rs!nombres)
         cDni = IIf(IsNull(rs!documento), "", rs!documento)
         cCodUpr = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)
         cCarest = IIf(IsNull(rs!carEst), "", rs!carEst)
      End If
      rs.Close
      If cCodUpr <> "" Then
         Set rs_d = New Recordset
         rs_d.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs_d.EOF Then
            cDscUpr = Trim(rs_d!descripcion)
         End If
         rs_d.Close
      End If
      If cCarest <> "" Then
         Set rs_d = New Recordset
         rs_d.Open "SELECT * FROM Cptablas WHERE tabla='27' and codigo='" & cCarest & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs_d.EOF Then
            cdscCarEst = Trim(rs_d!descripcion)
         End If
         rs_d.Close
      End If
      '-Procesa Directorio Trabajador
      For i = Day(primer) To Day(ultimo)
         z = i - 1
         cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", Trim(r_d.Fields(z + 4)))
         cmdcarga.Caption = cTarjeta + " " + Str(i)
         If InStr("D*D1D2D3D4_N*N2N3N4_H*H1_C*C1_P*P1_DNMTMXHCP", Trim(cTurno)) > 0 And (Trim(cTurno) <> "" And Trim(cTurno) <> "*") Then
            fechaB = Format(primer + z, "yyyymmdd")
            p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = ""
            '-HORARIO prog en dia
            sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
            Set rs = New Recordset
            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs.EOF Then
               cvalor = IIf(IsNull(rs!Horario), "", rs!Horario)
               Sq0 = "SELECT * FROM CpHorarioNew  WHERE Codigo='" & cvalor & "'"
               Set rs_d = New Recordset
               rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not rs_d.EOF Then
                  p_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
                  p_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                  p_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
                  p_sal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
               End If
               rs_d.Close
            End If
            rs.Close
            If Not Sw Then
               nvalor = nvalor + 1
               Sw = True
            End If
            rt.AddNew
            rt!Servicio = Left(cDscUpr, 50)
            rt!fecha = primer + z
            cdia = Left(WeekdayName(Weekday(primer + z)), 3)
            rt!Dia = UCase(Left(cdia, 1)) + Right(cdia, 2)
            rt!apellidos = Left(caPat + " " + caMat, 30)
            rt!nombres = Left(cnomb, 25)
            rt!Especialidad = Left(cdscCarEst, 15)
            If InStr("DNMT_MX_NA_D*D1D2D3D4_N*N2N3N4", Trim(cTurno)) > 0 Then
               rt!Horario = p_ing + "-" + p_sal
            ElseIf InStr("H*H1_C*C1_P*P1_PA_HCP", Trim(cTurno)) > 0 Then
                  rt!Horario = p_ing + "-" + p_sal + " / " + p_ing1 + "-" + p_sal1
            End If
            rt.Update
         End If
      Next i
      r_d.MoveNext
   Loop
   r_d.Close
   rt.Close
   cmdcarga.Caption = ""
   If nvalor > 0 Then
      MsgBox "Total Servidores " + Trim(Str(nvalor)), vbExclamation
      CmdPrint.SetFocus
   End If
   Exit Sub
   'Call Limpiacampos
End Sub

Private Sub CmdExporXML_Click()
    wTpoQuery = "5.7"
    Load ExporXML
    ExporXML.Left = 0: ExporXML.Top = 0
    ExporXML.Show
    Exit Sub
End Sub

Private Sub CmdPrint_Click()
   If Txmes.Text = "" Then
      MsgBox "Defina..  Periodo... ", vbExclamation
      Exit Sub
   End If
   nTipRpt = 17
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Call Limpiacampos
   CmdPrint.Caption = ""
   Exit Sub
End Sub
Private Sub Limpiacampos()
   CboMes.Text = "":   Txmes.Text = "": TxNombMes.Text = ""
   Txfecini.Text = "": TxFecFin.Text = ""
End Sub
Private Sub CmdClear_Click()
   Call Limpiacampos
End Sub
Private Sub cmdsalir_Click()
   Set r_d = Nothing
   Set rs = Nothing
   Set rt = Nothing
   wTpoQuery = ""
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "GenDirXML" Then
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
   Set rt = Nothing
   wTpoQuery = ""
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "GenDirXML" Then
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


