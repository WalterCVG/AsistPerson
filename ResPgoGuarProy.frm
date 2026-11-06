VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ResPgoGuarProy 
   Caption         =   "Resolucion - Pago de Guardias (Real proyectado)"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "ResPgoGuarProy.frx":0000
   ScaleHeight     =   2820
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   8760
      Picture         =   "ResPgoGuarProy.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Cerrar"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox TxFInicio 
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
      Left            =   2760
      TabIndex        =   25
      Top             =   600
      Width           =   1095
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
      Left            =   4440
      TabIndex        =   24
      Top             =   600
      Width           =   1095
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
      Height          =   195
      Left            =   1440
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
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
      Height          =   195
      Left            =   1440
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Txtarjeta 
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      MaxLength       =   5
      TabIndex        =   20
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Resolución de Pago Planilla"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox TxNroResolu 
      Height          =   285
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxMesAno 
      Enabled         =   0   'False
      Height          =   195
      Left            =   3360
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CmdExportaTXT 
      Height          =   615
      Left            =   7680
      Picture         =   "ResPgoGuarProy.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exporta a Txt"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxMesdeAno 
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   3135
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
      Height          =   195
      Left            =   1320
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox TxMesActual 
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   1095
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
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   5880
      TabIndex        =   11
      Top             =   1320
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "Alfabetico"
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   120
         Width           =   1095
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
         Left            =   1680
         Picture         =   "ResPgoGuarProy.frx":80FF2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir"
         Top             =   480
         Width           =   645
      End
      Begin VB.CommandButton mnuFicImpSelec 
         Height          =   675
         Left            =   360
         Picture         =   "ResPgoGuarProy.frx":818BC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Selec.Impresora"
         Top             =   480
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   1095
      Left            =   6480
      Picture         =   "ResPgoGuarProy.frx":8209A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ubicar Registros"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   8760
      Picture         =   "ResPgoGuarProy.frx":82275
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Otra Consulta"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "ResPgoGuarProy.frx":82B3F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   195
      Left            =   2640
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1935
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
      Left            =   4080
      TabIndex        =   27
      Top             =   600
      Width           =   375
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
      Left            =   2160
      TabIndex        =   26
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. Resolución"
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "ResPgoGuarProy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim r_d As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim rt0 As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim f_Fer, F_Existe, F_Modif As Boolean
Dim ntotpresu As Double
Dim nDiaI, nDiaF As Integer
Dim nLacta, nSonLacta, nposicion, nvalor As Integer
Dim ciFALTA, csexo, cAplica, cTarjeta, cDni, sql, SQ1, c_a, cplh, cTurno, cGrpGen, c_Costo, cDscCosto, cCodUpr, cDscUpr, ccadhep, cDcadhep, ccarpla, cDcarpla, cnivGuardia, ccadena, cPatron, wMuestra As String
Dim unafecha As Date
Private Sub Command1_Click()
    c_grupo = "04": c_secuencia = "043"
    Load AydAcerca
    AydAcerca.Left = 0: AydAcerca.Top = 0
    AydAcerca.Show
End Sub
Private Sub Form_Load()
  '-Pago Guardias HEP/RES
  TxMesActual = Trim(Mid(wMeses, (Month(Date) - 1) * 9 + 1, 9))

  TxFecActual = "La Victoria, " + Right$("00" + Trim(Str(Day(Date))), 2) + " de " + TxMesActual + " de " & Year(Date)

   Check2.Value = 0
   Check2.Enabled = True      'Posicion FINAL base pago PLANILLA
   
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
  F_Existe = False:  F_Modif = False
  TxNroResolu.Locked = False
End Sub
Private Sub Cbomes_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboMes.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbomes_LostFocus()
    Call Cbomes_KeyPress(13)
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
            Txmes = Right(CboMes.List(j), nlen - nposicion)
            unafecha = "01/" + Trim(Txmes) + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)            'primer/ultimo
            TxNombMes = Left(CboMes.List(j), nposicion - 1)
            TxMesdeAno = TxNombMes + " del " & genaño
            TxMesAno = TxNombMes + " " & genaño
            TxFecIni = primer
            nDiaI = Day(primer)
            TxFecFin = ultimo
            nDiaF = Day(ultimo)
         End If
         F_Existe = False: F_Modif = False
           '''TxNroResolu.Locked = True
      End If
      Exit Sub
   End If
End Sub

Private Sub txfInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If IsDate(TxFInicio) Then
        CboMes.Text = "": Txmes = ""
        TxFFin = TxFInicio
        TxFFin.SetFocus
        Exit Sub
     End If
  End If
End Sub
Private Sub txfFin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     CboMes.Text = "": Txmes = ""
     If IsDate(TxFFin) Then
        If Month(DateValue(TxFInicio)) <> Month(DateValue(TxFFin)) Then
           MsgBox "fechas no corresponden al mes. ", vbExclamation
           Exit Sub
        End If
        If Year(DateValue(TxFInicio)) <> Year(DateValue(TxFFin)) Then
           MsgBox "fechas de años diferentes.. ", vbExclamation
           Exit Sub
        End If
        If Year(DateValue(TxFFin)) <> Val(genaño) Then
           MsgBox "fecha no corresponde al periodo..", vbExclamation
           Exit Sub
        End If
        TxFecIni = TxFInicio
        nDiaI = Day(DateValue(TxFInicio))
        TxFecFin = TxFFin
        nDiaF = Day(DateValue(TxFFin))
        Txmes = Right("00" + Trim(Str(Month(DateValue(TxFFin)))), 2)
        
        unafecha = "01/" + Trim(Txmes) + "/" + Trim(Str(genaño))
        Call ObtenerPrimerUltimoDia(unafecha)            'primer/ultimo
        
        TxNombMes = Trim(Mid(wMeses, (Val(Txmes) - 1) * 9 + 1, 9))
        TxMesdeAno = TxNombMes + " del " & genaño
        TxMesAno = TxNombMes + " " & genaño
        F_Existe = False
        F_Modif = False
     End If
  End If
End Sub
Private Sub TxFFin_LostFocus()
   Call txfFin_KeyPress(13)
End Sub


Private Sub Check2_Click()
    F_Existe = False
    F_Modif = False
    TxNroResolu.Locked = False
    If Check2.Value = 1 Then
       '-Verifica si existe Nro Resolucion de "PAGO"  -*
       sql = "SELECT  * FROM CpDirectorio  WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes) & "' and nroResolu<>''"
       Set r_d = New Recordset
       r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       If Not r_d.EOF Then
          TxNroResolu = Trim(r_d!nroresolu)
          TxNroResolu.Locked = True
          F_Existe = True
          If MsgBox("Ya Existe, Resolución de Pago Planillas... " + TxNroResolu & ", reemplaza?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
             F_Modif = True
          End If
          Call cmdcarga_Click
       End If
    End If
End Sub

Private Sub cmdcarga_Click()
   If Txmes = "" And (TxFInicio = "" Or TxFFin = "") Then
      MsgBox "No indico periodo.. ", vbExclamation
      CboMes.SetFocus
      Exit Sub
   End If
       
   If CboMes.Text <> "" Then
      Call Cbomes_KeyPress(13)
   End If
   
   If Check2.Value = 1 And TxNroResolu = "" Then
      MsgBox "Favor, indique Nro. de Resolución.. ", vbCritical
      TxNroResolu.SetFocus
      Exit Sub
   End If
     
   wPer_resol = Txmes + Right(Trim(Str(genaño)), 2)
   cperiodo = "1047" + Txmes + Right(Trim(Str(genaño)), 2)
   Cnx.Execute "DELETE  FROM CpPG1047"
   Set rt0 = New Recordset
   SQ1 = "SELECT * FROM CpPG1047"
   rt0.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
     
   If Check2.Value = 1 Then        'Resolucion Base Pago Planilla
      If F_Existe Then
            '-Muestra  Resolucion Existente
            '------------------------------
            Cnx.Execute "DELETE  FROM CpTmPagoGuar"
            Set rt = New Recordset
            SQ1 = "SELECT * FROM CpTmPagoGuar"
            rt.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
            
            sql = "SELECT  * FROM CpDirectorio  WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes) & "' and nroResolu='" & TxNroResolu & "'"
            Set r_d = New Recordset
            r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            nvalor = 0
            ntotpresu = 0#
            Do While Not r_d.EOF
               cmdcarga.Caption = cTarjeta
               c_Costo = "": cDscCosto = "": cCodUpr = "": cDscUpr = ""
               cDni = "": cplh = "": caPat = "": caMat = "": cnomb = "": cnivGuardia = "": cGrpGen = "": ccadhep = "": cDcadhep = ""
               ccarpla = "": cDcarpla = ""
               c_Costo = IIf(IsNull(r_d!CCosto), "", r_d!CCosto)
               cCodUpr = IIf(IsNull(r_d!CodUpr), "", r_d!CodUpr)
               cTarjeta = r_d!TARJETA
               ccadhep = r_d!CadHep
               Set rs = New Recordset
               sql = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & cTarjeta & "'"
               rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If rs.EOF Then
               Else
                 cplh = IIf(IsNull(rs!PLH), "", rs!PLH)                  'guardias  HEP/RES
                 cDni = IIf(IsNull(rs!Documento), "", rs!Documento)
                 caPat = Trim(rs!apelpat)
                 caMat = Trim(rs!apelmat)
                 cnomb = Trim(rs!nombres)
                 cnivGuardia = IIf(IsNull(rs!NivGuardia), "", rs!NivGuardia)
                 ccarpla = IIf(IsNull(rs!carPla), "", rs!carPla)
                 If ccadhep = "" Then
                   ccadhep = IIf(IsNull(rs!CadHep), "", rs!CadHep)
                 End If
               End If
               rs.Close
               
               If ccarpla <> "" Then
                  Set rs = New Recordset
                  rs.Open "SELECT * FROM Cptablas WHERE tabla='28' and codigo='" & ccarpla & "'", Cnx, adOpenStatic, adLockReadOnly
                  If Not rs.EOF Then
                     cDcarpla = Trim(rs!descripcion)
                  End If
                  rs.Close
               End If
               nTarDO = 0#: nTarNO = 0#: nTarDF = 0#: nTarNF = 0#
               If cniv <> "" Then
                  Set rs = New Recordset
                  rs.Open "SELECT * FROM CpTarifas WHERE (baja is NULL)", Cnx, adOpenStatic, adLockReadOnly
                  Do While Not rs.EOF
                     '-Tarifa aplicado al servidor
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
               If c_Costo <> "" Then
                  Set rs = New Recordset
                  rs.Open "SELECT * FROM CpcCosto WHERE codigo='" & c_Costo & "'", Cnx, adOpenStatic, adLockReadOnly
                  If Not rs.EOF Then
                     cDscCosto = Trim(rs!descripcion)
                  End If
                  rs.Close
               End If
               If cCodUpr <> "" Then
                  Set rs = New Recordset
                  rs.Open "SELECT * FROM CpUndPrgRol WHERE codUpr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
                  If Not rs.EOF Then
                     cDscUpr = Trim(rs!descripcion)
                  End If
                  rs.Close
               End If
               If ccadhep <> "" Then
                  Set rs = New Recordset
                  rs.Open "SELECT * FROM Cpcadena WHERE ano='" & genaño & "' and cadena='" & ccadhep & "'", Cnx, adOpenStatic, adLockReadOnly
                  If Not rs.EOF Then
                     cDcadhep = Trim(rs!descripcion)
                  End If
                  rs.Close
               End If
               
               If rD!importe > 0 Then
                  cvalor = Format(rD!importe, "######.00")
                  rt0.AddNew
                  rt0!codigo = cperiodo
                  rt0!nombre = Left(caPat + " " + caMat + " " + cnomb, 45)
                  rt0!libele = Left(cDni, 8)
                  nlen = Len(cvalor)
                  If rD!importe > 0 Then
                     rt0!Monto = Left(cvalor, nlen - 3) + Right(cvalor, 2)
                  Else
                     rt0!Monto = "00"
                  End If
                  rt0.Update
               End If
               
               rt.AddNew
               rt!CCosto = c_Costo
               rt!Dsccosto = cDscCosto
               rt!CodUpr = cCodUpr
               rt!dscUPR = cDscUpr
               rt!CadHep = ccadhep
               rt!Dsccadhep = cDcadhep
               rt!TARJETA = cTarjeta
               rt!apelpat = caPat
               rt!apelmat = caMat
               rt!nombres = cnomb
               rt!Dni = cDni
               rt!NivGuardia = cnivGuardia
               rt!DsccarPla = cDcarpla
               rt!TarifaDO = nTarDO
               rt!TarifaNO = nTarNO
               rt!TarifaDF = nTarDF
               rt!TarifaNF = nTarNF
               rt!D_O = r_d!D_O
               rt!N_O = r_d!N_O
               rt!D_F = r_d!D_F
               rt!N_F = r_d!N_F
               rt!Monto = r_d!importe
               rt.Update
               ntotpresu = ntotpresu + r_d!importe
               nvalor = nvalor + 1
               r_d.MoveNext
            Loop
            r_d.Close
            cmdcarga.Caption = ""
            MsgBox "Total Servidores " + Trim(Str(nvalor)) + "  Monto S/." + Format(ntotpresu, "#######0.00"), vbExclamation
            Call CmdPrint_Click
            rt.Close
            If Not F_Modif Then
               Exit Sub
            End If
      End If
   End If
   
   Cnx.Execute "DELETE  FROM CpTmPagoGuar"
   Set rt = New Recordset
   SQ1 = "SELECT * FROM CpTmPagoGuar"
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
            If Len(Trim(rs!codigo)) = 6 Then       'codigo:mmddaa
               c_a = Right(rs!codigo, 2)
            ElseIf Len(Trim(rs!codigo)) = 4 Then   'Codigo:mmdd
                   c_a = Right(Year(unafecha), 2)
            End If
            If Val(Left(rs!codigo, 2)) = Val(Txmes) And Val(Mid(rs!codigo, 3, 2)) = i And Right(Year(unafecha), 2) = c_a Then
               f_Fer = True
            End If
            rs.MoveNext
         Wend
         rs.Close
      End If
      ccadena = ccadena + IIf(f_Fer, "F", "O")
   Next i
   
      '-MsgBox Format(unafecha, "dd/mm/yyyy") + ccadena
      
   '-carga  DIRECTORIO "TURNO ASISTENCIALES"
   sql = "SELECT  * FROM CpDirectorio  WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes) & "' and flag='0' " & _
        "ORDER BY Tarjeta, ano, mes "
   Set r_d = New Recordset
   r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not r_d.EOF Then
     nvalor = 0: nposicion = 0
     ntotpresu = 0#
     wMuestra = "N"
     Do While Not r_d.EOF
        c_Costo = "": cDscCosto = "": cCodUpr = "": cDscUpr = ""
        cDni = "": cplh = "": caPat = "": caMat = "": cnomb = "": cnivGuardia = "": cGrpGen = "": ccadhep = "": cDcadhep = ""
        ccarpla = "": cDcarpla = ""
        cTarjeta = r_d!TARJETA
        c_Costo = IIf(IsNull(r_d!CCosto), "", r_d!CCosto)
        cCodUpr = IIf(IsNull(r_d!CodUpr), "", r_d!CodUpr)
          '-SABER si una columna "cadhep" existe en tabla "CpDirectorio"
          SQ1 = "SELECT * FROM  INFORMATION_SCHEMA.COLUMNS AS c1  WHERE  c1.column_name = 'cadhep'  AND c1.table_name = 'cpdirectorio'"
          Set rs_2 = New Recordset
          rs_2.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
          If Not rs_2.EOF Then
             ccadhep = IIf(IsNull(r_d!CadHep), "", Trim(r_d!CadHep))
          End If
          rs_2.Close
      
        Set rs = New Recordset
        sql = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & cTarjeta & "'"
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If rs.EOF Then
        Else
           cplh = IIf(IsNull(rs!PLH), "", rs!PLH)                  'guardias  HEP/RES
           cDni = IIf(IsNull(rs!Documento), "", rs!Documento)
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
               rs_d.Open "SELECT * FROM CpcCosto WHERE codigo='" & c_Costo & "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rs_d.EOF Then
                  cDscCosto = Trim(rs_d!descripcion)
               End If
               rs_d.Close
           End If
           If cCodUpr <> "" Then
               Set rs_d = New Recordset
               rs_d.Open "SELECT * FROM CpUndPrgRol WHERE codUpr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
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
              rs.Open "SELECT * FROM CpTarifas  WHERE (baja is NULL)", Cnx, adOpenStatic, adLockReadOnly
              Do While Not rs.EOF
                 '-Tarifa aplicado al servidor
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
           '-Procesa DIRECTORIO del servidor
           For i = nDiaI To nDiaF          '--- Day(primer) To Day(ultimo)
              z = i - 1
              cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", Trim(r_d.Fields(z + 4)))
              cmdcarga.Caption = cTarjeta + " " + Str(i)
              '-Procesa Guardias
              If InStr("1DN4HCP", Left(cTurno, 1)) > 0 And cTurno <> "*" And Trim(cTurno) <> "" Then
                 F_falta = False
                 '-Verifica Asistencia Y/o Autorizaciones "D" que afectan Guardia
                 fechaB = Format(DateValue(primer) + z, "yyyymmdd")
                 dfentra = DateValue(primer) + z
                 dfsale = DateValue(primer) + z
                 p0_ing = "": p0_sal = "": p0_ing1 = "": p0_sal1 = ""   'HORARIO ORIGINAL
                 p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = ""
                 m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
                 nValorH = 0: nValorM = 0: mmintarde = 0:  mdscto = 0: nFaltaH = 0: nFaltaM = 0: nmotpersonal = 0: nComision = 0: ciFALTA = "": F_falta = False
                 cvalor = "": cAuPrg = "": cDiahora = "": cTipo = "": choraini = "": chorafin = "": cDescuenta = "": c_a = ""
                 cDiahora = "X"
                 cMantGuardia = "N"
                 
                 '-Afecto X LACTANCIA
                 '=============================
                 cAplica = "N": nSonLacta = 0
                 TxIniLac = "": TxFinLac = ""
                 If csexo = "F" Then
                    '-Ubica lactancia al periodo
                    Set rs = New Recordset
                    sql = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "' order by fecnac"
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
                 
                 '--DIRECTORIO REAL  y proyecta; Caso date > inicio MES
                 '-AUTORIZACIONES: Motivos Personales, Comisiones Serv,  V/L/R  etc...
                 sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and   tarjeta='" & cTarjeta & "' and FechaIni<='" + fechaB + "' and FechaFin>='" + fechaB + "'"
                 Set rs = New Recordset
                 rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                 Do While Not rs.EOF
                    If Not rs.EOF Then
                        cAuPrg = Trim(rs!codigo)
                        cDiahora = rs!diahora
                        cTipo = IIf(IsNull(rs!tipo), "", rs!tipo)            'I/S/O   caso horas
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
                        If cTipo = "I" Then
                            m_ing = choraini
                        ElseIf cTipo = "S" Then
                               m_sal = chorafin
                        End If
                    End If
                    rs.MoveNext
                 Loop
                 '-HORARIO prog en Dia -**
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
                          If rs_d!Horas < "1800" Then                              '========>
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
                          c2dias = "N"
                          If (p0_ing > p0_sal And Val(p0_ing1) = 0) Then         '======>Or p_horas = "2400"
                             c2dias = "S"        '1900-0700
                          ElseIf Val(p0_sal1) > 0 And p0_ing < p0_ing1 Then
                             c2dias = "S"   '1300-1900   1900-0700   0700-1200 1900-0700
                          End If
                          
                       End If
                       rs_d.Close
                    End If
                 End If
                 rs.Close
                    
                 If (cDiahora = "H" Or cDiahora = "D") And cMantGuardia = "N" Then
                    'Pierde Guardia
                    F_falta = True
                    ciFALTA = "a"
                    'If wMuestra = "S" Then
                    '   MsgBox (Format(primer + z, "dd/mm/yyyy") + "1")
                    'End If
                 Else
                    '-MARCAS del dia programado      fecha  hora =>  hhmm
                    If Val(p_horas) > 0 Then       '<======And primer + z < Date
                       If DateValue(primer) + z < Date Then   '<=Inicio: Evalua MARCAS hasta AYER
                          Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
                          If Val(p_sal1) > 0 Or Val(p_sal) < Val(p_ing) Then
                             '-dia siguiente  turno t/n, n/d
                             fecha1 = Format(DateValue(primer) + z + 1, "yyyymmdd")
                             Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha >='" & fechaB & "' AND fecha <='" & fecha1 & "' ORDER BY fecha,hora"
                             c2dias = "S"
                          End If
                          
                          Set rs_d = New Recordset
                          rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                          If Not rs_d.EOF Then
                             '-ASIGNA marcas basado en Horario
                             d_fentra = rs_d!fecha        '========> 1era marca del servidor para turno programado
                             Do While Not rs_d.EOF
                                cInicio = p_ing
                                cFin = rs_d!hora
                                dfsale = rs_d!fecha
                                If Val(m_ing) = 0 Then
                                   'if (nTotMints(cInicio, cFin) >= -60 And nTotMints(cInicio, cFin) <= 60) And Day(dfsale) = i Then
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
                       End If        '<=Fin verifica MARCAS antes de Hoy
                    End If
                       
                    '-Evalua  guardia Programada
                    If Val(p_horas) > 0 Then
                       '-Evalua MARCAS Ingreso
                       If (cDiahora = "D") And cDescuenta = "S" Then
                          F_falta = True
                          ciFALTA = "b"
                          'If wMuestra = "S" Then
                          '   MsgBox (Format(DateValue(Txfecini)  + z, "dd/mm/yyyy") + "2")
                          'End If
                          
                       'ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And p_sal <> p_ing1 Then      '======>
                       ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And Val(p_ing1) > 0 And p_sal <> p_ing1 Then  '======>
                              'Caso Horario:2000-0800   marca 0805-2010   O marca 0749-1404 1930-0820 ó  0803-2015/-0824
                              F_falta = True
                              ciFALTA = "c"

                       ElseIf (Val(m_ing) + Val(m_ing1) = 0) And DateValue(TxFecIni) + z < Date Then       '<======And primer + z < Date
                              '-No marco Ing,   valida vs autoriza
                              If cDiahora = "H" And cTipo = "I" Then
                                 cInicio = p_ing
                                 cFin = choraini
                                 If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                                    F_falta = True
                                    ciFALTA = "d"
                                    cDescuenta = "S"
                                    'If wMuestra = "S" Then
                                    '   MsgBox (Format(primer  + z, "dd/mm/yyyy") + "3")
                                    'End If
                                 End If
                              ElseIf cDiahora = "D" Then    'valida autoriza Dia  pierde
                                    If cDescuenta = "S" Then
                                       F_falta = True
                                       ciFALTA = "e"
                                       'If wMuestra = "S" Then
                                       '   MsgBox (Format(primer + z, "dd/mm/yyyy") + "4")
                                       'End If
                                       
                                    End If
                              Else
                                  F_falta = True
                                   ciFALTA = "f"
                                  'If wMuestra = "S" Then
                                  '   MsgBox (Format(primer  + z, "dd/mm/yyyy") + "5")
                                  'End If
                              End If
                       ElseIf (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) And DateValue(TxFecIni) + z < Date Then   '<====And primer + z < Date
                             'Marco tarde  Autoriza H al ingreso
                             If cDiahora = "H" And cTipo = "I" Then
                                cInicio = p_ing
                                cFin = choraini
                                If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                                   F_falta = True
                                   ciFALTA = "g"
                                    'If wMuestra = "S" Then
                                     '  MsgBox (Format(primer  + z, "dd/mm/yyyy") + "6")
                                     'End If
                                End If
                             End If
                       ElseIf Val(m_ing) > 0 And DateValue(primer) + z < Date Then     '<====And primer + z < Date
                              '--- INSERTADO Excepciones ----09.08.17
                              If Val(m_ing1) = 0 Then
                                 '  PI   PS            MI   MS     MI1  MS1
                                 '- 2000-0800      :   0740-1401   FALTA
                                 '                 :   2130-0835   FALTA no marco salida  ¿?
                                 '-                :   1750-0820   ok
                                 '--               :   0741-2044             marcas fuera Turno (antes/despues)-> FALTA
                                 '---  0700-1900 1900-0700  24Hras  :marcas   0715- / -0748
                                 If (m_sal > m_ing And m_sal < p_ing) And p_horas <> 2400 Then  '¿2400?
                                    '-salida Antes de Turno
                                    F_falta = True
                                    ciFALTA = "h"
                                    'If wMuestra = "S" Then
                                    '    MsgBox (Format(DateValue(Txfecini)  + z, "dd/mm/yyyy") + "7")
                                    'End If
                                    
                                 Else
                                    If c2dias = "S" And ((m_sal > p_sal And d_fentra = dfsale) Or (m_sal < p_sal And d_fentra < dfsale)) Then
                                       '-No completo el turno
                                       F_falta = True
                                       ciFALTA = "i"
                                       
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
                                        ciFALTA = "j"
                                       'If wMuestra = "S" Then
                                       '   MsgBox (Format(DateValue(Txfecini)  + z, "dd/mm/yyyy") + "8")
                                       'End If
                                        
                                    End If
                                 End If
                              End If
                       End If
                       '-Evalua MARCAS Salida
                       If F_ALTA Then
                       ElseIf (Val(m_sal) + Val(m_sal1) = 0) And DateValue(primer) + z < Date Then   '<===And primer + z < Date
                             '-No marco  Sal   valida vs autoriza
                             If cDiahora = "H" And cTipo = "S" Then
                                If (Val(p_sal) - Val(chorafin) > 0) Or (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                                   F_falta = True
                                   ciFALTA = "k"
                                   cDescuenta = "S"
                                     'If wMuestra = "S" Then
                                     '   MsgBox (Format(primer  + z, "dd/mm/yyyy") + "9")
                                     'End If
                                End If
                             ElseIf cDiahora = "D" Then     'valida autoriza Dia  pierde
                                If cDescuenta = "S" Then
                                   F_falta = True
                                   ciFALTA = "l"
                                     'If wMuestra = "S" Then
                                     '   MsgBox (Format(primer  + z, "dd/mm/yyyy") + "10")
                                     'End If
                                     
                                End If
                             Else
                                F_falta = True
                                ciFALTA = "m"
                                   'If wMuestra = "S" Then
                                   '  MsgBox (Format(primer  + z, "dd/mm/yyyy") + "11")
                                   'End If
                             End If
                      ElseIf Val(p_sal1) > 0 And p0_sal <> p0_ing1 And DateValue(primer) + z < Date Then     'p_horas<>"2400"
                            If Val(m_sal1) = 0 Then
                                  F_falta = True
                            Else
                                 If Val(p_sal1) > Val(m_sal1) Then
                                   F_falta = True
                                 End If
                            End If
                             
                       ElseIf Val(m_sal) > 0 And (Val(p0_sal) > Val(m_sal)) And DateValue(primer) + z < Date Then   'p_sal
                             'Marco antes de, Autoriza H
                             If cDiahora = "H" And cTipo = "S" Then     '====> And cTipo = "S" ** TO CHECK **
                                cInicio = p_sal
                                cFin = chorafin
                                If Val(p_sal) - Val(chorafin) > 0 Then
                                   F_falta = True
                                   ciFALTA = "n"
                                     'If wMuestra = "S" Then
                                     '   MsgBox (Format(DateValue(Txfecini)  + z, "dd/mm/yyyy") + "12")
                                     'End If
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
                                   ciFALTA = "o"
                                     'If wMuestra = "S" Then
                                     '   MsgBox (Format(primer  + z, "dd/mm/yyyy") + "13")
                                     'End If
                                     
                                ElseIf Val(p_sal1) - Val(cFin) > 0 Then
                                        F_falta = True
                                     
                                End If
                             End If
                       End If
                          
                       If F_falta Then
                       Else
                          If cDiahora = "H" And cTipo = "I" Then
                          ElseIf cDiahora <> "D" And primer + z < Date Then     '<===Evalua <>D y primer + z < Date
                                If (Val(m_ing) + Val(m_ing1) > 0) Then
                                    '-Hay marcas, aplica tolerancia, calcula tardanza
                                   If cDiahora = "H" And cTipo = "I" Then
                                      cInicio = p_ing
                                      cFin = choraini
                                      nmintarde = CalcMints(cInicio, cFin)
                                   Else
                                      nmintarde = CalcMTard(p_ing, m_ing)
                                   End If
                                   If nmintarde < 6 Then
                                      nmintarde = 0
                                   End If
                                   If nmintarde >= 6 And nmintarde < 31 Then
                                      If nmintarde >= 6 And nmintarde <= 10 Then
                                         mdscto = 10
                                      End If
                                      If nmintarde >= 11 And nmintarde <= 20 Then
                                         mdscto = 20
                                      End If
                                      If nmintarde >= 21 And nmintarde <= 30 Then
                                         mdscto = 30
                                      End If
                                   Else
                                      If nmintarde > 30 Then
                                         F_falta = True
                                         ciFALTA = "p"
                                           'If wMuestra = "S" Then
                                           '   MsgBox (Format(primer  + z, "dd/mm/yyyy") + "14")
                                           'End If
                                      End If
                                      nmintarde = 0
                                   End If
                                End If
                          End If    'Fin aplicar Tolerancia
                       End If
                          
                    End If 'Fin evaluar Guardia programada
                 End If     '<=========='Fin evaluar Marcas Guardia programada
                
                 '-06.03.2020 ajuste conteo D/N, caso Autoriza D y mantguardia
                 If cDiahora = "D" And (cAuPrg <> "" And cMantGuardia = "S") Then
                    If F_falta Then
                       '-MsgBox (Format(primer  + z, "dd/mm/yyyy") + "->" + ciFALTA)
                       F_falta = False
                    End If
                 End If
                 
                 'Acumula guardias Efectivas
                 If F_falta Then
                 Else
                    If Val(p_horas) > 0 Then      'guardia programada
                       '-Caso aplico  VA/COMP  contra turno programado
                       If InStr("1DN4HCP", Left(cTurno, 1)) > 0 And cDiahora = "D" And (cAuPrg = "VA" Or cAuPrg = "COMP") Then
                          '-pierde guardia equivale a 6hrs trabajadas
                       Else
                       
                          If cAuPrg = "" Or (cAuPrg <> "" And cMantGuardia = "S") Then
                              cPatron = Mid(ccadena, i, 1)     ' F/O
                              If cPatron = "F" Then
                                 If Left(Trim(cTurno), 1) = "D" Then
                                    ndf = ndf + 1
                                 ElseIf InStr("N_H_C", Left(Trim(cTurno), 1)) > 0 And cTurno <> "NA" Then
                                       nnf = nnf + 1
                                 ElseIf InStr("P", Left(Trim(cTurno), 1)) > 0 And cTurno <> "PA" Then
                                       ndf = ndf + 1
                                       nnf = nnf + 1
                                 End If
                              ElseIf cPatron = "O" Then
                                 If Left(Trim(cTurno), 1) = "D" Then
                                    ndo = ndo + 1
                                 ElseIf InStr("N_H_C", Left(Trim(cTurno), 1)) > 0 And cTurno <> "NA" Then
                                        nno = nno + 1
                                 ElseIf InStr("P", Left(Trim(cTurno), 1)) > 0 And cTurno <> "PA" Then
                                        ndo = ndo + 1
                                        nno = nno + 1
                                 End If
                              End If
                          End If       'Autorizado y mantguardia
                       
                       End If
                    End If
                 End If
                 
              End If    'Fin Dia
           Next i
           nimporte = 0#
           If ndo + nno + ndf + nnf > 0 Then
              nimporte = ndo * nTarDO + nno * nTarNO + ndf * nTarDF + nnf * nTarNF
              rt0.AddNew
              rt0!codigo = cperiodo
              rt0!nombre = Left(caPat + " " + caMat + " " + cnomb, 45)
              rt0!libele = Left(cDni, 8)
              cvalor = Format(nimporte, "######.00")
              nlen = Len(cvalor)
              If nimporte > 0 Then
                 rt0!Monto = Left(cvalor, nlen - 3) + Right(cvalor, 2)
              Else
                 rt0!Monto = "00"
              End If
              rt0.Update
              
              rt.AddNew
              rt!CCosto = c_Costo
              rt!Dsccosto = cDscCosto
              rt!CodUpr = cCodUpr
              rt!dscUPR = cDscUpr
              rt!CadHep = ccadhep
              rt!Dsccadhep = cDcadhep
              rt!TARJETA = cTarjeta
              rt!apelpat = caPat
              rt!apelmat = caMat
              rt!nombres = cnomb
              rt!Dni = cDni
              rt!NivGuardia = cnivGuardia
              rt!DsccarPla = cDcarpla
              rt!TarifaDO = nTarDO
              rt!TarifaNO = nTarNO
              rt!TarifaDF = nTarDF
              rt!TarifaNF = nTarNF
              rt!D_O = ndo
              rt!N_O = nno
              rt!D_F = ndf
              rt!N_F = nnf
              rt!Monto = nimporte
              rt.Update
              nvalor = nvalor + 1
              ntotpresu = ntotpresu + nimporte
              If Check2.Value = 1 And (Not F_Existe Or F_Modif) Then     'Programa Final
                 nposicion = nposicion + 1
                 SQ1 = "SELECT  * FROM CpDirectorio  WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes) & "' and tarjeta='" & cTarjeta & "'"
                 Set rs_d = New Recordset
                 rs_d.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
                 rs_d!nroresolu = UCase(TxNroResolu.Text)
                 rs_d!D_O = ndo
                 rs_d!N_O = nno
                 rs_d!D_F = ndf
                 rs_d!N_F = nnf
                 rs_d!importe = nimporte
                 rs_d!UsuResolu = wUsuario      'Firma Usuario
                 rs_d!Fecresolu = Date
                 rs_d.Update
                 rs_d.Close
              End If
           End If
        End If
        r_d.MoveNext
     Loop
  End If
  r_d.Close
  rt0.Close
  rt.Close
  cmdcarga.Caption = ""
  If nvalor > 0 Then
     MsgBox "Total Servidores " + Trim(Str(nvalor)) + "  Monto S/." + Format(ntotpresu, "#######0.00"), vbExclamation
     CmdPrint.SetFocus
  End If
  Exit Sub
  'Call Limpiacampos
End Sub

Private Sub TxNroResolu_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     TxNroResolu = UCase(TxNroResolu)
     Exit Sub
   End If
End Sub
Private Sub CmdExportaTXT_Click()
    '-Estructura   tabla 1047mmaa.dbf
    'codigo chr(8),    libele chr(8),   Nombre chr(45),    'Monto chr(9)   sin decimales  1,245.40  => 124540
    Load ResPGText
    ResPGText.Left = 0: ResPGText.Top = 0
    ResPGText.Refresh
    ResPGText.Show
    Exit Sub
End Sub

Private Sub CmdPrint_Click()
   If Txmes.Text = "" Then
      MsgBox "Defina..  Periodo... ", vbExclamation
      Exit Sub
   End If
   If Trim(TxNroResolu.Text) = "" And Check2.Value = 1 Then
      MsgBox "Indique Nº Resolución... ", vbExclamation
      Exit Sub
   End If
   TxNroResolu.Text = UCase(TxNroResolu.Text)
   nTipRpt = 13
   If Check1.Value = 1 Then
      nTipRpt = 13.01
   End If
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Call Limpiacampos
   CmdPrint.Caption = ""
   Exit Sub
End Sub
Private Sub CmdClear_Click()
   Call Limpiacampos
End Sub
Private Sub Limpiacampos()
   CboMes.Text = "":   Txmes = "": TxNombMes = "": TxNroResolu = "":  Check2.Value = 0:  Check2.Enabled = Enabled
   TxFecIni = "": TxFecFin = "": TxIniLac = "": TxFinLac = ""
   TxFInicio = "": TxFFin = "": nDiaI = 0: nDiaF = 0
   'Txtarjeta.Text = "": Txtarjeta.Enabled = False
End Sub
Private Sub cmdsalir_Click()
   Set r_d = Nothing
   Set rs = Nothing
   Set rs_d = Nothing
   Set rt0 = Nothing
   Set rt = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If UCase(Forms(cont).Name) <> "CPMENU" And UCase(Forms(cont).Name) <> "RESPGOGUAR" Then
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
   Set rt0 = Nothing
   Set rt = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "ResPgoGuar" Then
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
