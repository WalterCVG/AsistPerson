VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PresGuarPeriodo 
   Caption         =   "Presupuesto al 100% - Pago de Guardias"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "PresGuarPeriodo.frx":0000
   ScaleHeight     =   2820
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   3
      Top             =   720
      Width           =   1095
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
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   5880
      TabIndex        =   20
      Top             =   1440
      Width           =   2655
      Begin VB.CommandButton mnuFicImpSelec 
         Height          =   675
         Left            =   360
         Picture         =   "PresGuarPeriodo.frx":80266
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Selec.Impresora"
         Top             =   480
         Width           =   645
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
         Picture         =   "PresGuarPeriodo.frx":80A44
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Imprimir"
         Top             =   480
         Width           =   645
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Alfabetico"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox Txtarjeta 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      MaxLength       =   5
      TabIndex        =   19
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TxNroResolu 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1320
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox TxMesAno 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CmdExportaTXT 
      Height          =   615
      Left            =   7800
      Picture         =   "PresGuarPeriodo.frx":8130E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exporta a Txt"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxMesdeAno 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   2160
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
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox TxMesActual 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   2400
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
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Top             =   1800
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
      Left            =   1320
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   675
      Left            =   8760
      Picture         =   "PresGuarPeriodo.frx":81F50
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cerrar"
      Top             =   1920
      Width           =   765
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   1095
      Left            =   6480
      Picture         =   "PresGuarPeriodo.frx":82392
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ubicar Registros"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   8760
      Picture         =   "PresGuarPeriodo.frx":8256D
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Otra Consulta"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "PresGuarPeriodo.frx":82E37
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   1935
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
      TabIndex        =   25
      Top             =   720
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
      Left            =   4080
      TabIndex        =   24
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Presupuesto 100%"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   1320
      Width           =   1695
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
      Left            =   1920
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "PresGuarPeriodo"
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
Dim ntotpresu As Double
Dim nposicion, nvalor As Integer
Dim sql, SQ1, c_a, cplh, cTarjeta, cDni, cTurno, cGrpGen, cnivGuardia, ccadena, cPatron, wMuestra As String
Dim c_Costo, cDscCosto, cCodUpr, cDscUpr, ccadhep, cDcadhep, ccarpla, cDcarpla As String
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
     If dato = Trim(Txmes) Then
        CboMes.ListIndex = i
        Exit For
     End If
  Next i
  unafecha = "01/" + Trim(Txmes) + "/" + Trim(Str(genaño))
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
         Txmes = Right(CboMes.List(j), nlen - nposicion)
         TxNombMes = Left(CboMes.List(j), nposicion - 1)
         unafecha = "01/" + Trim(Txmes) + "/" + Trim(Str(genaño))
         Call ObtenerPrimerUltimoDia(unafecha)    'primer/ultimo
         TxNombMes = Left(CboMes.List(j), nposicion - 1)
         TxMesdeAno = TxNombMes.Text + " del " & genaño
         TxMesAno = TxNombMes + " " & genaño
         TxNroResolu = "PRESU_" + Trim(Str(genaño)) + Txmes
   
         TxFecIni = primer
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
            Txmes = Right(CboMes.List(j), nlen - nposicion)
            unafecha = "01/" + Trim(Txmes) + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)            'primer/ultimo
            TxNombMes = Left(CboMes.List(j), nposicion - 1)
            TxMesdeAno = TxNombMes + " del " & genaño
            TxMesAno = TxNombMes + " " & genaño
            TxNroResolu = "PRESU_" + Trim(Str(genaño)) + Txmes
            
            TxFecIni = primer
            TxFecFin = ultimo
         End If
      End If
      Exit Sub
   End If
End Sub

Private Sub txfInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If IsDate(TxFInicio) Then
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
            MsgBox "verifique meses en fechas ", vbExclamation
            Exit Sub
        End If
        If Year(DateValue(TxFInicio)) <> Year(DateValue(TxFFin)) Then
           MsgBox "fechas de años diferentes.. ", vbExclamation
           Exit Sub
        End If
        If Year(DateValue(TxFFin)) <> Val(genaño) Then
           MsgBox "fecha no corresponde al periodo.." + cPatron, vbExclamation
           Exit Sub
        End If
     End If
  End If
End Sub
Private Sub TxFFin_LostFocus()
   Call txfFin_KeyPress(13)
End Sub

Private Sub cmdcarga_Click()
   If Txmes = "" And (Trim(TxFInicio) = "" Or Trim(TxFFin) = "") Then
      MsgBox "No inidico perido a procesar ", vbExclamation
      CboMes.SetFocus
      Exit Sub
   End If
     
   If Val(Txmes) > 0 Then
      '-PERIODO
      Call Cbomes_KeyPress(13)
   Else
      '-lapso dias
      Txmes = "0" + Trim(Str(Month(TxFFin)))
      unafecha = "01/" + Trim(Txmes) + "/" + Trim(Str(genaño))
      Call ObtenerPrimerUltimoDia(unafecha)    'Obtiene primer/ultimo
      
      TxNombMes = Trim(Mid(wMeses, (Val(Txmes) - 1) * 9 + 1, 9))
      TxMesdeAno = TxNombMes + " del " & genaño
      TxMesAno = TxNombMes + " " & genaño
      TxNroResolu = "PRESU" + Trim(Str(genaño)) + Txmes + "de" + Right("0" + Trim(Str(Day(TxFInicio))), 2) + "a" + Right("0" + Trim(Str(Day(TxFFin))), 2)
      TxFecIni = TxFInicio
      TxFecFin = TxFFin
   End If
   
   wPer_resol = Txmes + Right(Trim(Str(genaño)), 2)
   cperiodo = "1047" + Txmes + Right(Trim(Str(genaño)), 2)
   Cnx.Execute "DELETE  FROM CpPG1047"
   Set rt0 = New Recordset
   SQ1 = "SELECT * FROM CpPG1047"
   rt0.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   
   Cnx.Execute "DELETE  FROM CpTmPagoGuar"
   Set rt = New Recordset
   SQ1 = "SELECT * FROM CpTmPagoGuar"
   rt.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
  
   ccadena = ""
   For i = Day(primer) To Day(DateValue(TxFecFin))
      unafecha = primer + i - 1
      f_Fer = False     ' verifica Feriado/domingo
      If UCase(Left(WeekdayName(Weekday(primer + i - 1)), 3)) = "DOM" Then
         f_Fer = True
         
              '''MsgBox Format(unafecha, "dd/mm/yyyy") + ":" + Str(i) + "-" + WeekdayName(Weekday(unafecha)) + "-" + Left(WeekdayName(Weekday(unafecha)), 3)
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
               
               '''  MsgBox rs!codigo + "-" + Format(unafecha, "dd/mm/yyyy") + "-" + c_a
               
            End If
            rs.MoveNext
         Wend
         rs.Close
      End If
      ccadena = ccadena + IIf(f_Fer, "F", "O")
   Next i
      
   '''  MsgBox ccadena
     
   
   '-carga  DIRECTORIO "TURNO ASISTENCIALES"
   sql = "SELECT  * FROM CpDirectorio  WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes) & "' and flag='0' " & _
        "ORDER BY Tarjeta, ano, mes "
   Set r_d = New Recordset
   r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not r_d.EOF Then
     nvalor = 0: nposicion = 0
     ntotpresu = 0#
     wMuestra = "N"
     
      Dim OpenForms
      v_an = 0
     
     Do While Not r_d.EOF
     
        '-dar control a windows y permitir multihilo
        v_an = v_an + 1
        If v_an Mod 10 = 0 Then     'if loop has repeated
           OpenForms = DoEvents 'Yield to operating system
        End If
     
        cTarjeta = r_d!tarjeta
        c_Costo = IIf(IsNull(r_d!CCosto), "", r_d!CCosto)
        cCodUpr = IIf(IsNull(r_d!CodUpr), "", r_d!CodUpr)
        '-SABER si una columna existe en una tabla
        SQ1 = "SELECT * FROM  INFORMATION_SCHEMA.COLUMNS AS c1  WHERE  c1.column_name = 'cadhep'  AND c1.table_name = 'cpdirectorio'"
        Set rs = New Recordset
        rs.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
           ccadhep = IIf(IsNull(r_d!CadHep), "", Trim(r_d!CadHep))
        End If
        rs.Close
        
        Set rs = New Recordset
        sql = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & cTarjeta & "'"
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        cDni = "": cplh = "": cDni = "": caPat = "": caMat = "": cnomb = "": cnivGuardia = "": cGrpGen = "": ccadhep = "": cDcadhep = ""
        ccarpla = "": cDcarpla = "": cCodSiaf = ""
        If Not rs.EOF Then
           cplh = IIf(IsNull(rs!PLH), "", rs!PLH)                  'guardias  HEP/RES
           cDni = IIf(IsNull(rs!Documento), "", rs!Documento)
           caPat = Trim(rs!apelpat)
           caMat = Trim(rs!apelmat)
           cnomb = Trim(rs!nombres)
           cnivGuardia = IIf(IsNull(rs!NivGuardia), "", rs!NivGuardia)
           ccarpla = IIf(IsNull(rs!carPla), "", rs!carPla)
           cGrpGen = IIf(IsNull(rs!grpgen), "", rs!grpgen)
           If ccadhep = "" Then
              ccadhep = IIf(IsNull(rs!CadHep), "", rs!CadHep)
           End If
        End If
        rs.Close
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
           '-Procesa Directorio-CALENDARIO
           For i = Day(DateValue(TxFecIni)) To Day(DateValue(TxFecFin))
              z = i - 1
              cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", Trim(r_d.Fields(z + 4)))
              cmdcarga.Caption = cTarjeta + " " + Str(i) + "-" + cTurno
              
              '--  If cTarjeta = "39439" Then
              '--     MsgBox cTarjeta + " " + Str(i) + "-" + cTurno
              '--  End If
              
              '-Procesa Guardias
              If InStr("1_DD*_D1D2D3D4_NN*_N2N3N4_4_HCP_H*_C*C1H1_P*_P1", cTurno) > 0 And cTurno <> "*" And Trim(cTurno) <> "" Then
                 F_falta = False
                 '-Verifica Asistencia Y/o Autorizaciones "D" que afectan Guardia
                 fechaB = Format(primer + z, "yyyymmdd")
                 dfentra = primer + z
                 dfsale = primer + z
                 p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = ""
                 m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
                 nValorH = 0: nValorM = 0: mmintarde = 0:  mdscto = 0: nFaltaH = 0: nFaltaM = 0: F_falta = False: nmotpersonal = 0: nComision = 0
                 cvalor = "": cAuPrg = "": cDiahora = "": cTipo = "": choraini = "": chorafin = "": cDescuenta = "": c_a = ""
                 
                 '-HORARIO prog en Dia
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
                 
                 'If cTarjeta = "39439" Then
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
                          p_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
                          p_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                          p_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
                          p_sal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
                          p_horas = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
                          nValorH = Val(Left(p_horas, 2))
                          nValorM = Val(Right(p_horas, 2))
                          nFaltaH = Val(Left(p_horas, 2))
                          nFaltaM = Val(Right(p_horas, 2))
                       End If
                       rs_d.Close
                    End If
                 End If
                 rs.Close
                 If Val(p_horas) > 0 Then     'Programado
                    cPatron = Mid(ccadena, i, 1)     ' F/O
                    If cPatron = "F" Then
                       If InStr("1_DD*_D1D2D3D4", cTurno) > 0 Then
                          ndf = ndf + 1
                       ElseIf InStr("4_NN*N2N3N4_HH*_CC*_H1C1", cTurno) > 0 Then
                              nnf = nnf + 1
                       ElseIf InStr("PP*P1", cTurno) > 0 Then
                              ndf = ndf + 1
                              nnf = nnf + 1
                       End If
                    ElseIf cPatron = "O" Then
                           If InStr("1_DD*_D1D2D3D4", cTurno) > 0 Then
                              ndo = ndo + 1
                           ElseIf InStr("4_NN*N2N3N4_HH*_CC*_H1C1", cTurno) > 0 Then
                                  nno = nno + 1
                           ElseIf InStr("PP*_P1", cTurno) > 0 Then
                                  ndo = ndo + 1
                                  nno = nno + 1
                           End If
                    End If
                 End If
              End If
              
        '--      If cTarjeta = "39439" Then  'And Trim(cTurno) <> ""
        '--         MsgBox cTarjeta + "->dia:" + Trim(Str(i)) + "-Hrs " + p_horas + " patron[" + cPatron + "]" + cTurno + "] van: d-" + Trim(Str(ndo)) + " n-" + Trim(Str(nno)) + " df-" + Trim(Str(ndf)) + " nf-" + Trim(Str(nnf))
        '--      End If
              
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
              rt!tarjeta = cTarjeta
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
   If Txmes = "" Then
      MsgBox "Defina..  Periodo... ", vbExclamation
      Exit Sub
   End If
   TxNroResolu = UCase(TxNroResolu)
   nTipRpt = 13.5
   If Check1.Value = 1 Then
      nTipRpt = 13.51
   End If
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Call Limpiacampos
   CmdPrint.Caption = ""
   Exit Sub
End Sub
Private Sub Limpiacampos()
   CboMes.Text = "":   Txmes = "": TxNombMes = "": TxNroResolu = ""
   TxFInicio = "": TxFFin = ""
   TxFecIni = "": TxFecFin = ""
   'Txtarjeta.Text = "": Txtarjeta.Enabled = False
End Sub
Private Sub CmdClear_Click()
   Call Limpiacampos
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
        If UCase(Forms(cont).Name) <> "CPMENU" And UCase(Forms(cont).Name) <> "PRESPGOGUAR" Then
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


