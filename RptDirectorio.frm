VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RptDirectorio 
   Caption         =   "Listar Directorio"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "RptDirectorio.frx":0000
   ScaleHeight     =   3525
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   8400
      Picture         =   "RptDirectorio.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cerrar"
      Top             =   2760
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8520
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton mnuFicImpSelec 
      Height          =   495
      Left            =   8280
      Picture         =   "RptDirectorio.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Selec.Impresora"
      Top             =   720
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   6855
      Begin VB.CommandButton CmdHuerfanos 
         BackColor       =   &H008080FF&
         Caption         =   "Asistenciales  No Programados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5160
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo 3ros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   1920
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo Asistentes"
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton Option0 
         Caption         =   "ADMINISTRATIVOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton Option0 
         Caption         =   "ASISTENCIALES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox TxDia 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Txtrabajador 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   480
         Width           =   3975
      End
      Begin VB.TextBox Txtarjeta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "DIARIO "
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "ANUAL del SERVIDOR"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6840
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha del Directorio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label_2 
         BackStyle       =   0  'Transparent
         Caption         =   "Trabajador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   855
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
      Height          =   915
      Left            =   8040
      Picture         =   "RptDirectorio.frx":80B8E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir"
      Top             =   1440
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Picture         =   "RptDirectorio.frx":81458
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "RptDirectorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim r_d As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim rs_t As New ADODB.Recordset
Dim Sw, F_add, lcreatmp, f_cierra As Boolean
Dim nLacta, nSonLacta, nposicion, nson, nPersonal, npersout, nGrabo As Integer
Dim csexo, cAplica, cDni, sql, ceditor, cTarjeta, cNombre, cano, c_ano, cmes, c_a, cCarest, cDscCargo, CCosto, cDscCosto, cTurno, cBloque, cOrden, cCadMarca, cMarca As String
Dim unafecha As Date
Dim fechaAnt, canoAnt, cmesant, cTurnoAnt, cDiaOUT      As String
Dim ndiaAnt As Integer

Private Sub Command1_Click()
     c_grupo = "03": c_secuencia = "031"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub
Private Sub Form_Load()
   RptDirectorio.Caption = "Listar Directorios " + genaño
   Txtarjeta = "": Txtrabajador = ""
   Option0(0).Value = False: Option0(1).Value = False
   TxDia = "": Check1.Value = 0: c_ano = Trim(Str(genaño))
End Sub

Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(Txtarjeta) = "" Or (Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta)) Then
        If Trim(Txtarjeta) <> "" And EsTexto(Txtarjeta) Then
           Ayuda1.TxUbiAP = Trim(Txtarjeta)
        End If
        Load Ayuda1
        Ayuda1.Tag = 7
        Ayuda1.Show 1
     Else
        Set rs = New Recordset
        sq2 = "SELECT * FROM CpGeneral (baja is NULL) and WHERE tarjeta = '" & Txtarjeta & "'"
        rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If rs.EOF Then
           MsgBox "No Ubicado", vbExclamation
        Else
           Txtarjeta = rs!tarjeta
           Txtrabajador = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
        End If
        rs.Close
     End If
  End If
End Sub
Private Sub txDia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If IsDate(TxDia) Then
        c_ano = Trim(Str(Year(DateValue(TxDia))))
        If Year(DateValue(TxDia)) <> Val(genaño) Then
           If MsgBox("No corresponde al año activo ", vbOKCancel + vbQuestion, "Confirme") = vbCancel Then
              Call Limpiacampos
           End If
        End If
     Else
        MsgBox "verifique fecha ..", vbExclamation
     End If
     Exit Sub
  End If
End Sub
Private Sub TxDia_LostFocus()
   If TxDia <> "" Then
      Call txDia_KeyPress(13)
   End If
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
      Case 1
   End Select
End Sub
Private Sub Limpiacampos()
   Txtarjeta = "":   Txtrabajador = ""
   Option0(0).Value = False: Option0(1).Value = False
   TxDia = "": TxIniLac = "": TxFinLac = "": c_ano = Str(genaño)
   Check1.Value = 0
End Sub
Private Sub CmdClear_Click()
   Limpiacampos
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

Private Sub CmdPrint_Click()
   If Txtrabajador = "" And Not IsDate(TxDia) Then
      MsgBox "No Definio..  parametros de Directorio ", vbExclamation
      Exit Sub
   End If
   ceditor = "X"
   If Wccosto = "" Or Wccosto = "14" Or Wccosto = "29" Then     'Nutricion/Comunicación
      ceditor = "S"
   Else
      If IsDate(TxDia) Then
         '-Directorio del dia
         fecha1 = Format(TxDia, "yyyymmdd")    'convierto a codigo ANSI
         cmes = Mid(fecha1, 5, 2)
         unafecha = TxDia
         Call ObtenerPrimerUltimoDia(unafecha)
         fecha1 = Format(primer, "yyyymmdd")
         fecha2 = Format(ultimo, "yyyymmdd")
         '-valida JG en  DIRECTORIO  -*
         sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & wTarjeta & "' AND ano ='" & c_ano & "' AND mes='" & cmes & "'  and FLAG <> '2'"
         Set rs = New Recordset
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then
            For i = Day(primer) To Day(ultimo)
               z = i - 1
               cTurno = Trim(IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4)))
               If InStr("D*_N*_H*_C*_P*", cTurno) > 0 And Len(cTurno) = 2 Then
                  ceditor = "S"
                  Exit For
               End If
            Next i
         End If
         rs.Close
      End If
   End If
   If ceditor = "X" Then
      MsgBox "Operador, no Autorizado...", vbExclamation
      Exit Sub
   End If

   If Trim(Txtrabajador) <> "" Then
      '-KARDEX Anual trabajador
      CargaDirect1
      nTipRpt = 6
      InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
      InfoCpSQL.Refresh
      InfoCpSQL.Show
   ElseIf IsDate(TxDia) Then
      '-Directorio Diario servidores
      If Check1.Value = 0 Then       '---Programado  100%
         Sw = False
         ''''If Weekday(DateValue(TxDia)) = 1 Then
         If UCase(Left(WeekdayName(Weekday(DateValue(TxDia))), 3)) = "DOM" Then
            Sw = True
         Else
            Set rs = New Recordset
            rs.Open "SELECT * FROM Feriados ORDER BY codigo", Cnx0, adOpenStatic, adLockReadOnly
            rs.MoveFirst
            While Not rs.EOF
               If Len(Trim(rs!Codigo)) = 6 Then     'MMddAA
                  c_a = Right(rs!Codigo, 2)
               ElseIf Len(Trim(rs!Codigo)) = 4 Then
                      c_a = Right(Year(TxDia), 2)
               End If
               If Val(Left(rs!Codigo, 2)) = Month(TxDia) And Val(Mid(rs!Codigo, 3, 2)) = i And Right(Year(TxDia.Text), 2) = c_a Then
                  Sw = True
               End If
               rs.MoveNext
            Wend
            rs.Close
         End If
         If Option0(1).Value = True And Sw Then
            MsgBox "Es dia FERIADO  ", vbExclamation
            Exit Sub
         End If
         
          ''  nPersonal = 0: npersout = 0: nGrabo = 0
            
         CargaDirect2
         
          ''   MsgBox "[" + Str(nPersonal) + "/" + Str(npersout) + "-" + Str(nGrabo) + "]"
             
         nTipRpt = 7
         unafecha = TxDia
         InfoCpRS.Left = 0: InfoCpRS.Top = 0
         InfoCpRS.Refresh
         InfoCpRS.Show
      Else           '----solo ASISTENTES
         Sw = False
         '''If Weekday(DateValue(TxDia)) = 1 Then
         If UCase(Left(WeekdayName(Weekday(DateValue(TxDia))), 3)) = "DOM" Then
            Sw = True
         Else
            Set rs = New Recordset
            rs.Open "SELECT * FROM Feriados ORDER BY codigo", Cnx0, adOpenStatic, adLockReadOnly
            'rs.Open "SELECT * FROM CpTablasAux WHERE grupo='FER' ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
            rs.MoveFirst
            While Not rs.EOF
               If Len(Trim(rs!Codigo)) = 6 Then
                  c_a = Right(rs!Codigo, 2)
               ElseIf Len(Trim(rs!Codigo)) = 4 Then
                      c_a = Right(Year(TxDia), 2)
               End If
               If Val(Left(rs!Codigo, 2)) = Month(TxDia) And Val(Mid(rs!Codigo, 3, 2)) = i And Right(Year(TxDia), 2) = c_a Then
                  Sw = True
               End If
               rs.MoveNext
            Wend
            rs.Close
         End If
         If Option0(1).Value = True And Sw Then
            MsgBox "Es dia FERIADO  ", vbExclamation
            Exit Sub
         End If
         f_cierra = True
         If DateValue(TxDia) = Date Then
            '-solo marcas ingreso ó programados
            f_cierra = False
         End If
         CargaDirect3
         nTipRpt = 7
         unafecha = TxDia
         InfoCpRS.Left = 0: InfoCpRS.Top = 0
         InfoCpRS.Refresh
         InfoCpRS.Show
      End If
   End If
   Call Limpiacampos
   CmdPrint.Caption = ""
   Call cmdsalir_Click
   Exit Sub
End Sub
Private Sub CargaDirect1()
'-CARGA  registros    1:Kardex Anual
    If Trim(Txtarjeta) = "" And Not IsDate(TxDia) Then
       MsgBox "No trabajador / fecha, para Directorio.. ", vbExclamation
       Exit Sub
    End If
    On Error Resume Next
    'Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmpDirKar"
    nposicion = 0
    Set rs_d = New Recordset
    sq_l = "SELECT * FROM CpTmpDirKar"
    rs_d.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    '-Directorio Anual del trabajador
    Set r_d = New Recordset
    sql = "SELECT  * FROM CpDirectorio WHERE Tarjeta ='" & Txtarjeta & "' AND ano='" & c_ano & "' and FLAG<>'2'" & _
    "ORDER BY Tarjeta, ano, mes "
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not r_d.EOF
       rs_d.AddNew
       CmdPrint.Caption = r_d!mes + " - " + r_d!tarjeta
       rs_d!tarjeta = r_d!tarjeta
       rs_d!ano = r_d!ano
       rs_d!mes = r_d!mes
       cCarest = IIf(IsEmpty(r_d!carEst), "", r_d!carEst)
       rs_d!carEst = cCarest
       unafecha = "01/" + Trim(r_d!mes) + "/" + c_ano
       Call ObtenerPrimerUltimoDia(unafecha)
       For i = Day(primer) To Day(ultimo)
          z = i - 1
          rs_d.Fields(z + 4) = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
          If nopc = 2 And z = nposicion Then
             cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
             rs_d!turno = cTurno
             If cTurno <> "" And InStr("*DN", Trim(cTurno)) = 0 Then
                cCarest = IIf(InStr("1234_D*_N*_H*_C*_P*", Trim(cTurno)) > 0, "JG", cCarest)
             End If
             rs_d!carEst = cCarest
             Set rs = New Recordset
             rs.Open "SELECT * FROM cptablas WHERE tabla='27' AND codigo='" & cCarest & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If rs.EOF Then
             Else
                rs_d!orden = IIf(IsNull(rs!campo), "", Trim(rs!campo))
             End If
             rs.Close
          End If
       Next i
       Set rs = New Recordset
       sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Txtarjeta & "'"
       rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       cCodigo = ""
       If rs.EOF Then
       Else
          cCodigo = IIf(IsNull(rs!actividad), "", rs!actividad)
       End If
       rs.Close
       rs_d!actividad = cCodigo
       Set rs = New Recordset
       rs.Open "SELECT * FROM Cptablas WHERE tabla='27' and codigo='" & cCarest & "'", Cnx, adOpenStatic, adLockReadOnly
       If rs.EOF Then
       Else
          rs_d!DscCargo = IIf(IsNull(rs!descripcion), "", Left(rs!descripcion, 30))
       End If
       rs_d.Update
       rs.Close
       r_d.MoveNext
    Loop
    'Cnx.CommitTrans
    rs_d.Close
    r_d.Close
End Sub
Private Sub CargaDirect2()
'-Directorio Diario PROGRAMADO - Asistencial /Administrativo  =>  Todos /3ros
    If Not IsDate(TxDia) Then
       MsgBox "No Indico fecha, del Directorio.. ", vbExclamation
       Exit Sub
    End If
    Cnx.BeginTrans
    On Error GoTo crea1
    Cnx.Execute "DELETE  FROM #tabla2"
    GoTo sigue1
crea1:
    Cnx.Execute ("SELECT * Into #tabla2  FROM CpTmpDirect WHERE 1 = 2")
sigue1:
    Cnx.CommitTrans
    nson = 0
    fecha1 = Format(TxDia, "yyyymmdd")    'convierto a codigo ANSI
    cmes = Mid(fecha1, 5, 2)
    nposicion = Day(TxDia) - 1
    Set r_d = New Recordset
    If Check2.Value = 0 Then  'no incluye plh=>'3RO'
    sql = "SELECT  * FROM CpDirectorio WHERE ano='" & c_ano & "' AND mes='" & cmes & "' AND (PLH is NULL OR PLH<>'3RO') and FLAG <>'2' " & _
       "ORDER BY Tarjeta, ano, mes "
    Else   'Solo plh='3RO'
       sql = "SELECT  * FROM CpDirectorio WHERE ano='" & c_ano & "' AND mes='" & cmes & "' AND PLH ='3RO'  and FLAG <> '2'" & _
       "ORDER BY Tarjeta, ano, mes "
    End If
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not r_d.EOF
       cNombre = "":  CCosto = "": cDscCosto = "": cCarest = "":  cDscCargo = "": cBloque = "": cBloque1 = "": cBloque2 = "": cdscbloqu1 = "": cdscbloqu2 = "": cOrden = ""
       CmdPrint.Caption = r_d!tarjeta + " - " + IIf(IsNull(r_d.Fields(nposicion + 4)), "", Trim(r_d.Fields(nposicion + 4)))
       If InStr("*LVR", Trim(r_d.Fields(nposicion + 4))) > 0 Or IsNull(r_d.Fields(nposicion + 4)) Or IsEmpty(r_d.Fields(nposicion + 4)) Then
          '-No descanso/licencia/vacaciones/Desc RX
       ElseIf (Option0(0).Value = True And Trim(r_d.Fields(nposicion + 4)) = "A") Then
           '-Dir Asistencial
           
       ElseIf (Option0(1).Value = True And Trim(r_d.Fields(nposicion + 4)) <> "A") Then
       'ElseIf (Option0(1).Value = True And InStr("14_D_M_T_N_HCP_H*H1_MT_NA_PA", Left(Trim(r_d.Fields(nposicion + 4)), 1)) > 0) Then
           '-Dir Administrativos
       Else
       
            nPersonal = nPersonal + 1
            
          cTarjeta = r_d!tarjeta
          
          cano = r_d!ano
          cmes = r_d!mes
          '-coge de Directorio
          cCarest = IIf(IsNull(r_d!carEst), "", r_d!carEst)
          Set rs = New Recordset
          sq_l = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta='" & cTarjeta & "'"
          
          rs.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If Not rs.EOF Then
             cNombre = Trim(rs!apelpat) + "," + Left(rs!nombres, 1) + ". ;"
             CCosto = rs!CCosto
             
              '-vacio coge parametro en maestra
              cCarest = IIf(IsEmpty(cCarest), IIf(IsNull(rs!carEst), "", rs!carEst), cCarest)
             
             If Len(cNombre) > 17 Then
                cNombre = Left(cNombre, 17)
             End If
          End If
          rs.Close
          cTurno = IIf(IsNull(r_d.Fields(nposicion + 4)), "", r_d.Fields(nposicion + 4))
          If cTurno <> "" And InStr("*DNHCP", Trim(cTurno)) = 0 Then
             cCarest = IIf(InStr("14_D*_N*_H*_C*_P*", Trim(cTurno)) > 0, "JG", cCarest)
          End If
          Set rs = New Recordset
          rs.Open "SELECT * FROM Cptablas WHERE tabla='27' and codigo='" & cCarest & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             cDscCargo = IIf(IsNull(rs!descripcion), "", Left(rs!descripcion, 30))
          End If
          rs.Close
          
          '-VALIDA NO exista AUTORIZACIONES Motivos Personales, Comisiones Serv - Excluye LVR
          FechaB = Format(TxDia, "yyyymmdd")
          cAuPrg = "": cDiahora = "N": Ctipo = "": cMantGuardia = "N"
          sql = "SELECT * FROM CpAutori" & c_ano & " WHERE (anulado is NULL) and   tarjeta='" & cTarjeta & "' and FechaIni<='" + FechaB + "' and FechaFin>='" + FechaB + "' and turno=''"
          Set rs = New Recordset
          rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If Not rs.EOF Then
             Do While Not rs.EOF
                cAuPrg = Trim(rs!Codigo)
                If rs!diahora = "D" Then
                   cDiahora = rs!diahora
                   Ctipo = IIf(IsNull(rs!tipo), "", rs!tipo)            'I/S/O   caso horas
                End If
                rs.MoveNext
             Loop
          End If
          rs.Close
          If cDiahora = "D" Then   'Autorizacion para el DIA
             If cAuPrg <> "" Then
                sql = "SELECT * FROM CpOcurrencias WHERE codigo='" & cAuPrg & "'"
                Set rs = New Recordset
                rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                If Not rs.EOF Then
                   cMantGuardia = IIf(IsNull(rs!MantGuardia), "N", rs!MantGuardia)
                End If
             End If
             rs.Close
             npersout = npersout + 1
          End If
          
          If cDirectorio <> "D" Or (cDiahora = "D" And cMantGuardia = "S") Then
                    
             Set rs = New Recordset
             rs.Open "SELECT * FROM cptablas WHERE tabla='31' AND codigo='" & cTurno & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If rs.EOF Then
             Else
                cBloque = IIf(IsNull(rs!campo), "", Trim(rs!campo))
                If Len(cBloque) > 1 Then
                   cBloque1 = Left(cBloque, 1)
                   cBloque2 = Right(cBloque, 1)
                Else
                   cBloque1 = cBloque
                End If
             End If
             If Option0(0).Value = True And Trim(cBloque) = "" Then
                MsgBox "[" + cTarjeta + "] Turno (" + cTurno + ") No definido para ASISTENCIALES (Tab.Aux 31)"
             End If
             rs.Close
             If cBloque1 <> "" Then
                Set rs = New Recordset
                rs.Open "SELECT * FROM cptablas WHERE tabla='43' AND codigo='" & cBloque1 & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
                If rs.EOF Then
                Else
                   cdscbloqu1 = IIf(IsNull(rs!descripcion), "", Trim(rs!descripcion))
                End If
                rs.Close
             End If
             If cBloque2 <> "" Then
                Set rs = New Recordset
                rs.Open "SELECT * FROM cptablas WHERE tabla='43' AND codigo='" & cBloque2 & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
                If rs.EOF Then
                Else
                   cdscbloqu2 = IIf(IsNull(rs!descripcion), "", Trim(rs!descripcion))
                End If
                rs.Close
             End If
             
             
             
             Set rs = New Recordset
             rs.Open "SELECT * FROM cptablas WHERE tabla='27' AND codigo='" & cCarest & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If rs.EOF Then
             Else
                cOrden = IIf(IsNull(rs!campo), "", Trim(rs!campo))
             End If
             If Option0(0).Value = True And Trim(cOrden) = "" Then
                MsgBox "[" + cTarjeta + "] Cargo Est. (" + cCarest + ") No definido para ASISTENCIALES (Tab.Aux 27)"
             End If
             rs.Close
             Set rs = New Recordset
             rs.Open "SELECT * FROM cpcCosto WHERE codigo='" & CCosto & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs.EOF Then
                cDscCosto = IIf(IsNull(rs!descripcion), "", Trim(rs!descripcion))
             End If
             rs.Close
             If Option0(1).Value = True Then    'Administrativo
                cOrden = CCosto
                cDscCargo = cDscCosto
             End If
             
             'If nPersonal < 5 Then
             '     MsgBox "[" + cBloque + "-" + cBloque1 + "-" + cBloque2 + "-" + cOrden + "]"
             'End If
             
              If cTarjeta = "ZZ718" Then
                 MsgBox "paso 2"
              End If

             
             
             
             If cBloque <> "" And cOrden <> "" Then
                If cBloque1 <> "" Then
                   F_add = True
                   nson = 1
                   Set rs_d = New Recordset
                      rs_d.Open "SELECT * FROM #tabla2 WHERE Bloque='" & cBloque1 & "' AND orden='" & cOrden & "'", Cnx, adOpenStatic, adLockOptimistic, adCmdText
                   If rs_d.EOF Then
                   Else
                      Do While Not rs_d.EOF
                         For i = 1 To 9
                            z = i - 1
                            If IsNull(rs_d.Fields(z + 10)) Or Trim(rs_d.Fields(z + 10)) = "" Then
                               nson = i
                               F_add = False
                               Exit For
                            End If
                         Next i
                         If Not F_add Then
                            Exit Do
                         End If
                         rs_d.MoveNext
                      Loop
                   End If
                   
                   If cTarjeta = "ZZ718" Then
                      MsgBox "paso3"
                   End If

                   
                   If F_add Then
                      nGrabo = nGrabo + 1
                      rs_d.AddNew
                      rs_d!carEst = cCarest
                      rs_d!DscCargo = cDscCargo
                      rs_d!CCosto = CCosto
                      rs_d!Dsccosto = cDscCosto
                      rs_d!bloque = cBloque1
                      rs_d!Dscbloque = cdscbloqu1
                      rs_d!orden = cOrden
                      rs_d!ano = cano
                      rs_d!mes = cmes
                      rs_d!Dia = TxDia
                      rs_d!Campo1 = cNombre
                      rs_d.Update
                   Else
                      rs_d.Fields(nson - 1 + 10) = cNombre
                      rs_d.Update
                   End If
                   rs_d.Close
                End If
                If cBloque2 <> "" Then
                   F_add = True
                   nson = 1
                   Set rs_d = New Recordset
                      rs_d.Open "SELECT * FROM #tabla2 WHERE Bloque='" & cBloque2 & "' AND orden='" & cOrden & "'", Cnx, adOpenStatic, adLockOptimistic, adCmdText
                   If rs_d.EOF Then
                   Else
                      Do While Not rs_d.EOF
                         For i = 1 To 9
                            z = i - 1
                            If IsNull(rs_d.Fields(z + 10)) Or Trim(rs_d.Fields(z + 10)) = "" Then
                               nson = i
                               F_add = False
                               Exit For
                            End If
                         Next i
                         If Not F_add Then
                            Exit Do
                         End If
                         rs_d.MoveNext
                      Loop
                   End If
                   
                     If cTarjeta = "ZZ718" Then
                        MsgBox "paso 4"
                     End If

                   
                   If F_add Then
                      rs_d.AddNew
                      rs_d!carEst = cCarest
                      rs_d!DscCargo = cDscCargo
                      rs_d!CCosto = CCosto
                      rs_d!Dsccosto = cDscCosto
                      rs_d!bloque = cBloque2
                      rs_d!Dscbloque = cdscbloqu2
                      rs_d!orden = cOrden
                      rs_d!ano = cano
                      rs_d!mes = cmes
                      rs_d!Dia = TxDia
                      rs_d!Campo1 = cNombre
                      rs_d.Update
                   Else
                      rs_d.Fields(nson - 1 + 10) = cNombre
                      rs_d.Update
                   End If
                   rs_d.Close
                End If
             End If   'Bloque/Orden
          End If    ' Autoriza
       End If
       r_d.MoveNext
    Loop
    'Cnx.CommitTrans
    r_d.Close
    Exit Sub
End Sub

Private Sub CargaDirect3()
'-Directorio Diario PRESENTES - turnos Asistencial D/N/HCP /Administrativo
    If Not IsDate(TxDia) Then
       MsgBox "No Fecha fecha, del Directorio.. ", vbExclamation
       Exit Sub
    End If
    
    Cnx.BeginTrans
  On Error GoTo CREA
     Cnx.Execute "DELETE  FROM #tabla2"
     GoTo sigue
CREA:
     Cnx.Execute ("SELECT * Into #tabla2  FROM CpTmpDirect WHERE 1 = 2")
   
sigue:
    Cnx.CommitTrans
        
    Set r_d = New Recordset
    nson = 0
    nposicion = 0
    fecha1 = Format(TxDia, "yyyymmdd")    'convierto a codigo ANSI
    cmes = Mid(fecha1, 5, 2)
    nposicion = Day(TxDia) - 1
    If Check2.Value = 0 Then
       sql = "SELECT  * FROM CpDirectorio WHERE ano='" & c_ano & "' AND mes='" & cmes & "' and (PLH is NULL OR PLH<>'3RO') AND FLAG <> '2' " & _
             "ORDER BY Tarjeta, ano, mes "
    Else
       sql = "SELECT  * FROM CpDirectorio WHERE ano='" & c_ano & "' AND mes='" & cmes & "' AND PLH='3RO'  and FLAG <> '2' " & _
             "ORDER BY Tarjeta, ano, mes "
    End If
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not r_d.EOF
       CmdPrint.Caption = r_d!tarjeta + " - " + IIf(IsNull(r_d.Fields(nposicion + 4)), "", Trim(r_d.Fields(nposicion + 4)))
       If InStr("*LVR", Trim(r_d.Fields(nposicion + 4))) > 0 Or IsNull(r_d.Fields(nposicion + 4)) Or IsEmpty(r_d.Fields(nposicion + 4)) Then
          'No descanso/licencia/vacaciones/Desc RX
       ElseIf (Option0(0).Value = True And Trim(r_d.Fields(nposicion + 4)) = "A") Then
           'Dir Asistencial
       ElseIf (Option0(1).Value = True And Trim(r_d.Fields(nposicion + 4)) <> "A") Then
           'Dir Administrativos
       Else                         '--------------->1
          cTarjeta = r_d!tarjeta
          cano = r_d!ano
          cmes = r_d!mes
          cCarest = IIf(IsEmpty(r_d!carEst), "", r_d!carEst)
          cDni = "": cNombre = "": CCosto = "": cDscCosto = "": cGrpGen = "": cBloque = ""
          cBloque1 = "": cBloque2 = "": cdscbloqu1 = "": cdscbloqu2 = "":  cOrden = ""
          Set rs = New Recordset
          sq_l = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta='" & cTarjeta & "'"
          rs.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If Not rs.EOF Then
             cNombre = Trim(rs!apelpat) + "," + Left(rs!nombres, 1) + ". ;"
             cDni = IIf(IsNull(rs!documento), "", Trim(rs!documento))
             csexo = IIf(IsNull(rs!sexo), "", rs!sexo)
             CCosto = rs!CCosto
             cGrpGen = IIf(IsNull(rs!grpgen), "", rs!grpgen)
             If Len(cNombre) > 17 Then
                cNombre = Left(cNombre, 17)
             End If
          End If
          rs.Close
                    
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
                If (DateValue(TxDia) >= rs!fecnac And DateValue(TxDia) <= dftope) Then
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
                    
          cTurno = IIf(IsNull(r_d.Fields(nposicion + 4)), "", r_d.Fields(nposicion + 4))
          If cTurno <> "" And InStr("*DNHCP_", Trim(cTurno)) = 0 Then
             cCarest = IIf(InStr("1234_D*_N*_H*_C*_P*", Trim(cTurno)) > 0, "JG", cCarest)
          End If
          Set rs = New Recordset
          rs.Open "SELECT * FROM Cptablas WHERE tabla='27' and codigo='" & cCarest & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             cDscCargo = IIf(IsNull(rs!descripcion), "", Left(rs!descripcion, 30))
          End If
          rs.Close
          Set rs = New Recordset
          rs.Open "SELECT * FROM cptablas WHERE tabla='31' AND codigo='" & cTurno & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If Not rs.EOF Then
             cBloque = IIf(IsNull(rs!campo), "", Trim(rs!campo))
             If Len(cBloque) > 1 Then
                cBloque1 = Left(cBloque, 1)
                cBloque2 = Right(cBloque, 1)
             Else
                cBloque1 = cBloque
             End If
          End If
          rs.Close
          If cBloque1 <> "" Then
             Set rs = New Recordset
             rs.Open "SELECT * FROM cptablas WHERE tabla='43' AND codigo='" & cBloque1 & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs.EOF Then
                cdscbloqu1 = IIf(IsNull(rs!descripcion), "", Trim(rs!descripcion))
             End If
             rs.Close
          End If
          If cBloque2 <> "" Then
             Set rs = New Recordset
             rs.Open "SELECT * FROM cptablas WHERE tabla='43' AND codigo='" & cBloque2 & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs.EOF Then
                cdscbloqu2 = IIf(IsNull(rs!descripcion), "", Trim(rs!descripcion))
             End If
             rs.Close
          End If
          Set rs = New Recordset
          rs.Open "SELECT * FROM cptablas WHERE tabla='27' AND codigo='" & cCarest & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If Not rs.EOF Then
             cOrden = IIf(IsNull(rs!campo), "", Trim(rs!campo))
          End If
          rs.Close
          Set rs = New Recordset
          rs.Open "SELECT * FROM cpcCosto WHERE codigo='" & CCosto & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If Not rs.EOF Then
             cDscCosto = IIf(IsNull(rs!descripcion), "", Trim(rs!descripcion))
          End If
          rs.Close
          '-VALIDA ASISTENCIA
          FechaB = Format(TxDia, "yyyymmdd")
            dfentra = DateValue(TxDia)
            dfsale = DateValue(TxDia)
          c_turno = "X": p0_ing = "": p0_sal = "": p0_ing1 = "": p0_sal1 = "" 'HORARIO ORIGINAL
          p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = ""
          m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
          nHorasP = 0: nValorH = 0: nValorM = 0: mdscto = 0: nFaltaH = 0: nFaltaM = 0: F_falta = False: nmotpersonal = 0: nComision = 0
          cvalor = "": cAuPrg = "": cDiahora = "": Ctipo = "": choraini = "": chorafin = "": cDescuenta = "": cMantGuardia = "N": c_a = ""
                    
          '-Tolerancia x LACTANCIA (I)ngreso/(S)alida
          '=======================
          nLacta = 0
          If cAplica <> "N" Then
             nLacta = IIf(nSonLacta = 1, 1, 2)
          End If
          '' MsgBox "Aplica:" + cAplica + " [" + Trim(Str(nLacta)) + "] Dia:" + Format(primer + z, "dd/mm/yyyy") + " en " + Format(TxIniLac, "dd/mm/yyyy") + " - " + Format(TxFinLac, "dd/mm/yyyy")
          
          '-HORARIO prog en dia -**
          sql = "SELECT * FROM CpHor" & c_ano & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & FechaB & "'"
          Set rs = New Recordset
          rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If rs.EOF Then
                If Trim(cGrpGen) <> "" Then
                   rs.Close
                   sql = "SELECT * FROM CpHor" & c_ano & " WHERE grupo='" & cGrpGen & "' AND fecha ='" & FechaB & "'"
                   Set rs = New Recordset
                   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                End If
          End If
          If Not rs.EOF Then
                cvalor = IIf(IsNull(rs!Horario), "", rs!Horario)
                Sq0 = "SELECT * FROM CpHorarioNew  WHERE Codigo='" & cvalor & "'"
                Set rs_t = New Recordset
                rs_t.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                If rs_t.EOF Then
                   rs_t.Close
                   Sq0 = "SELECT * FROM CpHorarios  WHERE Codigo='" & cvalor & "'"
                   Set rs_t = New Recordset
                   rs_t.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                End If
                If Not rs_t.EOF Then
                   c_turno = IIf(IsNull(rs_t!turno), " ", Trim(rs_t!turno))   'NHCP   2dias
                   p_ing = IIf(IsNull(rs_t!ingreso), "", rs_t!ingreso)
                   If rs_t!Horas < "1800" Then                               '========>
                      p_ing1 = IIf(IsNull(rs_t!Ingreso1), "", rs_t!Ingreso1)
                      p_sal = IIf(IsNull(rs_t!salida), "", rs_t!salida)
                   End If
                   p_sal1 = IIf(IsNull(rs_t!salida1), "", rs_t!salida1)
                                      
                   '-horario ORIGINAL
                   '=================
                   p0_ing = IIf(IsNull(rs_t!ingreso), "", rs_t!ingreso)
                   ''If rs_t!Horas <> "2400" And rs_t!Horas <> "1800" Then                               '=======>
                      p0_sal = IIf(IsNull(rs_t!salida), "", rs_t!salida)
                      p0_ing1 = IIf(IsNull(rs_t!Ingreso1), "", rs_t!Ingreso1)
                   ''End If
                   p0_sal1 = IIf(IsNull(rs_t!salida1), "", rs_t!salida1)
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
                   p_horas = IIf(IsNull(rs_t!Horas), "0000", rs_t!Horas)
                   nValorH = Val(Left(p_horas, 2))
                   nValorM = Val(Right(p_horas, 2))
                   nFaltaH = Val(Left(p_horas, 2))
                   nFaltaM = Val(Right(p_horas, 2))
                   nHorasP = nHorasP + nValorH + (nValorM / 60)     'acumula
                   c2dias = "N"
                   If InStr("NHCP", c_turno) > 0 Then
                      c2dias = "S"
                   End If
                   
                End If
                rs_t.Close
          End If
          rs.Close
             
          'If InStr("ZZ0608", cTarjeta) > 0 Then
          '   MsgBox (cTarjeta + "(" + c2dias + "): " + p_ing + "-" + p_sal + ";" + p_ing1 + "-" + p_sal1)
          'End If
             
          'AUTORIZACIONES: Motivos Personales, Comisiones Serv - excluye  LVR
          cDiahora = "X"
          cMantGuardia = "N"
          sql = "SELECT * FROM CpAutori" & c_ano & " WHERE (anulado is NULL) and   tarjeta='" & cTarjeta & "' and FechaIni<='" + FechaB + "' and FechaFin>='" + FechaB + "'"
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
                  Set rs_t = New Recordset
                  rs_t.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  If Not rs_t.EOF Then
                     cDescuenta = IIf(rs_t!DESCUENTO = "", "N", Trim(rs_t!DESCUENTO))
                     cMantGuardia = IIf(IsNull(rs_t!MantGuardia) Or Trim(rs_t!MantGuardia) = "", "N", rs_t!MantGuardia)
                  End If
                  rs_t.Close
                  '-Aplica Autorización Entrada(I)/(S)alida
                  If Ctipo = "I" Then
                     m_ing = choraini
                  ElseIf Ctipo = "S" Then
                         m_sal = chorafin
                  End If
               End If
               rs.MoveNext
          Loop
          rs.Close
             
          F_falta = False
          i = Day(TxDia)
          '-MARCAS del dia       fecha  hora =>  hhmm
          If Val(p_horas) > 0 Then              '-------->2
             Sq0 = "SELECT * FROM CpMarca" & c_ano & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & FechaB & "' ORDER BY fecha,hora"
             p_Sale = ""
             If cDiahora <> "D" Then
                If c2dias = "S" Then
                   '-Ubica marcas 2dias Actual-siguiente caso: turno t/n, n/d
                   fecha1 = Format(DateValue(TxDia) + 1, "yyyymmdd")
                   Sq0 = "SELECT * FROM CpMarca" & c_ano & " WHERE tarjeta='" & cTarjeta & "' AND fecha >='" & FechaB & "' AND fecha <='" & fecha1 & "' ORDER BY fecha,hora"
                   cdiaant = "S"
                   p_Sale = p_sal
                   If Val(p_sal1) > 0 Then
                      p_Sale = p_sal1         'Corrige fin turno
                   End If
                End If
             End If
             
             '-ASIGNA marcas basado en Horario   I - S  /I1 - S1
           ' If InStr("A0255", cTarjeta) > 0 Then
           '    MsgBox sq0
           ' End If
             
             Set rs_t = New Recordset
                rs_t.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                If Not rs_t.EOF Then
                   d_fentra = rs_t!fecha         '========> 1era marca del servidor para turno programado
                   Do While Not rs_t.EOF
                      cInicio = p_ing
                      cFin = rs_t!hora
                      dfsale = rs_t!fecha
                      If Val(m_ing) = 0 Then
                         'If (nTotMints(cInicio, cFin) >= -60 And nTotMints(cInicio, cFin) <= 60) And Day(dfsale) = i Then
                         If Day(dfsale) = i Then
                             m_ing = rs_t!hora
                            'If InStr("49911_51245_59103_59172_A0255_E0001_T0057_ZZ207", cTarjeta) > 0 Then
                            '    MsgBox "1-m_ing-" + Format(dfsale, "dd/mm/yyyy") + ":" + rs_t!hora
                            'End If
                         Else
                             m_sal = rs_t!hora
                            'If InStr("49911_51245_59103_59172_A0255_E0001_T0057_ZZ207", cTarjeta) > 0 Then
                            '    MsgBox "2-m_sal-" + Format(dfsale, "dd/mm/yyyy") + ":" + rs_t!hora
                            'End If
                         End If
                      ElseIf Val(m_ing) > 0 And (Val(rs_t!hora) - Val(m_ing) < 60) And c2dias = "S" And i = Day(dfsale) Then
                          'Doble Ingreso¿2dias?         ============>
                          m_ing1 = rs_t!hora
                            'If InStr("49911_51245_59103_59172_A0255_E0001_T0057_ZZ207", cTarjeta) > 0 Then
                            '    MsgBox "3-m_ing1-" + Format(dfsale, "dd/mm/yyyy") + ":" + rs_t!hora
                            'End If
                    
                      ElseIf i = Day(dfsale) And Val(m_ing) > 0 And (Val(rs_t!hora) - Val(m_ing) < 60 And Val(rs_t!hora) < Val(p_sal) And c2dias = "N") Then
                           'Doble Ingreso¿?
                           m_ing1 = rs_t!hora
                             'If InStr("49911_51245_59103_59172_A0255_E0001_T0057_ZZ207", cTarjeta) > 0 Then
                             '   MsgBox "4-m_ing1-" + Format(dfsale, "dd/mm/yyyy") + ":" + rs_t!hora
                             'End If
                     
                      ElseIf i = Day(dfsale) And Val(p_ing1) > 0 And Val(m_ing1) = 0 And (Val(rs_t!hora) - Val(p_ing1) < 10) Then
                             '-ing Turno partido
                             m_ing1 = rs_t!hora
                            'If InStr("49911_51245_59103_59172_A0255_E0001_T0057_ZZ207", cTarjeta) > 0 Then
                            '    MsgBox "5-m_ing1-" + Format(dfsale, "dd/mm/yyyy") + ":" + rs_t!hora
                            'End If
                     
                      ElseIf i = Day(dfsale) And Val(p_ing1) = 0 And Val(m_ing1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then
                              m_ing1 = rs_t!hora
                            'If InStr("49911_51245_59103_59172_A0255_E0001_T0057_ZZ207", cTarjeta) > 0 Then
                            '    MsgBox "6-m_ing1-" + Format(dfsale, "dd/mm/yyyy") + ":" + rs_t!hora
                            'End If
                    
                      ElseIf (Val(m_sal) = 0 And (Val(rs_t!hora) - Val(p_sal) >= 0)) Or (Val(m_sal) = 0 And Val(p_sal) > Val(rs_t!hora)) Then
                             '07001900-19000700   0701 0738x 0740    1352 0712
                             m_sal = rs_t!hora
                            'If InStr("49911_51245_59103_59172_A0255_E0001_T0057_ZZ207", cTarjeta) > 0 Then
                            '    MsgBox "7-m_ing1-" + Format(dfsale, "dd/mm/yyyy") + ":" + rs_t!hora
                            'End If
                   
                     ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) > 0 And Val(m_sal1) = 0 And (Val(rs_t!hora) - Val(p0_sal1) >= 0) Then
                             '-sal Turno partido  1300-1900/1900-0700
                             m_sal1 = rs_t!hora
                            'If InStr("49911_51245_59103_59172_A0255_E0001_T0057_ZZ207", cTarjeta) > 0 Then
                            '    MsgBox "8-m_sal1-" + Format(dfsale, "dd/mm/yyyy") + ":" + rs_t!hora
                            'End If
                          
                     '-Check carga marcaciones sin valor para turno
                      ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then     'And Val(m_ing1) > 0
                             m_sal1 = rs_t!hora
                             'If InStr("49911_51245_59103_59172_A0255_E0001_T0057_ZZ207", cTarjeta) > 0 Then
                             '   MsgBox "3-m_sal1-" + Format(dfsale, "dd/mm/yyyy") + ":" + rs_t!hora
                             'End If
                            
                      End If
                      rs_t.MoveNext
                   Loop
                End If
                rs_t.Close
              
            '    If InStr("A0255", cTarjeta) > 0 Then
            '       MsgBox (cTarjeta + ": " + m_ing + "-" + m_sal + ";" + m_ing1 + "-" + m_sal1)
            '    End If
            
               '-Evalua Marcas INGRESO
                wvx = ""
                If (c_turno <> "N" And c_turno <> "C" And Not f_cierra) Or f_cierra Then        '-------> (A)
                
                If cDiahora = "D" And cDescuenta = "S" Then
                   F_falta = True     'FALTA ¿JUSTIFICADA?
                   vwx = "1"
                ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And Val(p_ing1) > 0 And p_sal <> p_ing1 Then  '======>
                       'Caso Horario:2000-0800   marca 0805-2010   O marca 0749-1404 1930-0820 ó  0803-2015/-0824
                       '-07001900-19000700 marca 0701 -07180719 & 1252,1900- ,07
                       ' 13001900-19000700       12521905- ,0715
                       F_falta = True
                       vwx = "&"
                
                ElseIf (Val(m_ing) + Val(m_ing1) = 0) Then
                        '-NO MARCO Ing,   valida vs autoriza
                        If cDiahora = "H" And Ctipo = "I" Then
                           cInicio = p_ing
                           cFin = choraini
                           If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                              F_falta = True
                              cDescuenta = "S"
                              wvx = "2"
                           End If
                        ElseIf cDiahora = "D" Then
                               If cDescuenta = "S" Then
                                  F_falta = True
                                  wvx = "3"
                               End If
                        Else   'TO CHECK
                           F_falta = True
                           wvx = "4"
                        End If
                ElseIf (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                        'Marco tarde, check Autoriza H al ingreso
                        If cDiahora = "H" And Ctipo = "I" Then
                           cInicio = p_ing
                           cFin = choraini
                           If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                              F_falta = True
                              vwx = "5"
                           End If
                        End If
                ElseIf Val(m_ing) > 0 Then
                        '--- INSERTADO Excepciones ----09.08.17
                        If Val(m_ing1) = 0 Then
                           '-No hay doble marca
                           '  PI   PS            MI   MS     MI1  MS1
                           '- 2000-0800      :   0740-1401   FALTA
                           '                 :   2130-0835   FALTA Ingreso fuera de Tolerancia
                           '-                :   1750-0820   ok
                           '--               :   0741-2044             marcas fuera Turno (antes/despues)-> FALTA
                           '--               :   0812-1847        0900
                           '---  0700-1900 1900-0700  24Hras  :marcas   0715- / -0748
                           '---  1300-1900 1900-0700  18hrs   :marca    1252-   -0705
                           If (m_sal > m_ing And m_sal < p_ing) And p_horas <> 2400 Then  '¿2400?
                              '-trabajo Antes de Turno
                              F_falta = True
                              vwx = "6"
                           Else
                              ' If i = 31 Then
                              '    MsgBox ("[" + c2dias + "=> " + m_ing + "-" + m_sal + "]")
                              ' End If
                              If c2dias = "S" And ((m_sal > p_sal And d_fentra = dfsale) Or (m_sal < p_sal And d_fentra < dfsale)) Then
                                 '07001900-19000700   0716, -07580759
                                 '-No completo el turno
                                 F_falta = True
                                 vwx = "7"
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
                                 vwx = "8"
                              End If
                           End If
                        End If
                End If
                
                End If          '<------ (A)
                  
                  
                If f_cierra Then    '-------> (B)
                  
                '-Evalua MARCAS Salida
                If F_ALTA Then
                ElseIf (Val(m_sal) + Val(m_sal1) = 0) Then
                        '-NO marco  Sal   valida vs autoriza
                        If cDiahora = "H" And Ctipo = "S" Then
                           If (Val(p_sal) - Val(chorafin) > 0) Or (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                              F_falta = True
                              cDescuenta = "S"
                              vwx = "A"
                           End If
                        ElseIf cDiahora = "D" Then     'valida autoriza Dia Comp/ pierde
                               If cDescuenta = "S" Then
                                  F_falta = True
                                  vwx = "B"
                               End If
                        Else
                           F_falta = True
                           vwx = "C"
                        End If
                ElseIf Val(p_sal1) > 0 And p0_sal <> p0_ing1 Then    'p_horas<>"2400"
                        If Val(m_sal1) = 0 Then
                           F_falta = True
                           vwx = "D"
                        Else
                           If Val(p_sal1) > Val(m_sal1) Then
                              F_falta = True
                              vwx = "E"
                           End If
                        End If
                ElseIf Val(m_sal) > 0 And Val(p_sal) > Val(m_sal) Then
                        'Marco antes de, Autoriza H
                        If cDiahora = "H" And Ctipo = "S" Then    '====> And cTipo = "S"  ** TO CHECK **
                           cInicio = p_sal
                           cFin = chorafin
                           If Val(p_sal) - Val(chorafin) > 0 Then
                              F_falta = True
                              vwx = "F"
                           End If
                        ElseIf cDiahora = "X" Then
                               '----------cInicio = p_sal
                               If p_sal > m_sal And m_ing1 <> "" And m_sal1 = "" Then
                                  'Correcion caso   11/May/15
                                  '04769(20/05):0758-0800-2001
                                  '01951(12/05):0759-0800-1501
                                  cFin = m_ing1
                               Else
                                  cFin = IIf(m_sal1 <> "", m_sal1, m_sal)
                               End If
                               If Val(p_sal) - Val(cFin) > 0 Then
                                  F_falta = True
                                  vwx = "G"
                               End If
                        End If
                End If
                
                End If          '<------ (B)
                   
                   
                If Not F_falta Then
                   If cDiahora = "H" And Ctipo = "I" Then     '...1
                   ElseIf cDiahora <> "D" Then      '<===No Autorizaciones D  ...1
                         If (Val(m_ing) + Val(m_ing1) > 0) Then
                            '-Hay marcas, aplica tolerancia, calcula tardanza
                            If cDiahora = "H" And Ctipo = "I" Then
                               cInicio = p_ing
                               cFin = choraini
                               nmintarde = CalcMints(cInicio, cFin)
                            Else
                               nmintarde = CalcMTard(p_ing, m_ing)
                            End If
                            If nmintarde > 30 Then
                                F_falta = True
                                vwx = "H"
                            End If
                            nmintarde = 0
                         End If
                   End If
                End If
                
                'If InStr("A0255", cTarjeta) > 0 Then
                '   MsgBox (cTarjeta + ": " + p_ing + "-" + p_sal + ";" + p_ing1 + "-" + p_sal1 + " => " + m_ing + "-" + m_sal + ";" + m_ing1 + "-" + m_sal1 + " Falta (" + vwx + ")")
                'End If
          End If    'Hras Programadas / Trabajadas             '<--------2
       
       
          If F_falta Or (cDiahora = "D" And cMantGuardia = "N") Then         '------->3
          Else
             If Option0(1).Value = True Then    'Administrativo
                cOrden = CCosto
                cDscCargo = cDscCosto
             End If
             nson = 1
             If cBloque <> "" And cOrden <> "" Then
                If cBloque1 <> "" Then
                   F_add = True
                   Set rs_d = New Recordset
                   rs_d.Open "SELECT * FROM #tabla2 WHERE Bloque='" & cBloque1 & "' AND orden='" & cOrden & "'", Cnx, adOpenStatic, adLockOptimistic, adCmdText
                   If rs_d.EOF Then
                   Else
                      Do While Not rs_d.EOF
                         For i = 1 To 9
                            z = i - 1
                            If IsNull(rs_d.Fields(z + 10)) Or Trim(rs_d.Fields(z + 10)) = "" Then
                               F_add = False
                               nson = i
                               Exit For
                            End If
                         Next i
                         If Not F_add Then
                            Exit Do
                         End If
                         rs_d.MoveNext
                      Loop
                   End If
                   If F_add Then
                      rs_d.AddNew
                      rs_d!carEst = cCarest
                      rs_d!DscCargo = cDscCargo
                      rs_d!CCosto = CCosto
                      rs_d!Dsccosto = cDscCosto
                      rs_d!bloque = cBloque1
                      rs_d!Dscbloque = cdscbloqu1
                      rs_d!orden = cOrden
                      rs_d!ano = cano
                      rs_d!mes = cmes
                      rs_d!Dia = TxDia
                      rs_d!Campo1 = cNombre
                      rs_d.Update
                   Else
                      rs_d.Fields(nson + 10 - 1) = cNombre
                      rs_d.Update
                   End If
                   rs_d.Close
                End If
                If cBloque2 <> "" Then
                   F_add = True
                   Set rs_d = New Recordset
                       rs_d.Open "SELECT * FROM #tabla2 WHERE Bloque='" & cBloque2 & "' AND orden='" & cOrden & "'", Cnx, adOpenStatic, adLockOptimistic, adCmdText
                   If rs_d.EOF Then
                   Else
                      Do While Not rs_d.EOF
                         For i = 1 To 9
                            z = i - 1
                            If IsNull(rs_d.Fields(z + 10)) Or Trim(rs_d.Fields(z + 10)) = "" Then
                               F_add = False
                               nson = i
                               Exit For
                            End If
                         Next i
                         If Not F_add Then
                            Exit Do
                         End If
                         rs_d.MoveNext
                      Loop
                   End If
                   If F_add Then
                      rs_d.AddNew
                      rs_d!carEst = cCarest
                      rs_d!DscCargo = cDscCargo
                      rs_d!CCosto = CCosto
                      rs_d!Dsccosto = cDscCosto
                      rs_d!bloque = cBloque2
                      rs_d!Dscbloque = cdscbloqu2
                      rs_d!orden = cOrden
                      rs_d!ano = cano
                      rs_d!mes = cmes
                      rs_d!Dia = TxDia
                      rs_d!Campo1 = cNombre
                      rs_d.Update
                   Else
                      rs_d.Fields(nson + 10 - 1) = cNombre
                      rs_d.Update
                   End If
                   rs_d.Close
                End If
             End If
          End If           '<-------3
          
          
       End If               '<------------1
       r_d.MoveNext
    Loop
    'Cnx.CommitTrans
    r_d.Close
End Sub



Private Sub CmdHuerfanos_Click()
    '-Directorio Diario PROGRAMADO - Asistencial
    If Not IsDate(TxDia) Then
       MsgBox "No indico fecha para analizar.. ", vbExclamation, "[IMPORTANTE]"
       Exit Sub
    End If
    Cnx.BeginTrans
    On Error GoTo crea1
    Cnx.Execute "DELETE  FROM #tabla2"
    GoTo sigue1
crea1:
    Cnx.Execute ("SELECT * Into #tabla2  FROM CpTmpDirect WHERE 1 = 2")
sigue1:
    Cnx.CommitTrans
    nson = 0: wdato = 0: npersout = 0
    fecha1 = Format(TxDia, "yyyymmdd")
    
    cmes = Mid(fecha1, 5, 2)
    nposicion = Day(TxDia) - 1
    '-check Noche Dia anterior
    fechaAnt = Format(DateValue(TxDia) - 1, "yyyymmdd")
    canoAnt = Left(fechaAnt, 4)
    cmesant = Mid(fechaAnt, 5, 2)
    ndiaAnt = Val(Right(fechaAnt, 2)) - 1
    Set r_d = New Recordset
    sql = "SELECT  * FROM CpDirectorio WHERE ano='" & c_ano & "' AND mes='" & cmes & "' AND FLAG ='0' " & _
          "ORDER BY Tarjeta "
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not r_d.EOF
       cNombre = "":  CCosto = "": cDscCosto = "": cCarest = "":  cDscCargo = "": cBloque = "": cBloque1 = "": cBloque2 = "": cdscbloqu1 = "": cdscbloqu2 = "": cOrden = ""
       CmdPrint.Caption = r_d!tarjeta + " - " + IIf(IsNull(r_d.Fields(nposicion + 4)), "", Trim(r_d.Fields(nposicion + 4)))
       cTurno = IIf(IsNull(r_d.Fields(nposicion + 4)), "", r_d.Fields(nposicion + 4))
       If InStr("*LVR", Trim(cTurno)) > 0 Or IsNull(cTurno) Or IsEmpty(cTurno) Then
          '-En: Descanso/Licencia/Vacaciones/Desc RX/No prog
          
          'If InStr("51017_ZZzzz", r_d!tarjeta) > 0 Then
          '   MsgBox "paso:" + r_d!tarjeta + "->[" + cTurno + "]"
          'End If
          
          nPersonal = nPersonal + 1
          cTarjeta = r_d!tarjeta
          cano = r_d!ano
          cmes = r_d!mes
          '-verifica si hay Marca
          Set rs = New Recordset
          sq_l = "SELECT * FROM CpMarca" + cano + " WHERE  fecha='" + fecha1 + "' and tarjeta='" + cTarjeta + "'"
          rs.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          cCadMarca = ""
          Do While Not rs.EOF
             cCadMarca = cCadMarca + rs!hora
             rs.MoveNext
          Loop
          rs.Close
          
           'If InStr("51017_ZZzzz", r_d!tarjeta) > 0 Then
           '  MsgBox "paso1:" + cTarjeta + "->diant:" + cTurnoAnt + "[" + cCadMarca + "]"
           'End If
          
          If Len(cCadMarca) > 0 Then
             wdato = wdato + 1
             '-verifica Dia ANTERIOR noche CHPN
             cTurnoAnt = ""
             Set rs = New Recordset
             rs.Open "SELECT * FROM CpDirectorio WHERE ano='" + canoAnt + "' and mes='" & cmesant & "' and tarjeta='" + cTarjeta + "'", Cnx, adOpenStatic, adLockReadOnly
             If Not rs.EOF Then
                cTurnoAnt = IIf(IsNull(r_d.Fields(ndiaAnt + 4)), "", r_d.Fields(ndiaAnt + 4))
             End If
             rs.Close
             cDiaOUT = "S"
             If Trim(cTurnoAnt) > "" And InStr("CHPN", Left(cTurnoAnt, 1)) > 0 Then
                '-Verifica existe marcas del dia OUT
                For i = 1 To Len(cCadMarca) / 4                  'hhmmhhmm
                   cMarca = Mid$(cCadMarca, (i - 1) * 4 + 1, 4)
                   If cMarca <= "1100" Then
                      cDiaOUT = "N"
                   Else
                     cDiaOUT = "S"
                   End If
                Next
             End If
             
             'If InStr("51017_zzzzz", cTarjeta) > 0 Then
             '   MsgBox "paso2 " + cTarjeta + ">" + cTurnoAnt + "<" + cCadMarca + "[" + cDiaOUT + "]"
             'End If
             
             If cDiaOUT = "S" Then
                '-coge de Directorio
                cCarest = IIf(IsNull(r_d!carEst), "", r_d!carEst)
                Set rs = New Recordset
                sq_l = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta='" & cTarjeta & "'"
                rs.Open sq_l, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                If Not rs.EOF Then
                   cNombre = Trim(rs!apelpat) + "," + Left(rs!nombres, 1) + ". ;"
                   CCosto = rs!CCosto
                   '-vacio coge parametro en maestra
                   cCarest = IIf(IsEmpty(cCarest), IIf(IsNull(rs!carEst), "", rs!carEst), cCarest)
                   If Len(cNombre) > 17 Then
                      cNombre = Left(cNombre, 17)
                   End If
                End If
                rs.Close
             
                Set rs = New Recordset
                rs.Open "SELECT * FROM Cptablas WHERE tabla='27' and codigo='" & cCarest & "'", Cnx, adOpenStatic, adLockReadOnly
                If Not rs.EOF Then
                   cDscCargo = IIf(IsNull(rs!descripcion), "", Left(rs!descripcion, 30))
                End If
                rs.Close
          
                Set rs = New Recordset
                rs.Open "SELECT * FROM cptablas WHERE tabla='31' AND codigo='" & cTurno & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
                If Not rs.EOF Then
                   cBloque = IIf(IsNull(rs!campo), "", Trim(rs!campo))
                   If Len(cBloque) > 1 Then
                      cBloque1 = Left(cBloque, 1)
                      cBloque2 = Right(cBloque, 1)
                   Else
                      cBloque1 = cBloque
                   End If
                Else
                   cBloque = "0": cBloque1 = "0"
                   cdscbloqu1 = "HORARIO no Programado"
                End If
                rs.Close
                If Val(cBloque1) > 0 Then  ' <> ""
                   Set rs = New Recordset
                   rs.Open "SELECT * FROM cptablas WHERE tabla='43' AND codigo='" & cBloque1 & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
                   If Not rs.EOF Then
                      cdscbloqu1 = IIf(IsNull(rs!descripcion), "", Trim(rs!descripcion))
                   End If
                   rs.Close
                End If
                If cBloque2 <> "" Then
                   Set rs = New Recordset
                   rs.Open "SELECT * FROM cptablas WHERE tabla='43' AND codigo='" & cBloque2 & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
                   If Not rs.EOF Then
                      cdscbloqu2 = IIf(IsNull(rs!descripcion), "", Trim(rs!descripcion))
                   End If
                   rs.Close
                End If
             
                Set rs = New Recordset
                rs.Open "SELECT * FROM cptablas WHERE tabla='27' AND codigo='" & cCarest & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
                If Not rs.EOF Then
                   cOrden = IIf(IsNull(rs!campo), "", Trim(rs!campo))
                End If
                If Trim(cOrden) = "" Then
                   MsgBox "[" + cTarjeta + "] Cargo Est. (" + cCarest + ") No definido para ASISTENCIALES (Tab.Aux 27)"
                End If
                rs.Close
                Set rs = New Recordset
                rs.Open "SELECT * FROM cpcCosto WHERE codigo='" & CCosto & "'", Cnx, adOpenStatic, adLockReadOnly, adCmdText
                If Not rs.EOF Then
                   cDscCosto = IIf(IsNull(rs!descripcion), "", Trim(rs!descripcion))
                End If
                rs.Close
             
                'If InStr("51017_zzzzz", cTarjeta) > 0 Then
                '   MsgBox "paso3 " + cTarjeta + "[" + cDiaOUT + "-" + cBloque + "-" + cOrden + "]"
                'End If
 
                If cBloque <> "" And cOrden <> "" Then
                   If cBloque1 <> "" Then
                      F_add = True
                      nson = 1
                      Set rs_d = New Recordset
                      rs_d.Open "SELECT * FROM #tabla2 WHERE Bloque='" & cBloque1 & "' AND orden='" & cOrden & "'", Cnx, adOpenStatic, adLockOptimistic, adCmdText
                      If Not rs_d.EOF Then
                         Do While Not rs_d.EOF
                            For i = 1 To 9
                               z = i - 1
                               If IsNull(rs_d.Fields(z + 10)) Or Trim(rs_d.Fields(z + 10)) = "" Then
                                  nson = i
                                  F_add = False
                                  Exit For
                               End If
                            Next i
                            If Not F_add Then
                               Exit Do
                            End If
                            rs_d.MoveNext
                         Loop
                      End If
                   
                      If F_add Then
                         nGrabo = nGrabo + 1
                         rs_d.AddNew
                         rs_d!carEst = cCarest
                         rs_d!DscCargo = cDscCargo
                         rs_d!CCosto = CCosto
                         rs_d!Dsccosto = cDscCosto
                         rs_d!bloque = cBloque1
                         rs_d!Dscbloque = cdscbloqu1
                         rs_d!orden = cOrden
                         rs_d!ano = cano
                         rs_d!mes = cmes
                         rs_d!Dia = TxDia
                         rs_d!Campo1 = cNombre
                         rs_d.Update
                      Else
                         rs_d.Fields(nson - 1 + 10) = cNombre
                         rs_d.Update
                      End If
                      rs_d.Close
                      npersout = npersout + 1
                   End If
                   If cBloque2 <> "" Then
                      F_add = True
                      nson = 1
                      Set rs_d = New Recordset
                      rs_d.Open "SELECT * FROM #tabla2  WHERE Bloque='" & cBloque2 & "' AND orden='" & cOrden & "'", Cnx, adOpenStatic, adLockOptimistic, adCmdText
                      If Not rs_d.EOF Then
                         Do While Not rs_d.EOF
                            For i = 1 To 9
                               z = i - 1
                               If IsNull(rs_d.Fields(z + 10)) Or Trim(rs_d.Fields(z + 10)) = "" Then
                                  nson = i
                                  F_add = False
                                  Exit For
                               End If
                            Next i
                            If Not F_add Then
                               Exit Do
                            End If
                            rs_d.MoveNext
                         Loop
                      End If
                    
                      If F_add Then
                         rs_d.AddNew
                         rs_d!carEst = cCarest
                         rs_d!DscCargo = cDscCargo
                         rs_d!CCosto = CCosto
                         rs_d!Dsccosto = cDscCosto
                         rs_d!bloque = cBloque2
                         rs_d!Dscbloque = cdscbloqu2
                         rs_d!orden = cOrden
                         rs_d!ano = cano
                         rs_d!mes = cmes
                         rs_d!Dia = TxDia
                         rs_d!Campo1 = cNombre
                         rs_d.Update
                      Else
                         rs_d.Fields(nson - 1 + 10) = cNombre
                         rs_d.Update
                      End If
                      rs_d.Close
                      npersout = npersout + 1
                   End If
                End If   'Bloque/Orden
             End If 'DiaOUT
          End If    'cadena de marcas
       End If     'no programado
       r_d.MoveNext
    Loop
    r_d.Close
    
    Beep 500, 350
    Beep 500, 350
    If wdato > 0 Then
       MsgBox "Evaluados " & Trim(Str(wdato)) + " No prog: " + Trim(Str(npersout)), vbExclamation, "[SUMARIO]"
       If npersout > 0 Then
          nTipRpt = 7.5
          unafecha = TxDia
          InfoCpRS.Left = 0: InfoCpRS.Top = 0
          InfoCpRS.Refresh
          InfoCpRS.Show
       End If
    End If
   
    Call Limpiacampos
    CmdPrint.Caption = ""
    Exit Sub
   
End Sub



Private Sub cmdsalir_Click()
   Set r_d = Nothing
   Set rs_d = Nothing
   Set rs = Nothing
   Set rs_t = Nothing
   Me.Hide
   Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set r_d = Nothing
   Set rs_d = Nothing
   Set rs = Nothing
   Set rs_t = Nothing
   Me.Hide
   Unload Me
End Sub

