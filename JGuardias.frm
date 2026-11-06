VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form JGuardias 
   Caption         =   "Jefaturas de Guardia - Programadas"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "JGuardias.frx":0000
   ScaleHeight     =   5520
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton mnuFicImpSelec 
      Height          =   495
      Left            =   9600
      Picture         =   "JGuardias.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Selec.Impresora"
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   1095
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Width           =   8775
      Begin VB.TextBox Txtarifa 
         Height          =   285
         Left            =   7800
         MaxLength       =   9
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox TxNivGuardia 
         Height          =   285
         Left            =   8160
         MaxLength       =   2
         TabIndex        =   19
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Txdscturno 
         Height          =   285
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   17
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox Txturno 
         Height          =   285
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxDia 
         Height          =   285
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Txtrabajador 
         Height          =   285
         Left            =   1920
         MaxLength       =   40
         TabIndex        =   12
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox Txtarjeta 
         Height          =   285
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Tarifa"
         Height          =   255
         Left            =   7320
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Niv. Guardia"
         Height          =   255
         Left            =   7080
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Turno"
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Dia"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tarjeta"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   600
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
      Left            =   9480
      Picture         =   "JGuardias.frx":80A44
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir"
      Top             =   3840
      Width           =   765
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   2655
      Begin VB.CommandButton CmdClear 
         Height          =   735
         Left            =   1680
         Picture         =   "JGuardias.frx":8130E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Otra Consulta"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdcarga 
         Height          =   735
         Left            =   360
         Picture         =   "JGuardias.frx":81BD8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ubicar Registros"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   9480
      Picture         =   "JGuardias.frx":81DB3
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cerrar"
      Top             =   4800
      Width           =   735
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2640
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   195
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Jefaturas del Periodo"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "JGuardias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                               Dim DD As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim r_d As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim Sw, f_Fer As Boolean
Dim nposicion, ndifdias As Integer
Dim sql, SQ1, cano, cmes, c_a, cTarjeta, cTurno, cDscTurno, cnivGuardia, ccadena, cPatron As String
Dim nTarifa As Double
Dim unafecha As Date
Private Sub Form_Load()
  JGuardias.Caption = "Jefaturas de Guardia - Programadas " & genaño
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
End Sub
Private Sub Cbomes_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboMes.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbomes_LostFocus()
   If Not IsEmpty(CboMes.Text) Then
      Cbomes_KeyPress (13)
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
            TxNombMes.Text = Trim(Left(CboMes.List(j), nposicion - 1))
            unafecha = "01/" + Trim(Txmes.Text) + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)            'primer/ultimo
         End If
      End If
      Exit Sub
   End If
End Sub
Private Sub cmdcarga_Click()
   If Trim(Txmes.Text) = "" Then
       MsgBox "No Selecciono mes ", vbExclamation
       CboMes.SetFocus
      Exit Sub
   End If
   Set DD = New Recordset
   With DD
     .Fields.Append "Dia", adChar, 2
     .Fields.Append "turno", adChar, 2       ' cptabla 31
     .Fields.Append "Dscturno", adChar, 30
     .Fields.Append "Tarjeta", adChar, 5
     .Fields.Append "ApelPat", adChar, 20
     .Fields.Append "ApelMat", adChar, 20
     .Fields.Append "Nombres", adChar, 20
     .Fields.Append "NivGuardia", adChar, 2     ' cptarifas ->niveles
     .Fields.Append "tarifa", adDouble
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
          Case 0
          '.Caption = "Dia"
          .Width = Me.TextWidth(String(4, "S"))
          Case 1
          '.Caption =turno
          .Width = Me.TextWidth(String(3, "S"))
          Case 2
          .Caption = "Descripcion"
          .Width = Me.TextWidth(String(30, "S"))
          Case 3
          .Width = Me.TextWidth(String(7, "S"))
          Case 4
          .Width = Me.TextWidth(String(12, "S"))
          Case 5
          .Width = Me.TextWidth(String(12, "S"))
          Case 6
          .Width = Me.TextWidth(String(20, "S"))
          Case 7
           .Caption = "Cod.Niv."
          .Width = Me.TextWidth(String(8, "S"))
           Case 8
          .Width = Me.TextWidth(String(10, "S"))
          End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
  Next
  
 Cnx.Execute "DELETE  FROM CpTmpGuardia"
 Set rs_d = New Recordset
 SQ1 = "SELECT * FROM CpTmpGuardia"
 rs_d.Open SQ1, Cnx, adOpenStatic, adLockOptimistic, adCmdText
  
 '-carga  DIRECTORIO,  UPR "08" servicio  emergencia
  sql = "SELECT  * FROM CpDirectorio  WHERE ano='" & genaño & "' AND mes='" & Trim(Txmes.Text) & "'  AND CodUpr='08' AND  flag='0'" & _
        "ORDER BY Tarjeta, ano, mes "
  Set r_d = New Recordset
  r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  ccadena = ""
  For i = Day(primer) To Day(ultimo)
     ccadena = ccadena + "0"             ' 0xX =>  ' ','1/4','T'
  Next i
  If Not r_d.EOF Then
     Do While Not r_d.EOF
        cTarjeta = r_d!tarjeta
        For i = Day(primer) To Day(ultimo)
           z = i - 1
           cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", Trim(r_d.Fields(z + 4)))
           If cTurno = "*" Or cTurno = "" Then
           Else
           
           cmdcarga.Caption = cTarjeta + " " + Str(i)
           '-JEFATURA GUARDIA
           If InStr("1_2_3_4", cTurno) > 0 Or (InStr("D*_N*_H*_C*_P*_D3_N3", cTurno) > 0 And Len(cTurno) = 2) Then
              cPatron = Mid(ccadena, i, 1)
              If cPatron = "0" Or cPatron = "x" Then
                 f_Fer = False    ' verifica Feriado/domingo
                 '''If Weekday(unafecha + i - 1) = 1 Then
                 If UCase(Left(WeekdayName(Weekday(primer + z)), 3)) = "DOM" Then
                    f_Fer = True
                 Else
                    Set rs = New Recordset
                    rs.Open "SELECT * FROM Feriados ORDER BY codigo", Cnx0, adOpenStatic, adLockReadOnly
                    'rs.Open "SELECT * FROM CpTablasAux WHERE grupo='FER' ORDER BY codigo", Cnx, adOpenStatic, adLockReadOnly
                    rs.MoveFirst
                    While Not rs.EOF
                       If Len(Trim(rs!codigo)) = 6 Then       'codigo:mmddaa
                          c_a = Right(rs!codigo, 2)
                       ElseIf Len(Trim(rs!codigo)) = 4 Then   'Codigo:mmdd
                               c_a = Right(Year(unafecha), 2)
                       End If
                       If Val(Left(rs!codigo, 2)) = Val(Txmes.Text) And Val(Mid(rs!codigo, 3, 2)) = i And Right(Year(unafecha), 2) = c_a Then
                          f_Fer = True
                       End If
                       rs.MoveNext
                    Wend
                    rs.Close
                 End If
                 DD.AddNew
                 DD!Dia = Right("00" + Trim(Str(i)), 2)
                 DD!turno = cTurno
                 DD!tarjeta = cTarjeta
                 Set rs = New Recordset
                 sql = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & cTarjeta & "'"
                 rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                 caPat = "": caMat = "": cnomb = "": cnivGuardia = ""
                 If rs.EOF Then
                 Else
                    caPat = Trim(rs!apelpat)
                    caMat = Trim(rs!apelmat)
                    cnomb = Trim(rs!nombres)
                    cnivGuardia = IIf(IsNull(rs!NivGuardia), "", rs!NivGuardia)
                 End If
                 rs.Close
                 DD!apelpat = caPat
                 DD!apelmat = caMat
                 DD!nombres = cnomb
                 DD!NivGuardia = cnivGuardia
                 Set rs = New Recordset
                 rs.Open "SELECT * FROM Cptablas WHERE tabla='31' and codigo='" & cTurno & "'", Cnx, adOpenStatic, adLockReadOnly
                 cDscTurno = ""
                 If Not rs.EOF Then
                    cDscTurno = rs!descripcion
                 End If
                 rs.Close
                 DD!DscTurno = cDscTurno
                 Set rs = New Recordset
                 rs.Open "SELECT * FROM CpTarifas WHERE (baja is NULL)", Cnx, adOpenStatic, adLockReadOnly
                 nTarifa = 0#
                 Do While Not rs.EOF
                    If InStr(rs!Niveles, cnivGuardia) > 0 Then
                       If Trim(cTurno) = "1" Or Trim(cTurno) = "D*" Or Trim(cTurno) = "D3" Then
                          If f_Fer Then
                             nTarifa = rs!DiaFer
                          Else
                             nTarifa = rs!DiaOrd
                          End If
                       ElseIf InStr("4_N*_H*_N3_C*_P*", Trim(cTurno)) > 0 Then
                              If f_Fer Then
                                 nTarifa = rs!NocheFer
                              Else
                                 nTarifa = rs!NocheOrd
                              End If
                       End If
                       Exit Do
                    End If
                    rs.MoveNext
                 Loop
                 rs.Close
                 DD!Tarifa = nTarifa
                 DD.Update
              
                 rs_d.AddNew
                 rs_d!Dia = Right("00" + Trim(Str(i)), 2)
                 rs_d!turno = cTurno
                 rs_d!DscTurno = cDscTurno
                 rs_d!tarjeta = cTarjeta
                 rs_d!apelpat = caPat
                 rs_d!apelmat = caMat
                 rs_d!nombres = cnomb
                 rs_d!NivGuardia = cnivGuardia
                 rs_d!Tarifa = nTarifa
                 rs_d.Update
                 If cPatron = "0" Then
                    ccadena = Left(ccadena, i - 1) + "x" + Right(ccadena, Day(ultimo) - i)
                 ElseIf cPatron = "x" Then
                        ccadena = Left(ccadena, i - 1) + "X" + Right(ccadena, Day(ultimo) - i)
                 End If
              End If
           End If
           
           End If
        Next i
        r_d.MoveNext
     Loop
  End If
  r_d.Close
  cmdcarga.Caption = ""
  DataGrid1.Refresh
  Call Limpiacampos
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   TxDia = DataGrid1.Columns(0)
   Txturno = DataGrid1.Columns(1)
   Txdscturno = DataGrid1.Columns(2)
   Txtarjeta = DataGrid1.Columns(3)
   TxTrabajador = Left(Trim(DataGrid1.Columns(4)) + " " + Trim(DataGrid1.Columns(5)) + " " + Trim(DataGrid1.Columns(6)), 40)
   TxNivGuardia = DataGrid1.Columns(7)
   Txtarifa = DataGrid1.Columns(8)
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
   With DD      ' Adodc1.Recordset
      If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
         .Sort = .Fields(ColIndex).[Name] & " Desc"
      Else
         .Sort = .Fields(ColIndex).[Name] & " Asc"
      End If
   End With
   Call Limpiacampos
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
   If Txmes.Text = "" Then
      MsgBox "Defina..  Periodo... ", vbExclamation
      Exit Sub
   End If
   nTipRpt = 10
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Call Limpiacampos
   Cmdprint.Caption = ""
   Exit Sub
End Sub
Private Sub CmdClear_Click()
   DD.Close
   Txmes.Text = "": CboMes.Text = "": TxNombMes.Text = ""
   Call Limpiacampos
End Sub
Private Sub Limpiacampos()
   TxDia = "": Txturno = "": Txdscturno = "": TxNivGuardia = ""
   Txtarjeta = "": TxTrabajador = "": Txtarifa = ""
   'DataGrid1.SetFocus
End Sub
Private Sub cmdsalir_Click()
   Set r_d = Nothing
   Set rs_d = Nothing
   Set rs = Nothing
   Me.Hide
   Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set r_d = Nothing
   Set rs_d = Nothing
   Set rs = Nothing
   Me.Hide
   Unload Me
End Sub
