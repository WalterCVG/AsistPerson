VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form QueryPrograma 
   Caption         =   "Directorio Servidores ASISTENCIALES  (Programa SUSALUD)"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   8040
      Picture         =   "QueryPrograma.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Cerrar"
      Top             =   3720
      Width           =   735
   End
   Begin VB.CheckBox ChkAnual 
      Alignment       =   1  'Right Justify
      Caption         =   "Anual"
      Height          =   375
      Left            =   240
      TabIndex        =   31
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton CmdCarga1 
      Height          =   1095
      Left            =   1080
      Picture         =   "QueryPrograma.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CheckBox ChkXLS 
      Alignment       =   1  'Right Justify
      Caption         =   "XLS instalado"
      Height          =   195
      Left            =   4440
      TabIndex        =   29
      Top             =   3600
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
      Height          =   195
      Left            =   4680
      TabIndex        =   28
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox CboMes 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxNombMes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdClear 
      Height          =   615
      Left            =   8160
      Picture         =   "QueryPrograma.frx":0D8C
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Otra Consulta"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox TxUltDia 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      Height          =   195
      Left            =   7080
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Txpridia 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   3
      EndProperty
      Enabled         =   0   'False
      Height          =   195
      Left            =   5760
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxDni 
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      MaxLength       =   10
      TabIndex        =   23
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   8535
      Begin VB.TextBox TxCodUpr 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7680
         MaxLength       =   2
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CboUpr 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   6015
      End
      Begin VB.ComboBox CboCCosto 
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Width           =   6015
      End
      Begin VB.TextBox Txccosto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7680
         MaxLength       =   2
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Txtipopla 
         Enabled         =   0   'False
         Height          =   225
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CboTipoPla 
         Height          =   315
         Left            =   6120
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.ComboBox CboRegPlh 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox TxPlh 
         Enabled         =   0   'False
         Height          =   225
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "U.P.R."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ORGANO"
         Height          =   255
         Left            =   600
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Planilla"
         Height          =   255
         Left            =   5160
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Regimen Laboral"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdExportaXLS 
      Height          =   615
      Left            =   5880
      Picture         =   "QueryPrograma.frx":1656
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Exporta a Xls"
      Top             =   3480
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
      Height          =   195
      Left            =   3720
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Txdscccosto 
      Enabled         =   0   'False
      Height          =   195
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxhorasP 
      Enabled         =   0   'False
      Height          =   195
      Left            =   2880
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CmdCarga 
      Height          =   855
      Left            =   6960
      Picture         =   "QueryPrograma.frx":2298
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar Progreso1 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2760
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label_1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Lblrec 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LblRec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   3120
      Width           =   3615
   End
End
Attribute VB_Name = "QueryPrograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rD As New ADODB.Recordset
Dim Rsd As New ADODB.Recordset
Dim Criterio, f_iltro As String
Dim mFin, cTarjeta, cDni, csexo, cAplica, ccIng, CCsal, cHoraI, cHoraF, cCadenaC0, cCadenaCC, cCadenaDias As String
Dim nsonCC, nLacta, nSonLacta, nposicion, wdato, nreg, nson, nson1, nvez, Ncuenta0, ncuenta1 As Integer
Dim unafecha As Date

Private Sub Cbomes_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboMes.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbomes_LostFocus()
   If IsEmpty(CboMes.Text) Then
   Else
      Call Cbomes_KeyPress(13)
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
            TxNombMes = Trim(Left(CboMes.List(j), nposicion - 1))
            unafecha = "01/" + Txmes + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)
         End If
      End If
      Exit Sub
   End If
End Sub

Private Sub Form_Load()
  Lblrec.Caption = ""
  Me.Caption = "Directorio - programaciòn -" + genaño
  
  '-Meses
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
  Call Cbomes_KeyPress(13)
    
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpTablasAux WHERE grupo='REG' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboRegPlh.ListCount - 1
           If CboRegPlh.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboRegPlh.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  CboTipoPla.Clear
  CboTipoPla.AddItem "ASISTENCIAL"
  '--CboTipoPla.AddItem "ADMINISTRATIVO"
  CboTipoPla.ListIndex = 0
  
  Set rs = New Recordset
  '-rs.Open "SELECT * FROM CpCcosto ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  '-ORGANOS Asistenciales HABILITADOS
  rs.Open "SELECT * FROM CpTABLASAUX WHERE GRUPO='SUS' ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  cCadenaCC = " AND ccosto IN("
  cCadenaC0 = ""
  nsonCC = 0
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To CboCCosto.ListCount - 1
         If CboCCosto.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("Codigo") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        CboCCosto.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("Codigo")
        cCadenaC0 = cCadenaC0 + rs.Fields("codigo") + "_"
        nsonCC = nsonCC + 1
     End If
     rs.MoveNext
  Wend
  rs.Close
 
  For i = 1 To nsonCC
      cAplica = Mid(cCadenaC0, (i - 1) * 3 + 1, 2)
      If i < nsonCC Then
         cCadenaCC = cCadenaCC + "'" + cAplica + "',"
      Else
         cCadenaCC = cCadenaCC + "'" + cAplica + "')"
      End If
  Next i
  '-MsgBox cCadenaCC
 
  TxPlh = "":  CboRegPlh.Text = ""
  Txtipopla = "":  CboTipoPla.Text = ""
  Txccosto = "":  CboCCosto.Text = ""
  TxCodUpr = "": CboUpr.Text = ""
  
  '-mes anterior
  unafecha = DateValue("01/" + Right("0" + Trim(Str(genmes)), 2) + "/" + Trim(Str(genaño))) - 1
  Call ObtenerPrimerUltimoDia(unafecha)
  Txpridia = primer
  TxUltDia = ultimo
End Sub
Private Sub CboRegPlh_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboRegPlh.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboRegPlh_LostFocus()
  If IsEmpty(CboRegPlh.Text) Then
  Else
     j = CboRegPlh.ListIndex
     nposicion = InStr(CboRegPlh.List(j), ":")
     nlen = Len(CboRegPlh.List(j))
     TxPlh = Right(CboRegPlh.List(j), nlen - nposicion)
  End If
End Sub
Private Sub Cbotipopla_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboTipoPla.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbotipopla_LostFocus()
  If IsEmpty(CboTipoPla.Text) Then
  Else
     j = CboTipoPla.ListIndex
     Txtipopla = IIf(Left(CboTipoPla.List(j), 3) = "ASI", "0", "1")
  End If
End Sub
Private Sub Cboccosto_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboCCosto.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cboccosto_LostFocus()
  If IsEmpty(CboCCosto.Text) Then
  Else
     j = CboCCosto.ListIndex
     nposicion = InStr(CboCCosto.List(j), ":")
     nlen = Len(CboCCosto.List(j))
     Txccosto = Right(CboCCosto.List(j), nlen - nposicion)
     
     sql = "SELECT     tab1.CCOSTO, tab1.CODUPR, tab2.Descripcion " & _
     "FROM         CPCCostoUPR AS tab1 INNER JOIN " & _
     "                 CpUndPrgRol AS tab2 ON tab1.CODUPR = tab2.CodUPR " & _
     "WHERE CCOSTO='" & Txccosto + "'  ORDER BY tab2.descripcion "
     
     CboUpr.Clear
     Set rs = New Recordset
     rs.Open sql, Cnx, adOpenStatic, adLockReadOnly
     While Not rs.EOF
        ExisteYa = False
        For i = 0 To CboUpr.ListCount - 1
             If CboUpr.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR") Then
                ExisteYa = True
                Exit For
             End If
        Next i
        If Not ExisteYa Then
            CboUpr.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR")
        End If
        rs.MoveNext
     Wend
     rs.Close
  End If
End Sub

Private Sub CboUPR_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboUpr.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboUPR_LostFocus()
  If IsEmpty(CboUpr.Text) Then
  Else
     j = CboUpr.ListIndex
     nposicion = InStr(CboUpr.List(j), ":")
     nlen = Len(CboUpr.List(j))
     TxCodUpr = Right(CboUpr.List(j), nlen - nposicion)
  End If
End Sub

Private Sub cmdcarga_Click()
  
  If Val(Txmes) = 0 Then
     MsgBox "favor indicar mes de consulta..", vbCritical
     Exit Sub
  End If
  If Trim(Txtipopla) = "" Then
     MsgBox "No indica tipo de planilla..", vbCritical
     Exit Sub
  End If
  
  'If Trim(TxPlh) = "" And Trim(Txccosto) = "" And Trim(TxCodUpr) = "" Then
  '   MsgBox "Indicar parametros de consulta..", vbCritical, "[ ERROR ]"
  '   Exit Sub
  'End If

  Txpridia = primer
  TxUltDia = ultimo
  Lblrec.Caption = ""
  wdato = 0
  Criterio = "ano ='" & genaño & "' AND mes='" & Trim(Txmes) & "' and FLAG='" + Txtipopla + "'"
  f_iltro = ""
  If Txccosto = "" And TxCodUpr = "" Then
     Criterio = Criterio + cCadenaCC
  Else
     If Txccosto <> "" Then
       If Len(Criterio) > 0 Then
            Criterio = Criterio + " and  (ccosto='" & Txccosto & "')"
            f_iltro = f_iltro & ", C.Costo "
       Else
            Criterio = " (ccosto='" & Txccosto & "')"
            f_iltro = "C.Costo "
       End If
     End If
     If TxCodUpr <> "" Then
       If Len(Criterio) > 0 Then
            Criterio = Criterio + " and  (codupr='" & TxCodUpr & "')"
            f_iltro = f_iltro & ", UPR "
       Else
            Criterio = " (codupr='" & TxCodUpr & "')"
            f_iltro = "UPR "
       End If
     End If
  End If
  
  If Trim(TxPlh) <> "" Then
      Criterio = Criterio + "  and  (plh='" & TxPlh & "')"
      f_iltro = " PLH "
  End If
    
  If Criterio = "" Then
     MsgBox "Favor defina criterio ", vbExclamation
     Exit Sub
  End If
   
 '-----  MsgBox Criterio
 ' ndiasson = DateDiff("d", primer, ultimo) + 1
 
  
 Call CmdCarga21
 wTpoQuery = "7.7"
   
End Sub
Private Sub CmdClear_Click()
  Set rs = Nothing
  Set rRS = Nothing
  Set rD = Nothing
  CmdCarga.Caption = ""
  TxPlh = "":  CboRegPlh.Text = "":  Txtipopla = "":  CboTipoPla.Text = ""
  Txccosto = "":  CboCCosto.Text = ""
  TxCodUpr = "": CboUpr.Text = ""
  ChkXLS.Value = 1: ChkAnual.Value = 0
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
  Call Cbomes_KeyPress(13)

  TxDni = "":  Txdscccosto = ""
  Lblrec.Caption = ""
  wdato = 0
  
End Sub

Private Sub CmdCarga21()
   Lblrec.Caption = ""
   wdato = 0
   
   '-Directorio mensual
   Cnx.BeginTrans
   Cnx.Execute "DELETE  FROM CpDirPrograma"
   On Error GoTo Err7
   Cnx.Execute "DELETE  FROM #Tabla77"
   GoTo sigue7
Err7:
   Cnx.Execute ("SELECT * Into #Tabla77  FROM CpDirPrograma WHERE 1 = 2")
sigue7:
   Cnx.CommitTrans

   Set rRS = New Recordset
   sqRS = "SELECT * FROM #Tabla77"    'Directorio mensual -> export XLS
   rRS.Open sqRS, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   rRS.Close
   
   Set r_d = New ADODB.Recordset
   sqG = "SELECT * FROM CpDirectorio WHERE " & Criterio
   r_d.Open sqG, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If r_d.EOF Then
      Exit Sub
   End If
   nreg = r_d.RecordCount
  
   Progreso1.Value = 0
   Progreso1.Max = nreg
   Progreso1.Visible = True
   r_d.MoveFirst
   CmdCarga.Caption = Trim(Str(nreg))
   Cnx.BeginTrans
   Dim OpenForms
   v_an = 0
   diai = Day(Txpridia)    'Dias de inicio y cierre del periodo consultado
   diaf = Day(TxUltDia)    'Day(ultimo)
      
   Do While Not r_d.EOF         '=====> Servidores
   
      '-dar control a windows y permitir multihilo
      v_an = v_an + 1
      If v_an Mod 10 = 0 Then     'if loop has repeated
         OpenForms = DoEvents 'Yield to operating system
      End If
   
      With Progreso1
        If .Value < .Max Then
           .Value = .Value + 1
        End If
      End With
      capelPat = "": capelMat = "": cnombres = "": csexo = "":  wdsccosto = "": wdscUpr = ""
      cdescEspec = "SIN ESPECIALIDAD": ccodreg = "": CICR = "": cCadenaDias = ""
      cTarjeta = Trim(r_d!tarjeta)     '- servidor
      ccCosto = IIf(IsNull(r_d!CCosto), "", Trim(r_d!CCosto))
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
         ccodreg = Right("0000" + Trim(rs_d!CODREG), 6)     'colegiatura
      End If
      rs_d.Close
      If Trim(cCargo) <> "" Then
         Set rs_d = New Recordset
         sql = "SELECT * FROM CPCargos WHERE codigo='" + cCargo + "'"
         rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs_d.EOF Then
            CICR = IIf(IsNull(rs_d!colegio), "", Trim(rs_d!colegio))
         End If
      End If
      rs_d.Close
      If CICR = "CMP" Then
         Set rs_d = New Recordset
         sql = "SELECT * FROM Medico WHERE Numero='" + Trim(cDni) + "'"
         rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs_d.EOF Then
            ccodreg = Right("0000" + Trim(rs_d!cmp), 6)
            ccodespe = Trim(rs_d!codespe)
         End If
         rs_d.Close
         If ccodreg <> "" Then
             Set rs_d = New Recordset
             sql = "Select * From Especialidad where Codespe='" + ccodespe + "'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                cdescEspec = Trim(rs_d!descripcion)
                cCodigo = IIf(IsNull(rs_d!Codigo), "99-093", Trim(rs_d!Codigo))
             End If
             rs_d.Close
         End If
      
         If cCodigo <> "" Then
             Set rs_d = New Recordset
             sql = "Select * From RefEspecialidad where Codigo='" + cCodigo + "'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                cdescEspec = Trim(rs_d!Esp_descripcion)
             End If
             rs_d.Close
         End If
      End If
      If CICR = "CEP" Then
         cdescEspec = "ENFERMERÍA PEDIÁTRICA"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and tipo='ENFERM'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      If CICR = "CTSP" Then
         cdescEspec = "TRABAJO SOCIAL SALUD Y COMUNIDAD"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='ASISOC'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      If CICR = "CQFP" Then
         cdescEspec = "FARMACIA CLÍNICA"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='QUIFAR'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      
      If CICR = "CNP" Then
         cdescEspec = "NUTRICIÓN CLÍNICA"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='NUTRIC'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      If CICR = "CTMP" Then
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='TECMED'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      If CICR = "CPSP" Then
         cdescEspec = "VIOLENCIA FAMILIAR, VIOLENCIA Y GÉNERO Y DERECHOS HUMANOS"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='PSICOL'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      If CICR = "COPOBS" Then
         cdescEspec = "MEDICINA MATERNO FETAL"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='OBSTET'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      
      
      '------If ccodreg <> "" Then   ' <=== SOLO colegiados
      
      If cCodUpr <> "" Then
          Set rs_d = New Recordset
          rs_d.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs_d.EOF Then
             wdscUpr = Trim(rs_d!servicioSusalud)
          End If
          rs_d.Close
       End If
       For i = diai To diaf      '<<====dias calendario mes analizar, max 31
            z = i - 1
            fechaB = Format(primer + z, "yyyymmdd")
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
                  '--If (rt!fecnac >= DateValue(Txpridia) And DateValue(Txpridia) <= dftope) Or (DateValue(TxUltDia) >= rt!fecnac And DateValue(TxUltDia) <= dftope) Then
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
                             
            cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
            '-HORARIO prog en dia.
            sql = "SELECT * FROM CpHor" + genaño + "  WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
            Set rs = New Recordset
            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If rs.EOF Then
               If Trim(cGrpGen) <> "" Then
                  rs.Close
                  sql = "SELECT * FROM CpHor" + genaño + "  WHERE grupo='" & cGrpGen & "' AND fecha ='" & fechaB & "'"
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
               cHoraI = Left(p_ing, 2) + ":" + Right(p_ing, 2) 'p0_ing
               cHoraF = Left(p_sal, 2) + ":" + Right(p_sal, 2)
               If Val(p_ing1) > 0 Then
                  cHoraI1 = Left(p_ing1, 2) + ":" + Right(p_ing1, 2) 'p0_ing1
                  cHoraF1 = Left(p_sal1, 2) + ":" + Right(p_sal1, 2)
                  nvez = 2
               End If
               nson = 11 + (z * 2)
               nson1 = nson + 1
               ccIng = rRS.Fields(nson).Name
               CCsal = rRS.Fields(nson + 1).Name
               For k = 1 To nvez
                  If k = 1 Then
                     Set rRS = New Recordset
                     sqRS = "SELECT * FROM #Tabla77   where nro_doc='" + cDni + "'"      'Directorio mensual -> export XLS
                     rRS.Open sqRS, Cnx, adOpenStatic, adLockOptimistic, adCmdText
                     If rRS.EOF() Then         '1er dia programado
                        rRS.AddNew
                        rRS!IPRESS = "00006212"
                        rRS!nombre = "HOSPITAL DE EMERGENCIAS PEDIATRICAS"
                        rRS!Red = "NO PERTENECE A NINGUNA RED"
                        rRS!tIP_DOC = "DNI"
                        rRS!Nro_Doc = cDni
                        rRS!colegio_prof = CICR
                        rRS!Codigo = ccodreg
                        rRS!apellidos = Trim(capelPat) + " " + Trim(capelMat)
                        rRS!nombres = Trim(cnombres)
                        rRS!Especialidad = Trim(cdescEspec)
                        rRS!Servicio = Trim(wdscUpr)
                     End If
                     rRS.Fields(nson) = cHoraI
                     rRS.Fields(nson1) = cHoraF
                  Else
                     Set rRS = New Recordset
                     sqRS = "SELECT * FROM #Tabla77   where nro_doc='" + cDni + "'"      'Directorio mensual -> export XLS
                     rRS.Open sqRS, Cnx, adOpenStatic, adLockOptimistic, adCmdText
                     If Not IsNull(rRS.Fields(nson)) Or Trim(rRS.Fields(nson)) <> "" Then
                         rRS.AddNew
                         rRS!IPRESS = "00006212"
                         rRS!nombre = "HOSPITAL DE EMERGENCIAS PEDIATRICAS"
                         rRS!Red = "DIRIS LIMA-CENTRO"
                         rRS!tIP_DOC = "DNI"
                         rRS!Nro_Doc = cDni
                         rRS!colegio_prof = CICR
                         rRS!Codigo = ccodreg
                         rRS!apellidos = Trim(capelPat) + " " + Trim(capelMat)
                         rRS!nombres = Trim(cnombres)
                         rRS!Especialidad = Trim(cdescEspec)
                         rRS!Servicio = Trim(wdscUpr)
                     End If
                     rRS.Fields(nson) = cHoraI1
                     rRS.Fields(nson1) = cHoraF1
                  End If
                  rRS.Update
                  rRS.Close
               Next k
               
            End If
            
       Next i             '<<==== sgt Dia calendario mes analizar, max 31
       
       '-----End If  '<===
       
       If Len(cCadenaDias) > 0 Then
          wdato = wdato + 1
          CmdCarga.Caption = Trim(Str(wdato)) + "/" + Trim(Str(nreg))
       End If
       r_d.MoveNext    'Sgt Servidor
   Loop
   r_d.Close
   Cnx.CommitTrans
   CmdCarga.Caption = ""
   Beep 500, 350
   Progreso1.Visible = False
   If wdato > 0 Then
      Lblrec.Caption = "Son " & Trim(Str(wdato)) + " servidores "
   End If

End Sub

'Private Sub CmdPrint_Click()
'   If wdato = 0 Then Exit Sub
'   wForma = "QUERYPROGRAMA"
'   wTituRpt = ""
'   InfoCpRS.Left = 0: InfoCpRS.Top = 0
'   InfoCpRS.Show
'   Exit Sub
'End Sub

Private Sub CmdExportaXLS_Click()

On Error GoTo final
    '--   RecPerTarjeta
    If wdato = 0 Then Exit Sub
       If wTpoQuery = "7.7" Then
    
          If ChkXLS.Value = 0 Then
             Call ExportTableToXLS_click
          Else
             wNameXLS = "C:\RPT\DirPrograma.xls"
             Call Exportar_ADO_Excel
          End If
          Beep 500, 350
          MsgBox "Se genero C:\RPT\DirPrograma.xls", vbInformation
          Exit Sub
       ElseIf wTpoQuery = "7.8" Then
    
          If ChkXLS.Value = 0 Then
             Call ExportTableToXLS_click
          Else
             wNameXLS = "C:\RPT\DirProgAll.xls"
             Call Exportar_ADO_Excel
          End If
          Beep 500, 350
          MsgBox "Se genero C:\RPT\DirProgAll.xls", vbInformation
          Exit Sub
       End If
          
final:
   MsgBox Err.Description & " Error: " & Err.Number & vbCr
    
End Sub
Private Sub cmdsalir_Click()
  Set rD = Nothing
  Set rs = Nothing
  Set rRS = Nothing
  Set DD = Nothing
  For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        '- Forms Abiertos
        If UCase(Forms(cont).Name) <> "CPMENU" And UCase(Forms(cont).Name) <> "RECMARCASQUERY" Then
           Unload Forms(cont)
        End If
      End If
  Next
  Me.Hide
  Unload Me
End Sub

Private Sub CmdCarga1_Click()
  If Val(Txmes) = 0 And ChkAnual.Value = 0 Then
     MsgBox "favor indicar periodo de consulta..", vbCritical
     Exit Sub
  End If
  If Trim(Txtipopla) = "" Then
     MsgBox "No indica tipo de planilla..", vbCritical
     Exit Sub
  End If
  
  'If Trim(Txccosto) = "" And Trim(Txcodupr) = "" Then
  '   MsgBox "Indicar parametros de consulta..", vbCritical, "[ ERROR ]"
  '   Exit Sub
  'End If

  Lblrec.Caption = ""
  wdato = 0
  Criterio = " FLAG='" + Txtipopla + "' AND ano ='" & genaño & "' "
  If ChkAnual.Value = 0 Then
     Criterio = Criterio & "  AND mes='" & Trim(Txmes) & "' "
     Txpridia = primer
     TxUltDia = ultimo

  Else
     If Year(Date) = Val(genaño) Then
        'año actual
        mFin = Right("0" + Trim(Str(Month(Date))), 2)
     Else
        'año anterior
        mFin = "12"
     End If
     Criterio = Criterio & "  AND mes<='" & mFin & "' "
  End If
  
  f_iltro = ""
  If Txccosto = "" And TxCodUpr = "" Then
     Criterio = Criterio + cCadenaCC
  Else
     If Txccosto <> "" Then
       If Len(Criterio) > 0 Then
            Criterio = Criterio + " and  (ccosto='" & Txccosto & "')"
            f_iltro = f_iltro & ", C.Costo "
       Else
            Criterio = " (ccosto='" & Txccosto & "')"
            f_iltro = "C.Costo "
       End If
     End If
     If TxCodUpr <> "" Then
       If Len(Criterio) > 0 Then
            Criterio = Criterio + " and  (codupr='" & TxCodUpr & "')"
            f_iltro = f_iltro & ", UPR "
       Else
            Criterio = " (codupr='" & TxCodUpr & "')"
            f_iltro = "UPR "
       End If
     End If
  End If
    
  If Criterio = "" Then
     MsgBox "Favor defina criterio ", vbExclamation
     Exit Sub
  End If
   
 '-----  MsgBox Criterio
 ' ndiasson = DateDiff("d", primer, ultimo) + 1
  
  wTpoQuery = "7.8"
  If ChkAnual.Value = 0 Then
     Call Carga22
  Else
     Call Carga22Anual
  End If
  
End Sub

Private Sub Carga22()
   Lblrec.Caption = ""
   wdato = 0
   
   '-Directorio mensual
   Cnx.BeginTrans
   Cnx.Execute "DELETE  FROM CpDirProgALL"
   On Error GoTo Err71
   Cnx.Execute "DELETE  FROM #Tabla78"
   GoTo sigue71
Err71:
   Cnx.Execute ("SELECT * Into #Tabla78  FROM CpDirProgALL WHERE 1 = 2")
sigue71:
   Cnx.CommitTrans

   Set rRS = New Recordset
   sqRS = "SELECT * FROM #Tabla78"    'Directorio mensual -> export XLS
   rRS.Open sqRS, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   rRS.Close
   
   Set r_d = New ADODB.Recordset
   sqG = "SELECT * FROM CpDirectorio WHERE " & Criterio
   r_d.Open sqG, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If r_d.EOF Then
      Exit Sub
   End If
   nreg = r_d.RecordCount
  
   Progreso1.Value = 0
   Progreso1.Max = nreg
   Progreso1.Visible = True
   r_d.MoveFirst
   CmdCarga1.Caption = Trim(Str(nreg))
   Cnx.BeginTrans
   Dim OpenForms
   v_an = 0
   '---   Ncuenta0 = 0: ncuenta1 = 0
   Do While Not r_d.EOF         '=====> Servidores
   
      '-dar control a windows y permitir multihilo
      v_an = v_an + 1
      If v_an Mod 10 = 0 Then     'if loop has repeated
         OpenForms = DoEvents 'Yield to operating system
      End If
   
      With Progreso1
        If .Value < .Max Then
           .Value = .Value + 1
        End If
      End With
      
      cCodProg = "": cprogUPR = "": cDscCondicion = ""
      capelPat = "": capelMat = "": cnombres = "": csexo = "":  wdsccosto = "": wdscUpr = ""
      cdescEspec = "SIN ESPECIALIDAD": ccodreg = "": CICR = "": cCadenaDias = ""
      
      cCodProg = Trim(r_d!usuario)
      cTarjeta = Trim(r_d!tarjeta)     '- servidor
      ccCosto = IIf(IsNull(r_d!CCosto), "", Trim(r_d!CCosto))
      cCodUpr = IIf(IsNull(r_d!CodUpr), "", Trim(r_d!CodUpr))
      cCargo = IIf(IsNull(r_d!Cargo), "", Trim(r_d!Cargo))
      cCondicion = IIf(IsNull(r_d!PLH), "", Trim(r_d!PLH))
      
      Set rs_d = New Recordset
      sql = "SELECT * FROM CPUsuarios WHERE usuario='" & cCodProg & "'"
      rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs_d.EOF Then
         rs_d.Close
         Set rs_d = New Recordset
         sql = "SELECT * FROM CPGeneral WHERE documento='" & cCodProg & "'"
         rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs_d.EOF Then
            cprogUPR = Left(Trim(rs_d!apelpat) + " " + Trim(rs_d!apelpat) + ", " + Trim(rs_d!nombres), 45)
         End If
      Else
         cprogUPR = Left(rs_d!nombre, 45)
      End If
      rs_d.Close
      If cprogUPR = "" Then
         cprogUPR = r_d!usuario
      End If
      
      Set rs_d = New Recordset
      sql = "SELECT * FROM CpGeneral WHERE tarjeta='" + cTarjeta + "'"
      rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs_d.EOF Then
         capelPat = IIf(IsNull(rs_d!apelpat), "", Trim(rs_d!apelpat))
         capelMat = IIf(IsNull(rs_d!apelmat), "", Trim(rs_d!apelmat))
         cnombres = IIf(IsNull(rs_d!nombres), "", Trim(rs_d!nombres))
         cDni = Trim(rs_d!documento)
         csexo = IIf(IsNull(rs_d!sexo), "", Trim(rs_d!sexo))
            ccodreg = Right("0000" + Trim(rs_d!CODREG), 6)     'colegiatura
      End If
      rs_d.Close
          '---    Ncuenta0 = Ncuenta0 + 1 ' servidores  asistenciales   programados
      If Trim(cCargo) <> "" Then
         Set rs_d = New Recordset
         sql = "SELECT * FROM CPCargos WHERE codigo='" + cCargo + "'"
         rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs_d.EOF Then
            CICR = IIf(IsNull(rs_d!colegio), "", Trim(rs_d!colegio))
         End If
      End If
      rs_d.Close
      If CICR = "CMP" Then
         Set rs_d = New Recordset
         sql = "SELECT * FROM Medico WHERE Numero='" + Trim(cDni) + "'"
         rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs_d.EOF Then
            ccodreg = Right("0000" + Trim(rs_d!cmp), 6)
            ccodespe = Trim(rs_d!codespe)
         End If
         rs_d.Close
         If ccodreg <> "" Then
             Set rs_d = New Recordset
             sql = "Select * From Especialidad where Codespe='" + ccodespe + "'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                cdescEspec = Trim(rs_d!descripcion)
                cCodigo = IIf(IsNull(rs_d!Codigo), "99-093", Trim(rs_d!Codigo))
             End If
             rs_d.Close
         End If
         
           '  If wdato < 4 Or ccodreg = "0523661" Then
           '     MsgBox "1-" + codreg + capelPat + " " + cnombres + "-" + cdescEspec
           '  End If
             
         If cCodigo <> "" And cdescEspec = "" Then
             Set rs_d = New Recordset
             sql = "Select * From RefEspecialidad where Codigo='" + cCodigo + "'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                cdescEspec = Trim(rs_d!Esp_descripcion)
             End If
             rs_d.Close
         End If
      End If
      If CICR = "CEP" Then
         cdescEspec = "ENFERMERÍA PEDIÁTRICA"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and tipo='ENFERM'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      If CICR = "CTSP" Then
         cdescEspec = "TRABAJO SOCIAL SALUD Y COMUNIDAD"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='ASISOC'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      If CICR = "CQFP" Then
         cdescEspec = "FARMACIA CLÍNICA"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='QUIFAR'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      
      If CICR = "CNP" Then
         cdescEspec = "NUTRICIÓN CLÍNICA"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='NUTRIC'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      If CICR = "CTMP" Then
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='TECMED'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      If CICR = "CPSP" Then
         cdescEspec = "VIOLENCIA FAMILIAR, VIOLENCIA Y GÉNERO Y DERECHOS HUMANOS"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='PSICOL'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      If CICR = "COPOBS" Then
         cdescEspec = "MEDICINA MATERNO FETAL"
         'Set rs_d = New Recordset
         'sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='OBSTET'"
         'rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
         'If Not rs_d.EOF Then
         '   ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
         'End If
         'rs_d.Close
      End If
      
      If cCodUpr <> "" Then
          Set rs_d = New Recordset
          rs_d.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs_d.EOF Then
             wdscUpr = Trim(rs_d!servicioSusalud)
          End If
          rs_d.Close
       End If
       Txpridia = primer
       TxUltDia = ultimo
       diai = Day(Txpridia)    'Dias de inicio y cierre del periodo consultado
       diaf = Day(TxUltDia)    'Day(ultimo)
       For i = diai To diaf      '<<====dias calendario mes analizar, max 31
            z = i - 1
            fechaB = Format(primer + z, "yyyymmdd")
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
                  '--If (rt!fecnac >= DateValue(Txpridia) And DateValue(Txpridia) <= dftope) Or (DateValue(TxUltDia) >= rt!fecnac And DateValue(TxUltDia) <= dftope) Then
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
                             
            cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
            If InStr("MTDNHCP", Left(cTurno, 1)) > 0 And cTurno <> "" Then
               If cTurno = "MT" Then
                  cDscTurno = "Ma/Ta"
               ElseIf Left(cTurno, 1) = "D" Then
                      cDscTurno = "Dia"
               ElseIf Left(cTurno, 1) = "M" Then
                      cDscTurno = "Mañana"
               ElseIf Left(cTurno, 1) = "T" Then
                      cDscTurno = "Tarde"
               ElseIf Left(cTurno, 1) = "N" Then
                      cDscTurno = "Noche"
               ElseIf Left(cTurno, 1) = "H" Then
                      cDscTurno = "Ma/No"
               ElseIf Left(cTurno, 1) = "C" Then
                      cDscTurno = "Ta/No"
               ElseIf Left(cTurno, 1) = "P" Then
                      cDscTurno = "Di/No"
               End If
                      
               '-HORARIO prog en dia.
               sql = "SELECT * FROM CpHor" + genaño + "  WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
               Set rs = New Recordset
               rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If rs.EOF Then
                  If Trim(cGrpGen) <> "" Then
                     rs.Close
                     sql = "SELECT * FROM CpHor" + genaño + "  WHERE grupo='" & cGrpGen & "' AND fecha ='" & fechaB & "'"
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
            End If
            
            If Val(p_horas) > 0 Then
               cCadenaDias = cCadenaDias + Right("0" + Trim(Str(i)), 2)
               cHoraI1 = Left(p_ing, 2) + ":" + Right(p_ing, 2) 'p0_ing
               cHoraF1 = Left(p_sal, 2) + ":" + Right(p_sal, 2)
               If Val(p_ing1) > 0 Then
                  cHoraI2 = Left(p_ing1, 2) + ":" + Right(p_ing1, 2) 'p0_ing1
                  cHoraF2 = Left(p_sal1, 2) + ":" + Right(p_sal1, 2)
               End If
               Set rRS = New Recordset
               sqRS = "SELECT * FROM #Tabla78   where nrodoc='" + cDni + "' and fecha='" + fechaB + "'"  'Directorio mensual -> export XLS
               rRS.Open sqRS, Cnx, adOpenStatic, adLockOptimistic, adCmdText
               If rRS.EOF() Then         'dia programado
                  rRS.AddNew
                  rRS!fecha = primer + z
                  rRS!Servicio = Trim(wdscUpr)
                  rRS!Especialidad = Trim(cdescEspec)
                  rRS!Condicion = cCondicion
                  rRS!tipdoc = "DNI"
                  rRS!NroDoc = cDni
                  rRS!trabajador = Trim(capelPat) + " " + Trim(capelMat) + ", " + Trim(cnombres)
               End If
               rRS!DscTurno = cDscTurno
               rRS!HrInicio1 = cHoraI1
               rRS!HrFin1 = cHoraF1
               rRS!HrInicio2 = cHoraI2
               rRS!HrFin2 = cHoraF2
               rRS!Horas = nHrProg
               rRS!Programador = cprogUPR
               rRS.Update
               rRS.Close
            End If
            
       Next i             '<<==== sgt Dia calendario mes analizar, max 31
       
       If Len(cCadenaDias) > 0 Then
       
         '--- ncuenta1 = ncuenta1 + 1 ' servidores  asistenciales programados COLEGIADOS excluye periodos LVD
         ' If wdato < 4 Or ccodreg = "0523661" Then
         '    MsgBox "2-" + ccodreg + capelPat + " " + cnombres + "-" + cdescEspec
         ' End If
        
          wdato = wdato + 1
          CmdCarga1.Caption = Trim(Str(wdato)) + "/" + Trim(Str(nreg))
          
       Else
          '- no genero cadena
          'MsgBox Trim(capelPat) + " " + (capelMat) + " " + Trim(cnombres) + "-" + cDni
          
       End If
       r_d.MoveNext    'Sgt Servidor
   Loop
   r_d.Close
   Cnx.CommitTrans
    '---If Ncuenta0 <> ncuenta1 Then
    '---   MsgBox "hay servidores no incluidos " + Str(Ncuenta0) + "<>" + Str(ncuenta1)
    '---End If
   
   CmdCarga1.Caption = ""
   Beep 500, 350
   Progreso1.Visible = False
   If wdato > 0 Then
      Lblrec.Caption = "Son " & Trim(Str(wdato)) + " servidores "
   End If

End Sub

Private Sub Carga22Anual()
   Lblrec.Caption = ""
   wdato = 0
   '-Directorio mensual
   Cnx.BeginTrans
   Cnx.Execute "DELETE  FROM CpDirProgALL"
   On Error GoTo Err78
   Cnx.Execute "DELETE  FROM #Tabla78"
   GoTo sigue78
Err78:
   Cnx.Execute ("SELECT * Into #Tabla78  FROM CpDirProgALL WHERE 1 = 2")
sigue78:
   Cnx.CommitTrans

   Set rRS = New Recordset
   sqRS = "SELECT * FROM #Tabla78"    'Directorio mensual -> export XLS
   rRS.Open sqRS, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   rRS.Close
   
   Set r_d = New ADODB.Recordset
   sqG = "SELECT * FROM CpDirectorio WHERE " & Criterio & " ORDER BY ano,mes"
   r_d.Open sqG, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If r_d.EOF Then
      Exit Sub
   End If
   nreg = r_d.RecordCount
  
   Progreso1.Value = 0
   Progreso1.Max = nreg
   Progreso1.Visible = True
   
   CmdCarga1.Caption = Trim(Str(nreg))
   Cnx.BeginTrans
   Dim OpenForms
   v_an = 0
   r_d.MoveFirst
   cmes = "99"
   Do While Not r_d.EOF         '=====> Servidores
      If cmes <> r_d!mes Then
          cmes = r_d!mes
          unafecha = "01/" + cmes + "/" + Trim(Str(genaño))
          Call ObtenerPrimerUltimoDia(unafecha)
          Txpridia = primer
          TxUltDia = ultimo
          diai = Day(Txpridia)    'Dias de inicio y cierre del periodo consultado
          diaf = Day(TxUltDia)    'Day(ultimo)
      End If
         '-dar control a windows y permitir multihilo
          v_an = v_an + 1
          If v_an Mod 10 = 0 Then     'if loop has repeated
             OpenForms = DoEvents 'Yield to operating system
          End If
   
          With Progreso1
            If .Value < .Max Then
               .Value = .Value + 1
            End If
          End With
          cCadenaDias = ""
          cCodProg = "": cprogUPR = "": cDscCondicion = ""
          capelPat = "": capelMat = "": cnombres = "": csexo = "":  wdsccosto = "": wdscUpr = ""
          cdescEspec = "SIN ESPECIALIDAD": ccodreg = "": CICR = ""
          
          cTarjeta = Trim(r_d!tarjeta)     '- servidor
          ccCosto = IIf(IsNull(r_d!CCosto), "", Trim(r_d!CCosto))
          cCodUpr = IIf(IsNull(r_d!CodUpr), "", Trim(r_d!CodUpr))
          cCargo = IIf(IsNull(r_d!Cargo), "", Trim(r_d!Cargo))
          cCondicion = IIf(IsNull(r_d!PLH), "", Trim(r_d!PLH))
          cCodProg = Trim(r_d!usuario)
          Set rs_d = New Recordset
          sql = "SELECT * FROM CPUsuarios WHERE usuario='" & cCodProg & "'"
          rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
          If rs_d.EOF Then
             rs_d.Close
             Set rs_d = New Recordset
             sql = "SELECT * FROM CPGeneral WHERE documento='" & cCodProg & "'"
             rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                cprogUPR = Left(Trim(rs_d!apelpat) + " " + Trim(rs_d!apelpat) + ", " + Trim(rs_d!nombres), 45)
             End If
          Else
             cprogUPR = Left(rs_d!nombre, 45)
          End If
          rs_d.Close
          If cprogUPR = "" Then
             cprogUPR = r_d!usuario
          End If
      
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
          If Trim(cCargo) <> "" Then
             Set rs_d = New Recordset
             sql = "SELECT * FROM CPCargos WHERE codigo='" + cCargo + "'"
             rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                CICR = IIf(IsNull(rs_d!colegio), "", Trim(rs_d!colegio))
             End If
          End If
          rs_d.Close
          If CICR = "CMP" Then
             Set rs_d = New Recordset
             sql = "SELECT * FROM Medico WHERE Numero='" + Trim(cDni) + "'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                ccodreg = Right("0000" + Trim(rs_d!cmp), 6)
                ccodespe = Trim(rs_d!codespe)
             End If
             rs_d.Close
             If ccodreg <> "" Then
                Set rs_d = New Recordset
                sql = "Select * From Especialidad where Codespe='" + ccodespe + "'"
                rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
                If Not rs_d.EOF Then
                   cdescEspec = Trim(rs_d!descripcion)
                   cCodigo = IIf(IsNull(rs_d!Codigo), "99-093", Trim(rs_d!Codigo))
                End If
                rs_d.Close
             End If
         
             '  If wdato < 4 Or ccodreg = "0523661" Then
             '     MsgBox "1-" + codreg + capelPat + " " + cnombres + "-" + cdescEspec
             '  End If
             
             If cCodigo <> "" And cdescEspec = "" Then
                Set rs_d = New Recordset
                sql = "Select * From RefEspecialidad where Codigo='" + cCodigo + "'"
                rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
                If Not rs_d.EOF Then
                   cdescEspec = Trim(rs_d!Esp_descripcion)
                End If
                rs_d.Close
             End If
          End If
          If CICR = "CEP" Then
             cdescEspec = "ENFERMERÍA PEDIÁTRICA"
             Set rs_d = New Recordset
             sql = "Select * From ProfSalud where numero='" + cDni + "' and tipo='ENFERM'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
             End If
             rs_d.Close
          End If
          If CICR = "CTSP" Then
             cdescEspec = "TRABAJO SOCIAL SALUD Y COMUNIDAD"
             Set rs_d = New Recordset
             sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='ASISOC'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
             End If
             rs_d.Close
          End If
          If CICR = "CQFP" Then
             cdescEspec = "FARMACIA CLÍNICA"
             Set rs_d = New Recordset
             sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='QUIFAR'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
             End If
             rs_d.Close
          End If
      
          If CICR = "CNP" Then
             cdescEspec = "NUTRICIÓN CLÍNICA"
             Set rs_d = New Recordset
             sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='NUTRIC'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
             End If
             rs_d.Close
          End If
          If CICR = "CTMP" Then
             Set rs_d = New Recordset
             sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='TECMED'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
             End If
             rs_d.Close
          End If
          If CICR = "CPSP" Then
             cdescEspec = "VIOLENCIA FAMILIAR, VIOLENCIA Y GÉNERO Y DERECHOS HUMANOS"
             Set rs_d = New Recordset
             sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='PSICOL'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
             End If
             rs_d.Close
          End If
          If CICR = "COPOBS" Then
             cdescEspec = "MEDICINA MATERNO FETAL"
             Set rs_d = New Recordset
             sql = "Select * From ProfSalud where numero='" + cDni + "' and TIPO='OBSTET'"
             rs_d.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
             If Not rs_d.EOF Then
                ccodreg = Left("0000" + Trim(rs_d!Codigo), 6)
             End If
             rs_d.Close
          End If
      
          If cCodUpr <> "" Then
             Set rs_d = New Recordset
             rs_d.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
             If Not rs_d.EOF Then
                wdscUpr = Trim(rs_d!servicioSusalud)
             End If
             rs_d.Close
          End If
          For i = diai To diaf      '<<====dias calendario mes analizar, max 31
            z = i - 1
            fechaB = Format(primer + z, "yyyymmdd")
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
                  '--If (rt!fecnac >= DateValue(Txpridia) And DateValue(Txpridia) <= dftope) Or (DateValue(TxUltDia) >= rt!fecnac And DateValue(TxUltDia) <= dftope) Then
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
                             
            cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
            If InStr("MTDNHCP", Left(cTurno, 1)) > 0 And cTurno <> "" Then
               If cTurno = "MT" Then
                  cDscTurno = "Ma/Ta"
               ElseIf Left(cTurno, 1) = "D" Then
                      cDscTurno = "Dia"
               ElseIf Left(cTurno, 1) = "M" Then
                      cDscTurno = "Mañana"
               ElseIf Left(cTurno, 1) = "T" Then
                      cDscTurno = "Tarde"
               ElseIf Left(cTurno, 1) = "N" Then
                      cDscTurno = "Noche"
               ElseIf Left(cTurno, 1) = "H" Then
                      cDscTurno = "Ma/No"
               ElseIf Left(cTurno, 1) = "C" Then
                      cDscTurno = "Ta/No"
               ElseIf Left(cTurno, 1) = "P" Then
                      cDscTurno = "Di/No"
               End If
                      
               '-HORARIO prog en dia.
               sql = "SELECT * FROM CpHor" + genaño + "  WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
               Set rs = New Recordset
               rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If rs.EOF Then
                  If Trim(cGrpGen) <> "" Then
                     rs.Close
                     sql = "SELECT * FROM CpHor" + genaño + "  WHERE grupo='" & cGrpGen & "' AND fecha ='" & fechaB & "'"
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
            End If
            
            If Val(p_horas) > 0 Then
               cCadenaDias = cCadenaDias + Right("0" + Trim(Str(i)), 2)
               cHoraI1 = Left(p_ing, 2) + ":" + Right(p_ing, 2) 'p0_ing
               cHoraF1 = Left(p_sal, 2) + ":" + Right(p_sal, 2)
               If Val(p_ing1) > 0 Then
                  cHoraI2 = Left(p_ing1, 2) + ":" + Right(p_ing1, 2) 'p0_ing1
                  cHoraF2 = Left(p_sal1, 2) + ":" + Right(p_sal1, 2)
               End If
               Set rRS = New Recordset
               sqRS = "SELECT * FROM #Tabla78   where nrodoc='" + cDni + "' and fecha='" + fechaB + "'"  'Directorio mensual -> export XLS
               rRS.Open sqRS, Cnx, adOpenStatic, adLockOptimistic, adCmdText
               rRS.AddNew
               rRS!fecha = primer + z
               rRS!Servicio = Trim(wdscUpr)
               rRS!Especialidad = Trim(cdescEspec)
               rRS!Condicion = cCondicion
               rRS!tipdoc = "DNI"
               rRS!NroDoc = cDni
               rRS!trabajador = Trim(capelPat) + " " + Trim(capelMat) + ", " + Trim(cnombres)
               rRS!DscTurno = cDscTurno
               rRS!HrInicio1 = cHoraI1
               rRS!HrFin1 = cHoraF1
               rRS!HrInicio2 = cHoraI2
               rRS!HrFin2 = cHoraF2
               rRS!Horas = nHrProg
               rRS!Programador = cprogUPR
               rRS.Update
               rRS.Close
            End If
            
          Next i             '<<==== sgt Dia calendario mes analizar, max 31
       
          If Len(cCadenaDias) > 0 Then
       
            ' If wdato < 4 Or ccodreg = "0523661" Then
            '    MsgBox "2-" + ccodreg + capelPat + " " + cnombres + "-" + cdescEspec
            ' End If

             wdato = wdato + 1
             CmdCarga1.Caption = cmes + " - " + Trim(Str(wdato)) + "/" + Trim(Str(nreg))
          End If
          r_d.MoveNext    'Sgt Servidor
          
   Loop
   r_d.Close
   Cnx.CommitTrans
   CmdCarga1.Caption = ""
   Beep 500, 350
   Progreso1.Visible = False
   If wdato > 0 Then
      Lblrec.Caption = "Son " & Trim(Str(wdato)) + " servidores "
   End If

End Sub


