VERSION 5.00
Begin VB.Form CPFaltas 
   Caption         =   "Analisis AUSENCIAS / ASISTENTES del Periodo"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "CPFaltas.frx":0000
   ScaleHeight     =   4020
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkXLS 
      Alignment       =   1  'Right Justify
      Caption         =   "XLS instalado"
      Height          =   195
      Left            =   2520
      TabIndex        =   40
      Top             =   3120
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resumen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   480
      TabIndex        =   22
      Top             =   1680
      Width           =   6255
      Begin VB.Label LblPorc1 
         Caption         =   "Porc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5520
         TabIndex        =   39
         Top             =   960
         Width           =   495
      End
      Begin VB.Label LblPORC0 
         BackStyle       =   0  'Transparent
         Caption         =   "Porc0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   5520
         TabIndex        =   38
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Lbla1 
         BackStyle       =   0  'Transparent
         Caption         =   "a1"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4680
         TabIndex        =   37
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Lbla0 
         Caption         =   "a0"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4680
         TabIndex        =   36
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Lblp1 
         BackStyle       =   0  'Transparent
         Caption         =   "p1"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         Top             =   960
         Width           =   495
      End
      Begin VB.Label LblP0 
         BackStyle       =   0  'Transparent
         Caption         =   "p0"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3720
         TabIndex        =   34
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Lbld1cas 
         Caption         =   "d1cas"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2640
         TabIndex        =   33
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Lbld0Cas 
         Caption         =   "d0cas"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2640
         TabIndex        =   32
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Lbld1hep 
         Caption         =   "d1hep"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1440
         TabIndex        =   31
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Lbld0hep 
         BackStyle       =   0  'Transparent
         Caption         =   "d0hep"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Administrativos"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Asistenciales"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "% Ausencia"
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   5400
         TabIndex        =   27
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ausentes"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   4560
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Presentes"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "C.A.S."
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nombrados"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "PRESENTES"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "AUSENTES"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "No 3ros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox TxHrasMP 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5400
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxHrasDM 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox TxHrasFalt 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5040
      TabIndex        =   19
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxPorcPerd 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5880
      TabIndex        =   18
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxTotHrasProg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4560
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton CmdExportaXLS 
      Height          =   615
      Left            =   2760
      Picture         =   "CPFaltas.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Exporta a Xls"
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton CmdSalir 
      Height          =   735
      Left            =   7080
      Picture         =   "CPFaltas.frx":80EA8
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton CmdClear 
      Height          =   615
      Left            =   4800
      Picture         =   "CPFaltas.frx":80FF2
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Otra Consulta"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   1095
      Left            =   6840
      Picture         =   "CPFaltas.frx":8110C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ubicar Registros"
      Top             =   720
      Width           =   1215
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
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   975
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
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   975
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
      Left            =   1320
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   975
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
      Left            =   2760
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox TxDsccargo 
      Enabled         =   0   'False
      Height          =   195
      Left            =   4680
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Txtipopla 
      Enabled         =   0   'False
      Height          =   225
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboTipoPla 
      Height          =   315
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   2055
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
      Left            =   1200
      Picture         =   "CPFaltas.frx":812E7
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir"
      Top             =   3240
      Width           =   765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas         Del :"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Planilla"
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "CPFaltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim r_t As New ADODB.Recordset
Dim Sw, f_cierra As Boolean
Dim nson, nsonC, nSonLacta, nLacta, nposicion, nlen, dfi, dff, z As Integer
Dim nmintarde, nTotHprog As Double

Dim nTNomb, ntAsiNomb, ntAdmNomb, nTCas, ntAsiCas, ntAdmCas As Integer
Dim ndNomb, ndAsiNomb, ndAdmNomb, ndCas, ndAsiCas, ndAdmCas As Integer
Dim p0, a0, p1, a1 As Double

Dim c_riterio, cDni, csexo, cAplica, sql, t_urnoD, c_turno, c_Costo, cdscccosto, cCodUpr, cDscUpr, cvalor, c_opc, cDActividad As String
Dim cPresente, cDscCargo, cCargo, cDscCategoria, cCategoria As String
Dim unafecha As Date

Private Sub Form_Load()
    Frame1.Visible = False
    CPFaltas.Caption = "Faltas / Presentes - " & genaño
    TxFecIni = Date - 7: TxFecFin = Date - 1
    Txtipopla = ""
    Option1(0).Value = False: Option1(1).Value = False
    ChkXLS.Value = 1
    '-LLena el Combo TipoPLA
    CboTipoPla.Clear
    CboTipoPla.AddItem "ASISTENCIAL"
    CboTipoPla.AddItem "ADMINISTRATIVO"
    CboTipoPla.ListIndex = 0
    CboTipoPla.Text = "": cPresente = ""
    Check1.Value = 1
End Sub

Private Sub txfecini_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If IsDate(TxFecIni) Then
        Exit Sub
     Else
        MsgBox "verifique fecha ..", vbExclamation
     End If
  End If
End Sub
Private Sub txfecFin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If IsDate(TxFecFin) Then
        If Month(DateValue(TxFecIni)) > Month(DateValue(TxFecFin)) Then
            MsgBox "verifique  fechas ", vbExclamation
            Exit Sub
        End If
        CboTipoPla.SetFocus
     Else
        MsgBox "verifique  fechas...", vbCritical
        TxFecFin = ""
     End If
  End If
End Sub
Private Sub TxFecFin_LostFocus()
   Call txfecFin_KeyPress(13)
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
        cPresente = "N"
      Case 1
        cPresente = "S"
   End Select
   
  '-- MsgBox cPresente
End Sub

Private Sub Cbotipopla_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboTipoPla.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboTipoPla_KeyPress(KeyAscii As Integer)
  If IsEmpty(CboTipoPla.Text) Then
  Else
     j = CboTipoPla.ListIndex
     Txtipopla = Left(CboTipoPla.List(j), 3)
  End If
End Sub
Private Sub Cbotipopla_LostFocus()
   Call CboTipoPla_KeyPress(13)
End Sub

Private Sub cmdcarga_Click()
   If cPresente = "" Then
      MsgBox "Favor indique Faltas / Presentes..", vbExclamation
      Exit Sub
   End If

   If DateValue(TxFecFin) > Date Then
      MsgBox "verifique fecha..", vbExclamation
      Exit Sub
   End If
   
   f_cierra = True
   If DateValue(TxFecFin) = Date Then
      f_cierra = False
   End If
   
   mesi = Month(DateValue(TxFecIni))
   mesf = Month(DateValue(TxFecFin))
      
   '''On Error Resume Next
   Cnx.Execute "DELETE  FROM CpTmpAnalis"
   Cnx.BeginTrans
   On Error GoTo CREA
   Cnx.Execute ("DELETE  FROM #tabla9 ")
   GoTo sigue
CREA:
   Cnx.Execute ("SELECT * Into #tabla9 FROM CpTmpAnalis WHERE 1 = 2")
 
sigue:
   Cnx.CommitTrans
       
   sql = "SELECT * FROM #tabla9 "
   Set rt = New Recordset
   rt.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   
   dfi = Day(TxFecIni)
   dff = Day(TxFecFin)
   cmes = Right("00" + Trim(Str(Month(TxFecIni))), 2)
   unafecha = TxFecIni
   Call ObtenerPrimerUltimoDia(unafecha)
   TxTotHrasProg = "": TxHrasFalt = "": TxHrasMP = "": TxHrasDM = ""
   nTotHprog = 0: nTothTrab = 0: nTothFalt = 0:  ntotMP = 0: ntotDM = 0: nposicion = 0: nson = 0
   
   nTNomb = 0: ntAsiNomb = 0: ntAdmNomb = 0: ntAdmCas = 0: ntAsiCas = 0: nTCas = 0
   ndNomb = 0: ndAsiNomb = 0: ndAdmNomb = 0: ndAdmCas = 0: ndAsiCas = 0: ndCas = 0
   
   For k = mesi To mesf  '<========  (0)  PERIODO
       cmesk = Right("00" + Trim(Str(k)), 2)
       If cmesk = cmes Then
          unafecha = TxFecIni
       Else
          unafecha = "01/" + cmesk + "/" + genaño
       End If
       Call ObtenerPrimerUltimoDia(unafecha)
       dfi = Day(primer)
       dff = Day(ultimo)
       '-Corrige dia de inicio y fin del periodo
       If k = mesi Then
          dfi = Day(TxFecIni)
       End If
       If k = mesf Then
          dff = Day(TxFecFin)
       End If
       
       '''MsgBox Str(mesi) + "-" + Str(mesf) + "->" + cmesk + "del:" + Str(dfi) + " al:" + Str(dff)
   
   '-carga  "TURNO" Directorio
   TxDscCargo = ""
   c_riterio = ""
   If Check1.Value = 1 Then
      '-Excluye 3ROS
      c_riterio = " AND (PLH is NULL OR PLH<>'3RO') "
   End If
   If Trim(Txtipopla) = "" Then
   Else
      If Txtipopla = "ASI" Then
         c_riterio = c_riterio + " AND Flag='0' "
      ElseIf Txtipopla = "ADM" Then
         c_riterio = c_riterio + " AND Flag='1' "
      End If
   End If
      
   Dim OpenForms
   j = 0
      
   sql = "SELECT * FROM CpDirectorio  WHERE ano ='" & genaño & "' AND mes='" & cmesk & "' " & c_riterio + " ORDER BY tarjeta "
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   wson = rs.RecordCount
   rs.MoveFirst
   wvan = 0
   Do While Not rs.EOF
           '-Dar control a Windows y retomar hilo
           j = j + 1
           If j Mod 20 = 0 Then        'if loop has repeated
              OpenForms = DoEvents     'Yield to operating system
           End If
      wvan = wvan + 1
      '--cmdcarga.Caption = rs!tarjeta
      cmdcarga.Caption = Trim(Str(k)) + ":" + Trim(Str(wvan)) + "/" + Trim(Str(wson))
      cTarjeta = Trim(rs!tarjeta)
      cplh = IIf(IsNull(rs!PLH), "", rs!PLH)                '========>            HEP  / CAS
      cFLAG = IIf(IsNull(rs!Flag), "", rs!Flag)             '========>            0:Asist  / 1:Adm
      cDni = "": cApelP = "": cApelM = "": cnomb = ""
      cCargo = "": cDscCargo = "": cCategoria = "": cDscCategoria = ""
      '-Contenido del Directorio
      c_Costo = "": cDscCosto = "": cCodUpr = "": cDscUpr = ""
      c_Costo = IIf(IsNull(rs!CCosto), "", rs!CCosto)
      cCodUpr = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)
      c_actividad = "": cDActividad = ""
      cGrpGen = "": cGrpPer = ""
      sql = "SELECT * FROM Cpgeneral WHERE (baja is NULL) and tarjeta='" & cTarjeta & "'"
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs_d.EOF Then
         cApelP = IIf(IsEmpty(rs_d!apelpat), "", UCase(rs_d!apelpat))
         cApelM = IIf(IsEmpty(rs_d!apelmat), "", UCase(rs_d!apelmat))
         cnomb = IIf(IsEmpty(rs_d!nombres), "", UCase(rs_d!nombres))
         cDni = IIf(IsNull(rs_d!Documento), "", Trim(rs_d!Documento))
         csexo = IIf(IsNull(rs_d!sexo), "", rs_d!sexo)
         '--cplh = IIf(IsEmpty(rs_d!plh), "", rs_d!plh)                '========>            HEP  / CAS
         cGrpGen = "" + Trim(rs_d!grpgen)
         cGrpPer = "" + Trim(rs_d!grpPer)
         c_actividad = IIf(IsNull(rs_d!actividad), "", rs_d!actividad)
         '-Corrige contenidos caso Vacios
         c_Costo = IIf(c_Costo = "", rs_d!CCosto, c_Costo)
         cCodUpr = IIf(cCodUpr = "", rs_d!CodUpr, cCodUpr)
         cCargo = IIf(IsNull(rs_d!Cargo), "", Trim(rs_d!Cargo))
         cCategoria = IIf(IsNull(rs_d!Categoria), "", Trim(rs_d!Categoria))
      Else
         cTarjeta = "@@@@@"
      End If
      rs_d.Close
      If cTarjeta = "@@@@@" Then            '#####=>TARJETA
         '-Skip Registro Invalido
      Else
      
      Set rs_d = New Recordset
      rs_d.Open "SELECT * FROM CpActividad WHERE codigo='" & c_actividad & "'", Cnx, adOpenStatic, adLockReadOnly
      If Not rs_d.EOF Then
         cDActividad = Trim(rs_d!descripcion)
      End If
      rs_d.Close
          If cCargo <> "" Then
             Set rs_d = New Recordset
             rs_d.Open "SELECT * FROM CpCargos WHERE codigo='" & cCargo & "'", Cnx, adOpenStatic, adLockReadOnly
             If Not rs_d.EOF Then
                cDscCargo = Trim(rs_d!descripcion)
             End If
             rs_d.Close
          End If
          If cCategoria <> "" Then
             Set rs_d = New Recordset
             rs_d.Open "SELECT * FROM CpTablasAux WHERE grupo='CAT' and codigo='" + cCategoria + "'", Cnx, adOpenStatic, adLockReadOnly
             If Not rs_d.EOF Then
                cDscCategoria = Trim(rs_d!descripcion)
             End If
             rs_d.Close
          End If
      
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
          
      '-Establece Afecto X LACTANCIA
      '=============================
      cAplica = "N": nSonLacta = 0   '-caso parto multiple max aplica 2Hras
      TxIniLac = "": TxFinLac = ""
      If csexo = "F" Then
         '-Ubica lactancia al periodo
         Set rs_d = New Recordset
         sql = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "' order by fecnac"
         rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         Do While Not rs_d.EOF
            dftope = DateAdd("YYYY", 1, rs_d!fecnac)   'Tope Lactancia
            If (rs_d!fecnac >= DateValue(TxFecIni) And DateValue(TxFecFin) < dftope) Or (rs_d!fecnac < DateValue(TxFecIni) And DateValue(TxFecFin) >= dftope) Or (rs_d!fecnac < DateValue(TxFecIni) And DateValue(TxFecFin) <= dftope) Then
               '-Carga Parametros de Ajuste horario por Lactancia
               TxIniLac = rs_d!fecnac
               TxFinLac = rs_d!FecFIN
               cAplica = rs_d!Aplica
            End If
            rs_d.MoveNext
         Loop
         rs_d.Close
      End If
      
      cdiaant = "N"      'controla marcacion caso Cierra turno NOCHE (2dias  19:30-07:30)
      For i = dfi To dff  '--------->dias
         z = i - 1
         c_turno = Trim(IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4)))
         If Trim(c_turno) = "*" Or Trim(c_turno) = "L" Or Trim(c_turno) = "V" Or Trim(c_turno) = "R" Then
         Else
            fechaB = Format(primer + z, "yyyymmdd")
            dfentra = primer + z
            dfsale = primer + z
            t_urnoD = "": p0_ing = "": p0_sal = "": p0_ing1 = "": p0_sal1 = "" 'HORARIO ORIGINAL
            p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": c2dias = "N"
            m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
            nValorH = 0: nValorM = 0: ndscto = 0: nmintarde = 0: F_falta = False: nmotpersonal = 0: nDsctoMP = 0:  nComision = 0
            cvalor = "": cAuPrg = "": cDiahora = "": cTipo = "": choraini = "": chorafin = "": cDescuenta = ""
            cMantGuardia = "N"
            
            '-Establece Afecto X LACTANCIA
            '=============================
            cAplica = "N": nSonLacta = 0   '-caso parto multiple max aplica 2Hras
            TxIniLac = "": TxFinLac = ""
            If csexo = "F" Then
               '-Ubica lactancia al periodo
               Set rs_d = New Recordset
               sql = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "' order by fecnac"
               rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               Do While Not rs_d.EOF
                 dftope = DateAdd("YYYY", 1, rs_d!fecnac)   'Tope Lactancia
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
            'If cTarjeta = "01979" Then  '¿?
            '   MsgBox (Format(DateValue(TxIniLac), "dd/mm/yyyy") + " - " + Format(DateValue(TxFinLac), "dd/mm/yyyy"))
            'End If
            
            '-Tolerancia x LACTANCIA (I)ngreso/(S)alida
            '=======================
            nLacta = 0
            If cAplica <> "N" Then
               nLacta = IIf(nSonLacta = 1, 1, 2)
            End If
            'If cTarjeta = "01979" Then
            '   MsgBox "Aplica:" + cAplica + " [" + Trim(Str(nLacta)) + "] Dia:" + Format(primer + z, "dd/mm/yyyy") + " en " + Format(TxIniLac, "dd/mm/yyyy") + " - " + Format(TxFinLac, "dd/mm/yyyy")
            'End If
            
            '-HORARIO prog en dia
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
                  t_urnoD = IIf(IsNull(rs_d!turno), "", Trim(rs_d!turno))
                  p_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
                  If rs_d!Horas < "1800" Then                               '========>
                     p_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
                     p_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                  End If
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
                  'If InStr("07749", cTarjeta) > 0 Then
                  '   MsgBox (cTarjeta + " : " + cAplica + " => " + p_ing + "-" + p_sal + " ; " + p_ing1 + "-" + p_sal1)
                  'End If
                  
                  p_horas = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
                  nValorH = Val(Left(p_horas, 2))
                  nValorM = Val(Right(p_horas, 2))
                  If InStr("NHCP", t_urnoD) > 0 Then
                     c2dias = "S"        '1900-0700
                  End If
                  
               End If
            End If
            rs_d.Close
            
            '-Ubica AUTORIZACIONES: Motivos Personales, Comisiones Serv, vaca
            cDiahora = "X"
            sql = "SELECT * FROM CpAutori" & genaño & " WHERE (anulado is NULL) and   tarjeta='" & Trim(cTarjeta) & "' and FechaIni<='" + fechaB + "' and FechaFin>='" + fechaB + "'"     '   and turno=''"
            Set rs_d = New Recordset
            rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            Do While Not rs_d.EOF
               If Not rs_d.EOF Then
                  cAuPrg = Trim(rs_d!codigo)
                  cDiahora = rs_d!diahora                              'D/H
                  cTipo = IIf(IsNull(rs_d!tipo), "", rs_d!tipo)        'I/S/O   caso horas
                  choraini = IIf(IsNull(rs_d!horaini), "", rs_d!horaini)
                  chorafin = IIf(IsNull(rs_d!horafin), "", rs_d!horafin)
               End If
               If cAuPrg <> "" Then
                  sql = "SELECT * FROM CpOcurrencias WHERE codigo='" & cAuPrg & "'"
                  Set r_t = New Recordset
                  r_t.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                  If Not r_t.EOF Then
                     cDescuenta = IIf(r_t!DESCUENTO = "", "N", Trim(r_t!DESCUENTO))
                     cMantGuardia = IIf(IsNull(r_t!MantGuardia), "N", r_t!MantGuardia)
                  End If
                  r_t.Close
                  '-Aplica Autorización Entrada(I)/(S)alida
                  If cTipo = "I" Then
                     m_ing = choraini
                  ElseIf cTipo = "S" Then
                         m_sal = chorafin
                  End If
               End If
               rs_d.MoveNext
            Loop
            rs_d.Close
            
            If Val(p_horas) > 0 Then   '-------> (1)
               If cplh = "HEP" Then
                  nTNomb = nTNomb + 1
                  If cFLAG = "0" Then
                     ntAsiNomb = ntAsiNomb + 1
                  ElseIf cFLAG = "1" Then
                         ntAdmNomb = ntAdmNomb + 1
                  End If
               ElseIf cplh = "CAS" Then
                      nTCas = nTCas + 1
                      If cFLAG = "0" Then
                         ntAsiCas = ntAsiCas + 1
                      ElseIf cFLAG = "1" Then
                             ntAdmCas = ntAdmCas + 1
                      End If
               End If
                
               '-ACUMULA TOT. HRAS PROG del PERIODO
               nTotHprog = nTotHprog + (nValorH + (nValorM / 60))
            
               '-Marcas en Dia Programado
               Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
               '-Marcas del dia, solo si Justifica "D"
               p_Sale = ""
               If cDiahora <> "D" Then
                  If Val(p_sal1) > 0 Or Val(p_sal) < Val(p_ing) Then
                     '-Ubica marcas 2dias Actual-siguiente caso: turno t/n, n/d
                     fecha1 = Format(primer + z + 1, "yyyymmdd")
                     Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha >='" & fechaB & "' AND fecha <='" & fecha1 & "' ORDER BY fecha,hora"
                     cdiaant = "S"
                     p_Sale = p_sal
                     If Val(p_sal1) > 0 Then
                        p_Sale = p_sal1         'Corrige fin turno
                     End If
                  End If
               End If
               
               '-ASIGNA marcas basado en Horario   I - S  /I1 - S1
               Set rs_d = New Recordset
               rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not rs_d.EOF Then
                  d_fentra = rs_d!fecha            '========> 1era marca del servidor para turno programado
                  Do While Not rs_d.EOF
                     cInicio = p_ing
                     cFin = rs_d!hora
                     dfsale = rs_d!fecha
                     If Val(m_ing) = 0 Then
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
                 
                     ElseIf ((Val(m_sal) = 0 And (Val(rs_d!hora) - Val(p_sal) >= 0)) Or (Val(m_sal) = 0 And Val(p_sal) > Val(rs_d!hora))) Then
                            '07001900-19000700   0701 0738x 0740    1352 0712
                            m_sal = rs_d!hora
                 
                     ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) > 0 And Val(m_sal1) = 0 And (Val(rs_d!hora) - Val(p0_sal1) >= 0) Then       '<=60
                            '-sal Turno partido  1300-1900/1900-0700
                            m_sal1 = rs_d!hora
                        
                     '-Check carga marcaciones sin valor para turno
                     ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then     'And Val(m_ing1) > 0
                            m_sal1 = rs_d!hora
                 
                     End If
                     rs_d.MoveNext
                  Loop
               End If
               rs_d.Close
          
            End If      '<------ (1)
            
             '-Muestra  marcas realizadas
             'If InStr("07749", cTarjeta) > 0 Then
             '   MsgBox cTarjeta + ": Marcas [2dias->" + c2dias + "] " + Trim(Str(Day(dfentra))) + "/" + Trim(Str(Day(dfsale))) + " ->  Ing/sal :" + m_ing + " , " + m_sal + " / " + m_ing1 + " , " + m_sal1
             'End If
            
            ww = ""
            If Val(p_horas) > 0 Then   '------> (A)
            
               If (t_urnoD <> "N" And t_urnoD <> "C" And Not f_cierra) Or f_cierra Then '-----> (b)
               '-Evalua MARCAS  Ingreso
               If cDiahora = "D" And cDescuenta = "S" Then
                  F_falta = True
                  ww = ww + "1_"
               ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And Val(p_ing1) > 0 And p_sal <> p_ing1 Then  '======>
                      'Caso Horario:2000-0800   marca 0805-2010   O marca 0749-1404 1930-0820 ó  0803-2015/-0824
                      '-07001900-19000700 marca 0701 -07180719 & 1252,1900- ,07
                      ' 13001900-19000700       12521905- ,0715
                      F_falta = True
                      ww = ww + "2_"
               ElseIf (Val(m_ing) + Val(m_ing1) = 0) Then
                      '-No marco Ing,   valida vs AUTORIZA     check DM/MP
                      If cDiahora = "H" And cTipo = "I" Then
                         cInicio = p_ing
                         cFin = choraini
                         If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                            F_falta = True
                            cDescuenta = "S"
                            ww = ww + "3_"
                         End If
                      ElseIf cDiahora = "D" Then
                             If cDescuenta = "S" Then
                                F_falta = True
                                ww = ww + "4_"
                             End If
                      Else
                         F_falta = True
                         ww = ww + "5_"
                      End If
               ElseIf (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                      'Marco tarde  Autoriza H
                      If cDiahora = "H" And cTipo = "I" Then
                         cInicio = p_ing
                         cFin = choraini
                         If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                            F_falta = True
                            ww = ww + "6_"
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
                            '--               :   0812-1847        0900
                            '---  0700-1900 1900-0700  24Hras  :marcas   0715- / -0748
                            '---  1300-1900 1900-0700  18hrs   :marca    1252-   -0705
                            If (m_sal > m_ing And m_sal < p_ing) And p_horas <> 2400 Then   '¿2400?
                               '-salida Antes de Turno
                               F_falta = True
                               ww = ww + "7_"
                            Else
                               If c2dias = "S" And ((m_sal > p_sal And d_fentra = dfsale) Or (m_sal < p_sal And d_fentra < dfsale)) Then
                                  '07001900-19000700   0716, -07580759
                                  '-No completo el turno
                                  F_falta = True
                                  ww = ww + "8_"
                                  
                                  'If InStr("56775_59123_ZZ1085", cTarjeta) > 0 Then
                                  '   MsgBox cTarjeta + ": Marcas [2dias->" + c2dias + "] " + Trim(Str(Day(dfentra))) + "/" + Trim(Str(Day(dfsale))) + " ->  Horas: " + p_horas + " Ing/sal" + p_ing + "/" + p_sal + " marcas Ing/sal :" + m_ing + " , " + m_sal + " / " + m_ing1 + " , " + m_sal1
                                  'End If
                                  
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
                                  ww = ww + "9_"
                               End If
                            End If
                      End If
               End If
               End If      '<------ (b)
               
               
               If f_cierra Then         '-----> (c)
               '-Evalua MARCAS Salida
               If F_ALTA Then
               ElseIf (Val(m_sal) + Val(m_sal1) = 0) Then
                  'No marco  Sal   valida vs autoriza
                  If cDiahora = "H" And cTipo = "S" Then
                     If (Val(p_sal) - Val(chorafin) > 0) Or (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                        F_falta = True
                        cDescuenta = "S"
                        ww = ww + "s1_"
                     End If
                  ElseIf cDiahora = "D" Then
                         If cDescuenta = "S" Then
                            F_falta = True
                            ww = ww + "s2_"
                         End If
                  Else
                     F_falta = True
                     ww = ww + "s3_"
                  End If
               
               ElseIf Val(p_sal1) > 0 And p0_sal <> p0_ing1 Then    'p_horas<>"2400"
                      If Val(m_sal1) = 0 Then
                         F_falta = True
                         ww = ww + "s4_"
                      Else
                         If Val(p_sal1) > Val(m_sal1) Then
                            F_falta = True
                            ww = ww + "s5_"
                         End If
                      End If
                  
               ElseIf Val(m_sal) > 0 And Val(p_sal) > Val(m_sal) Then
                     'Marco antes de, Autoriza H
                     If cDiahora = "H" And cTipo = "S" Then    '====> And cTipo = "S"  ** TO CHECK **
                        cInicio = p_sal
                        cFin = chorafin
                        If Val(p_sal) - Val(chorafin) > 0 Then
                           F_falta = True
                           ww = ww + "s6_"
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
                               ww = ww + "s7_"
                        
                              'If InStr("56775_59123_ZZ1085", cTarjeta) > 0 Then
                              '   MsgBox cTarjeta + ": Marcas [2dias->" + c2dias + "] " + Trim(Str(Day(dfentra))) + "/" + Trim(Str(Day(dfsale))) + " Falt(" + ww + ") -> Horas: " + p_horas + " p_sal:" + p_sal + "salida Fin  :" + cFin
                              'End If
                            End If
                     End If
               End If
               End If          '<----- (c)
               
               If F_falta Then
               Else
                  If cDiahora = "H" And cTipo = "I" Then
                  ElseIf cDiahora <> "D" Then      '<===No Autorizaciones
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
                           ww = ww + "s8_"
                        End If
                     End If
                  End If
               End If
               If cDiahora = "H" Then
                  If F_falta Then
                  Else
                     If cDescuenta = "S" Then
                        'Hras Mot personales
                        nmotpersonal = Val(Left(chorafin, 2)) - Val(Left(choraini, 2)) + (Val(Right(chorafin, 2)) - Val(Right(choraini, 2))) / 60
                        If cAuPrg = "MP" Then
                           '-ACUMULA TOT. HRAS MP del PERIODO
                           ntotMP = ntotMP + nmotpersonal
                        End If
                        If cTipo = "I" Then
                           nDsctoMP = nmotpersonal
                           nvalor = (Val(Left(chorafin, 2)) - Val(Left(choraini, 2))) * 60 + Val(Right(chorafin, 2)) - Val(Right(choraini, 2))
                           If nvalor > 30 And nvalor < 60 Then
                              nDsctoMP = 1    'hora
                           Else
                              nDsctoMP = nmotpersonal
                           End If
                        End If
                     Else
                        nComision = Val(Left(chorafin, 2)) - Val(Left(choraini, 2)) + (Val(Right(chorafin, 2)) - Val(Right(choraini, 2))) / 60
                     End If
                  End If
               End If
            End If          '<------  (A)
            
            If F_falta Then
               nvalor = nValorH + (nValorM / 60)
            Else
               nvalor = nValorH + ((nValorM - ndscto) / 60)
               If cDiahora = "H" Then
                  If cDescuenta = "S" Then
                     'MotPersonal
                  Else
                     'Comision
                  End If
               End If
            End If
            
            'If InStr("07749", cTarjeta) > 0 Then
            '   MsgBox cTarjeta + "[" + c_turno + "]: " + IIf(F_falta, "F:" + ww, "T:" + Str(nvalor))
            'End If
            
            Sw = False
            
            If (F_falta Or cAuPrg = "MP" Or cAuPrg = "DM") And Val(p_horas) > 0 Then
               If cAuPrg = "DM" Then
                  '-ACUMULA TOT. HRAS DM del PERIODO
                  ntotDM = ntotDM + nvalor
               End If
               If cPresente = "N" Then
                  '-Falta
                  nposicion = nposicion + 1
                  Sw = True
               End If
            End If
            If Not F_falta And Val(p_horas) > 0 Then
               If cPresente = "S" Then
                  nTothTrab = nTothTrab + nvalor
                  nposicion = nposicion + 1
                  Sw = True
               End If
            End If
            
            If Sw Then
               rt.AddNew
               rt!tarjeta = cTarjeta
               rt!Dni = Trim(cDni)
               rt!apelpat = cApelP
               rt!apelmat = cApelM
               rt!nombres = cnomb
               rt!PLH = cplh
               rt!dscactividad = cDActividad
               rt!nombdia = Left(WeekdayName(Weekday(primer + z)), 3)
               rt!fecha = primer + z       '<==FECHA de falta
               rt!turno = c_turno          '<== Turno
             '  rt!horario = cvalor
               '-- Turno Original
               rt!ping = "" + p0_ing
              ' rt!psal = "" + p0_sal
             '  rt!ping1 = "" + p0_ing1
              ' rt!psal1 = "" + p0_sal1
               rt!hrasprog = 0 + (nValorH + nValorM / 60)  '<== HRAS PROG
               rt!ming = "" + m_ing
             '  rt!msal = "" + m_sal
              ' rt!ming1 = "" + m_ing1
             '  rt!msal1 = "" + m_sal1
                 rt!tardanza = 0 + nmintarde    ' 0 + IIf(F_falta, 0, nmintarde)
             '  rt!DESCUENTO = 0 + IIf(F_falta, 0, ndscto)
             '  rt!autorizac = "" + cAuPrg                 '<== Codigo AUTORIZACION
             '  rt!descuenta = "" + cDescuenta
             '  rt!diahora = "" + cDiahora
             '  rt!tipo = "" + cTipo
            '   rt!horaini = "" + choraini
             '  rt!horafin = "" + chorafin
               rt!falta = IIf(F_falta, "S", "N")            'INDICADOR FALTA
             '  If cAuPrg = "DM" Then
             '     rt!DM = (nValorH + nValorM / 60)
             '  End If
             '  rt!Comision = nComision
             '  rt!motpersonal = nmotpersonal                'IND MOT PERSONAL
             '  rt!DescuentoMP = 0 + IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", nDsctoMP, 0))
               rt!hrasfalt = IIf(F_falta, IIf(cDescuenta <> "N", (nValorH + (nValorM / 60)), 0), 0)
               rt!HrasTrab = IIf(F_falta, IIf(cDescuenta <> "N", 0, nvalor), IIf(p_ing <> "" Or cdiaant = "S", nvalor - nDsctoMP, 0))
               rt!Valido = IIf(p_ing <> "" Or cdiaant = "S", "S", "")
               rt!CCosto = c_Costo
               rt!dscccosto = Left(cDscCosto, 50)
               rt!CodUpr = cCodUpr
               rt!dscUPR = Left(cDscUpr, 55)
                    rt!Cargo = cCargo
                    rt!DscCargo = Trim(cDscCargo)
                    rt!Categoria = cCategoria
                    rt!dsccategoria = Trim(cDscCategoria)
               
               rt.Update
               ''nTothFalt = nTothFalt + IIf(F_falta, IIf(cDescuenta <> "N", (nValorH + (nValorM / 60)), 0), 0)
               nTothFalt = nTothFalt + rt!hrasfalt
               nson = nson + 1
               
               '-registros PEDIDOS
               If cplh = "HEP" Then
                  ndNomb = ndNomb + 1
                  If cFLAG = "0" Then
                     ndAsiNomb = ndAsiNomb + 1
                  ElseIf cFLAG = "1" Then
                         ndAdmNomb = ndAdmNomb + 1
                  End If
               ElseIf cplh = "CAS" Then
                      ndCas = ndCas + 1
                      If cFLAG = "0" Then
                         ndAsiCas = ndAsiCas + 1
                      ElseIf cFLAG = "1" Then
                             ndAdmCas = ndAdmCas + 1
                      End If
               End If

               
               'MsgBox Format(primer + z, "dd/mm/yyyy") + "-" + cTarjeta + ":" + Str(rt!hrasprog) + "-" + Str(rt!hrasfalt)
               
            End If
            If c2dias = "N" And cdiaant = "S" Then
               '-No turno 2dias y marca corresponde turno Dia Anterior
               cdiaant = "N"
            End If
         End If
      Next      '---------> nextdia
      End If    '<=##### TARJETA
      rs.MoveNext
   Loop
   rs.Close
   cmdcarga.Caption = ""
   Next k       '========>    (0)
   rt.Close
   TxTotHrasProg = Format(nTotHprog, "##.#0")
   TxHrasFalt = Format(nTothFalt, "##.#0")
   TxHrasMP = Format(ntotMP, "##.#0")
   TxHrasDM = Format(ntotDM, "##.#0")

   Frame1.Caption = IIf(cPresente = "S", "Presentes", "Ausentes")
   Lbld0hep.Caption = Format(ndAsiNomb, "###0")
   Lbld0Cas.Caption = Format(ndAsiCas, "###0")
   Lbld1hep.Caption = Format(ndAdmNomb, "###0")
   Lbld1cas.Caption = Format(ndAdmCas, "###0")
   If cPresente = "S" Then
     
      p0 = ndAsiNomb + ndAsiCas
      p1 = ndAdmNomb + ndAdmCas
      a0 = ntAsiNomb + ntAsiCas - p0
      a1 = ntAdmNomb + ntAdmCas - p1
   Else
      a0 = ndAsiNomb + ndAsiCas
      a1 = ndAdmNomb + ndAdmCas
      p0 = ntAsiNomb + ntAsiCas - a0
      p1 = ntAdmNomb + ntAdmCas - a1
   End If
   
   LblP0.Caption = Format(p0, "###0")
   Lbla0.Caption = Format(a0, "###0")
   Lblp1.Caption = Format(p1, "###0")
   Lbla1.Caption = Format(a1, "###0")
   LblPORC0.Caption = ""
   LblPorc1.Caption = ""
   If ntAsiNomb + ntAsiCas > 0 Then
     LblPORC0.Caption = Format(a0 * 100 / (ntAsiNomb + ntAsiCas), "###.#0")
   End If
   If ntAdmNomb + ntAdmCas > 0 Then
     LblPorc1.Caption = Format(a1 * 100 / (ntAdmNomb + ntAdmCas), "###.#0")
   End If
   Frame1.Visible = True
   
   If nTotHprog > 0 Then
      If cPresente = "S" Then
         TxPorcPerd = Format((nTothTrab / nTotHprog) * 100, "##.#0")
         MsgBox "Servidores PRESENTES:" + Str(nson) + " => Hras.Trab:" + Format(nTothTrab, "####.#0")
      Else
         TxPorcPerd = Format(((nTothFalt + ntotMP + ntotDM) / nTotHprog) * 100, "##.#0")
         MsgBox "Servidores Ausentes:" + Str(nson) + " => Hras.Falt:" + Format(nTothFalt, "####.#0") + " MP: " + Format(ntotMP, "##.#0") + " DM:" + Format(ntotDM, "##.#0")
      End If
   End If
   If nposicion > 0 Then
      Call CmdPrint_Click
   Else
      MsgBox "No Registra Información..", vbExclamation
   End If
   
   
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
   cmdcarga.Caption = ""
   If nson = 0 Then
      MsgBox "No registra resultados...", vbExclamation
      Exit Sub
   End If
   '-Faltas
   nTipRpt = 9
   InfoCpRS.Left = 0: InfoCpRS.Top = 0
   InfoCpRS.Refresh
   InfoCpRS.Show
   Exit Sub
End Sub
Private Sub CmdExportaXLS_Click()
    cmdcarga.Caption = ""
    If ChkXLS.Value = 0 Then
       wTpoQuery = "9"
       Call ExportTableToXLS_click
    Else
       wTpoQuery = "9.0"
       wNameXLS = "C:\RPT\CpFaltas.xls"
       Call Exportar_ADO_Excel
    End If
    Beep 500, 350
    MsgBox "Se generó C:\RPT\CpFaltas.xls", vbInformation
    Exit Sub
End Sub
Private Sub CmdClear_Click()
  Frame1.Visible = False
  TxFecIni.Text = Date - 1: TxFecFin.Text = Date - 1: TxIniLac.Text = "": TxFinLac.Text = ""
  CboTipoPla.Text = "": Txtipopla = ""
  c_turno = "": TxTotHrasProg = "": Option1(0).Value = False: Option1(1).Value = False: cPresente = ""
  ChkXLS.Value = 1
  ChkDATAP.Caption = "Data pedida": ChkDATAP.Value = 0
  nsonC = 0: nposicion = 0:
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


