VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RecMasRS 
   Caption         =   "SUMARIO de marcación del Periodo"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "RecMasRS.frx":0000
   ScaleHeight     =   4050
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExportaXLS 
      Height          =   735
      Left            =   7560
      Picture         =   "RecMasRS.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Exporta a Xls"
      Top             =   3000
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar Progreso1 
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   2160
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
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
      Left            =   120
      TabIndex        =   17
      Top             =   2640
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
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton mnuFicImpSelec 
      Height          =   675
      Left            =   2640
      Picture         =   "RecMasRS.frx":80EA8
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Selec.Impresora"
      Top             =   3000
      Width           =   765
   End
   Begin VB.TextBox TxUltDia 
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
      Left            =   7680
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox TxPridia 
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
      Left            =   7680
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Solo Activos"
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   360
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   9240
      Picture         =   "RecMasRS.frx":81686
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Otra Consulta"
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdCargaSUM 
      Height          =   615
      Left            =   9240
      Picture         =   "RecMasRS.frx":81F50
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ubicar Registros"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox Txtipopla 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6240
      MaxLength       =   3
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox CboTipoPla 
      Height          =   315
      Left            =   5520
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   9360
      Picture         =   "RecMasRS.frx":8212B
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cerrar"
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox Txccosto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4200
      MaxLength       =   2
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Cmdprint 
      Caption         =   "Imprime"
      Height          =   735
      Left            =   5040
      Picture         =   "RecMasRS.frx":826AD
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox TxPlh 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3240
      MaxLength       =   3
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   375
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
      Left            =   5400
      TabIndex        =   15
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10200
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Al :"
      Height          =   255
      Left            =   7200
      TabIndex        =   14
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Del :"
      Height          =   255
      Left            =   7200
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Planilla"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "RecMasRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim r_s As New ADODB.Recordset
Dim r_2 As New ADODB.Recordset
Dim rt1 As New ADODB.Recordset
Dim rRS As New ADODB.Recordset
Dim Criterio, f_iltro, cTarjeta, ceditor, VerDia, ccadena, cNombDia, cms As String
Dim ctrabajador, cDni, csexo, cAplica, chorasp, chtarda, chdscto, chorasmp, chorasfj, chorasfi, chorast As String
Dim nHrFalt, nHrProg, nDescuento, nDesCta, nHrTrab As Double
Dim nLacta, nposicion, wdato, nreg, msi, diai, msf, diaf As Integer
Dim unafecha As Date
Private Sub Form_Load()
  Cmdprint.Enabled = False
  Txmes = RecMarcasRS.Txmes
  For X = 1 To 12
    If Val(Txmes) = X Then
       RecMasRS.Caption = "SUMARIO de Marcación " + Trim(Mid(wMeses, (X - 1) * 9 + 1, 9)) + " - " + genaño
    End If
  Next X
  
  TxPridia = RecMarcasRS.TxPridia
  If DateValue(RecMarcasRS.TxUltDia) >= Date Then
     TxUltDia = Date - 1
  Else
     TxUltDia = DateValue(RecMarcasRS.TxUltDia) - 1
  End If
  
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
  'LLena el Combo TipoPLA
  CboTipoPla.Clear
  CboTipoPla.AddItem "ASISTENCIAL"
  CboTipoPla.AddItem "ADMINISTRATIVO"
  CboTipoPla.ListIndex = 0
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpCcosto ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
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
     End If
     rs.MoveNext
  Wend
  rs.Close
 
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpUndPrgRol ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
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
  TxPlh = "":  CboRegPlh.Text = ""
  Txtipopla = "":  CboTipoPla.Text = ""
  Txccosto = "":  CboCCosto.Text = ""
  TxCodUpr = "": CboUpr.Text = ""
  Lblrec.Caption = ""
  
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
     TxPlh.Text = Right(CboRegPlh.List(j), nlen - nposicion)
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
     Txtipopla.Text = Left(CboTipoPla.List(j), 3)
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
     Txccosto.Text = Right(CboCCosto.List(j), nlen - nposicion)
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
     TxCodUpr.Text = Right(CboUpr.List(j), nlen - nposicion)
  End If
End Sub
Private Sub CmdClear_Click()
    
  Set rs = Nothing
  Set rRS = Nothing
  Set rt = Nothing
  Set rt1 = Nothing
  Set DD = Nothing
  
  TxPlh.Text = "":  CboRegPlh.Text = ""
   Txtipopla.Text = "":  CboTipoPla.Text = ""
   Txccosto.Text = "":  CboCCosto.Text = ""
   TxCodUpr.Text = "": CboUpr.Text = ""
   Check1.Value = 1
   Lblrec.Caption = ""
   wdato = 0
End Sub


Private Sub TxUltDia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If DateValue(TxUltDia) < DateValue(TxPridia) Or Year(TxUltDia) <> Year(TxPridia) Then
        MsgBox "<E> Verifique fechas de consulta. ", vbExclamation
        TxUltDia = ""
        Exit Sub
     End If
     msi = Month(TxPridia)
     msf = Month(TxUltDia)
  End If
End Sub

Private Sub TxUltDia_LostFocus()
     TxUltDia_KeyPress (13)
End Sub

Private Sub cmdCargaSUM_Click()
   
  If Month(TxPridia) <> Month(TxUltDia) Then
     Me.Caption = "SUMARIO de Marcación del Periodo"
     Me.Refresh
  End If
   msi = Month(TxPridia)
   diai = Day(TxPridia)
   msf = Month(TxUltDia)
   diaf = Day(TxUltDia)
   Lblrec.Caption = ""
   wdato = 0
   Cmdprint.Enabled = False
   If Wccosto <> "" And TxCodUpr = "" Then
      MsgBox "Defina UPR...", vbExclamation
      Exit Sub
   End If
   ceditor = "X"
   If Wccosto = "" Then
      ceditor = "S"
   Else
      '-Valida derechos en Upr
      npos = InStr(wCadenaUpr, TxCodUpr)
      If npos > 0 Then
         ceditor = Right(Mid(wCadenaUpr, npos, 3), 1)
      End If
   End If
   If ceditor = "X" Then
      MsgBox "Acceso, No autorizado...", vbExclamation
      Exit Sub
   End If
   
   Criterio = " (cpgeneral.baja is NULL) "
   f_iltro = ""
   If Check1.Value = 1 Then        'activo
      If Len(Criterio) > 0 Then
          Criterio = Criterio + " and ( cpgeneral.estado='A')"
          f_iltro = f_iltro & ", Activos "
      Else
          Criterio = " ( cpgeneral.estado='A')"
          f_iltro = " Activos "
      End If
   End If
   If Trim(TxPlh) <> "" Then
      If Len(Criterio) > 0 Then
          Criterio = Criterio + " and ( cpgeneral.plh='" & TxPlh & "')"
          f_iltro = f_iltro & ", PLH "
      Else
          Criterio = " ( cpgeneral.plh='" & TxPlh & "')"
          f_iltro = " PLH "
      End If
   End If
   If Txtipopla <> "" Then
      If Len(Criterio) > 0 Then
        Criterio = Criterio + " and  ( cpgeneral.Tipopla='" & Txtipopla & "')"
        f_iltro = f_iltro & ", Tip.Pla. "
      Else
        Criterio = "(cpgeneral.Tipopla='" & Txtipopla & "')"
        f_iltro = "Tip.Pla. "
      End If
   End If
   If Txccosto <> "" Then
      If Len(Criterio) > 0 Then
        Criterio = Criterio + " and  (cpgeneral.ccosto='" & Txccosto & "')"
        f_iltro = f_iltro & ", C.Costo "
      Else
        Criterio = " (cpgeneral.ccosto='" & Txccosto & "')"
        f_iltro = "C.Costo "
      End If
   End If
   If TxCodUpr <> "" Then
      If Len(Criterio) > 0 Then
         Criterio = Criterio + " and  (cpgeneral.codupr='" & TxCodUpr & "')"
         f_iltro = f_iltro & ", UPR "
      Else
         Criterio = " (cpgeneral.codupr='" & TxCodUpr & "')"
         f_iltro = "UPR "
      End If
   End If
   If Criterio = "" Then
      MsgBox "No definio Criterio ", vbExclamation
      Exit Sub
   End If
   
  Cnx.Execute "DELETE  FROM CpTmGeneral"

    On Error GoTo Err0
    Cnx.Execute "DELETE  FROM #Tabla5"
     GoTo sigue0
Err0:
    Cnx.Execute ("SELECT * Into #Tabla5  FROM CpTmRecSum WHERE 1 = 2")
sigue0:
   
  On Error GoTo Err1
  Cnx.Execute ("SELECT * Into #tabla1  FROM CpTmRecord WHERE 1 = 2")
  GoTo sigue1
Err1:
  Cnx.Execute "DELETE  FROM #tabla1"
sigue1:

   Set rRS = New Recordset
   sqRS = "SELECT * FROM #Tabla5"    'SUMARIO marcacion -> export XLS
   rRS.Open sqRS, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   
   Set rt = New Recordset
   sq_l = "SELECT * FROM CpTmGeneral"
   rt.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   rt.Close
   
   Set rt1 = New Recordset
   sql = "SELECT * FROM #tabla1"
   rt1.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   
   Set r_d = New ADODB.Recordset
   sql = "SELECT * FROM CpGeneral WHERE " & Criterio & " ORDER BY ccosto,codUpr,apelpat,apelmat"
   r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If r_d.EOF Then
      Exit Sub
   End If
   nreg = r_d.RecordCount
   
   Progreso1.Value = 0
   Progreso1.Max = nreg
    Progreso1.Visible = True
   r_d.MoveFirst
   Cnx.BeginTrans
   
   Do While Not r_d.EOF         '=====> Servidores
      With Progreso1
        If .Value < .Max Then
           .Value = .Value + 1
        End If
      End With
      
      cTarjeta = r_d!tarjeta
      csexo = IIf(IsNull(r_d!sexo), "", Trim(r_d!sexo))
      cDni = IIf(IsNull(r_d!Documento), "", Trim(r_d!Documento))
      cActividad = r_d!actividad
      cCargo = IIf(IsNull(r_d!cargo), "", Trim(r_d!cargo))
      ccCosto = IIf(IsNull(r_d!cCosto), "", Trim(r_d!cCosto))
      cCodUpr = IIf(IsNull(r_d!CodUpr), "", Trim(r_d!CodUpr))
      wdsccosto = "": wdscUpr = ""
      Set rs = New Recordset
      rs.Open "SELECT * FROM Cpccosto WHERE codigo='" & ccCosto & "'", Cnx, adOpenStatic, adLockReadOnly
      If Not rs.EOF Then
         wdsccosto = Left(rs!descripcion, 50)
      End If
      rs.Close
      If cCodUpr <> "" Then
         Set rs = New Recordset
         rs.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs.EOF Then
            wdscUpr = Left(rs!descripcion, 55)
         End If
         rs.Close
      End If
      '- Totales acumulado  por servidor en el periodo
      nThorasP = 0:  nTtarda = 0: nTDscto = 0
      nThorasMP = 0: nTotros = 0: nThorasFj = 0: nThorasFi = 0: ntHorasFMP = 0: nthorasT = 0
      
      diai = Day(TxPridia)    'Dias de inicio y cierre del periodo consultado
      diaf = Day(TxUltDia)    'Day(ultimo)
      ''''dff0 = Day(ultimo)  ultimo dia calendario en cada mes
      
      '-Establece Afecto X LACTANCIA
      '=============================
      cAplica = "N": nSonLacta = 0
      TxIniLac = "": TxFinLac = ""
      If csexo = "F" Then
         '-Ubica lactancia al periodo
         Set rs = New Recordset
         sql = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "'  order by fecnac"
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         Do While Not rs.EOF
            dftope = DateAdd("YYYY", 1, rs!fecnac)   'Tope Lactancia
            'MsgBox Format(dftope, "dd/mm/yyyy")
            If (rs!fecnac >= DateValue(TxPridia) And DateValue(TxUltDia) < dftope) Or (rs!fecnac < DateValue(TxPridia) And DateValue(TxUltDia) >= dftope) Or (rs!fecnac < DateValue(TxPridia) And DateValue(TxUltDia) <= dftope) Then
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
      
      For j = msi To msf      '====> directorios del SERVIDOR
         '-Carga  DIRECTORIO "TURNO"
         cms = Right("00" + Trim(Str(j)), 2)
         sql = "SELECT * FROM CpDirectorio  WHERE tarjeta='" & cTarjeta & "' AND ano ='" & genaño & "' AND mes='" & cms & "'"
         
            'If wdato < 5 Then
            '    MsgBox sql
            'End If
         
         Set r_s = New Recordset
         r_s.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If r_s.EOF Then
            '-Servidores s/programacion
            ccadena = ccadena + (cTarjeta + "_")
         Else
            cdiaant = "N"         'marcacion caso Cierra turno NOCHE (2dias  19:30-07:30)
            p_Sale = ""
            
            unafecha = "01/" + cms + "/" + Trim(Str(genaño))
            Call ObtenerPrimerUltimoDia(unafecha)
            dff0 = Day(ultimo)
            If j = msi Then
               dfi = diai
            Else
               dfi = 1
            End If
            If j = msf Then
               dff = diaf
            Else
               dff = dff0
            End If
        
            For i = dfi To dff      '=====================>>dias calendario mes analizar, max 31
                z = i - 1
                If i <= dff0 Then      '---Y  Marcas hasta Dia ultimo de mes dff0
                   VerDia = IIf(i <= dff, "S", "N")
                   FechaB = Format(primer + z, "yyyymmdd")
                   dfentra = primer + z
                   dfsale = primer + z
                   cTurno = ""
                   p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": c2dias = "N"
                   m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
                   nValorH = 0: nValorM = 0: nHrProg = 0: nmintarde = 0: ndscto = 0: F_falta = False: nmotpersonal = 0: nDsctoMP = 0: notros = 0: nComision = 0
                   cvalor = "": cAuPrg = "": cDiahora = "": cTipo = "": choraini = "": chorafin = "": cDescuenta = "": cMantGuardia = ""
                   p0_ing = "": p0_sal = "":  p0_ing1 = "":  p0_sal1 = ""
             
                    '-Tolerancia x LACTANCIA (I)ngreso/(S)alida
                    '=======================
                   nLacta = 0
                   If cAplica <> "N" Then
                      If DateValue(primer + z) >= DateValue(TxIniLac) And DateValue(primer + z) <= DateValue(TxFinLac) Then
                         nLacta = IIf(nSonLacta = 1, 1, 2)
                      End If
                   End If
                    '' MsgBox "Aplica:" + cAplica + " [" + Trim(Str(nLacta)) + "] Dia:" + Format(primer + z, "dd/mm/yyyy") + " en " + Format(TxIniLac, "dd/mm/yyyy") + " - " + Format(TxFinLac, "dd/mm/yyyy")
                              
                   If VerDia = "S" Then      '----X
                      cTurno = IIf(IsNull(r_s.Fields(z + 4)), "", r_s.Fields(z + 4))
                      '-HORARIO prog en dia
                      sql = "SELECT * FROM CpHor" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & FechaB & "'"
                      Set rs = New Recordset
                      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                      If rs.EOF Then
                         If Trim(cGrpGen) <> "" Then
                            rs.Close
                            sql = "SELECT * FROM CpHor" & genaño & " WHERE grupo='" & cGrpGen & "' AND fecha ='" & FechaB & "'"
                            Set rs = New Recordset
                            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                         End If
                      End If
                      If Not rs.EOF Then
                         cvalor = IIf(IsNull(rs!horario), "", rs!horario)
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
                            If rs_d!Horas < "1800" Then                               '========>
                               p_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                               p_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
                            End If
                            p_sal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
                                                
                            '-horario ORIGINAL
                            '=================
                            p0_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
                            ''If rs_d!Horas <> "2400" And rs_d!Horas <> "1800" Then                               '=======>
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
                            nHrProg = (nValorH + nValorM / 60)
                            nThorasP = nThorasP + nHrProg       'acumula Hrs. Prog
                            If (p0_ing > p0_sal And Val(p0_ing1) = 0) Then          '======> Or p_horas = "2400"
                                c2dias = "S"
                            ElseIf Val(p0_sal1) > 0 And p0_ing < p0_ing1 Then
                                c2dias = "S"
                            End If
                         End If
                         rs_d.Close
                      End If
                      rs.Close
            
                      '-AUTORIZACIONES: Motivos Personales, Comisiones Serv
                      cDiahora = "X"
                      cMantGuardia = "N"
                      sql = "SELECT * FROM CpAutori" & genaño & " WHERE tarjeta='" & cTarjeta & "' and FechaIni<='" + FechaB + "' and FechaFin>='" + FechaB + "'"
                      Set rs = New Recordset
                      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                      Do While Not rs.EOF
                         If Not rs.EOF Then
                            cAuPrg = Trim(rs!Codigo)
                            cDiahora = rs!diahora
                            cTipo = IIf(IsNull(rs!tipo), "", rs!tipo)            'I/S/O   caso horas
                            choraini = IIf(IsNull(rs!horaini), "", rs!horaini)
                            chorafin = IIf(IsNull(rs!horafin), "", rs!horafin)
                         End If
                         If cAuPrg <> "" Then
                            sql = "SELECT * FROM CpOcurrencias WHERE codigo='" & cAuPrg & "'"
                            Set r_2 = New Recordset
                            r_2.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                            If Not r_2.EOF Then
                               cDescuenta = IIf(r_2!DESCUENTO = "", "N", Trim(r_2!DESCUENTO))
                               cMantGuardia = IIf(IsNull(r_2!MantGuardia) Or Trim(r_2!MantGuardia) = "", "N", Trim(r_2!MantGuardia))
                            End If
                            r_2.Close
                         End If
                         rs.MoveNext
                         '-Aplica Autorización Entrada(I)/(S)alida
                         If cTipo = "I" Then
                            m_ing = choraini
                         ElseIf cTipo = "S" Then
                                m_sal = chorafin
                         End If
                      Loop
                      rs.Close
                      If Val(p_horas) > 0 Then
                         '-Marcas de Turno PROGRAMADO
                         Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & FechaB & "' ORDER BY fecha,hora"
                         '-Marcas del dia, solo si Justifica "D"
                         p_Sale = ""
                         If cDiahora <> "D" Then
                            If Val(p_sal1) > 0 Or Val(p_sal) < Val(p_ing) Then
                               '-Ubica marcas 2dias Actual-siguiente caso: turno t/n, n/d
                               fecha1 = Format(primer + z + 1, "yyyymmdd")
                               Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha >='" & FechaB & "' AND fecha <='" & fecha1 & "' ORDER BY fecha,hora"
                               cdiaant = "S"
                               p_Sale = p_sal
                               If Val(p_sal1) > 0 Then
                                  p_Sale = p_sal1        'fin turno
                               End If
                            End If
                         End If
                     
                         '-ASIGNA marcas basado en Horario    I - S  /I1 - S1
                         Set rs_d = New Recordset
                         rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                         If Not rs_d.EOF Then
                            d_fentra = rs_d!fecha      '========> 1era marca del servidor para turno programado
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
                                   
                               ElseIf (Val(m_sal) = 0 And (Val(rs_d!hora) - Val(p_sal) >= 0)) Or (Val(m_sal) = 0 And Val(p_sal) > Val(rs_d!hora)) Then
                                      '07001900-19000700   0701 0738x 0740
                                      m_sal = rs_d!hora
                               ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) > 0 And Val(m_sal1) = 0 And (Val(rs_d!hora) - Val(p0_sal1) >= 0) Then
                                      '-sal Turno partido  1300-1900/1900-0700
                                      m_sal1 = rs_d!hora
                               'To check carga marcaciones sin valor para turno
                               ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 And Val(m_ing1) > 0 Then
                                      m_sal1 = rs_d!hora
                                  
                               End If
                               rs_d.MoveNext
                            Loop
                         End If
                         rs_d.Close
                         '-Evalua MARCAS Ingreso
                         If cDiahora = "D" And cDescuenta = "S" Then
                            F_falta = True
                         ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And p_sal <> p_ing1 Then      '======>
                                'Caso Horario:2000-0800   marca 0805-2010   O marca 0749-1404 1930-0820
                                F_falta = True
                         ElseIf (Val(m_ing) + Val(m_ing1) = 0) Then
                                'No marco Ing,   valida vs autoriza
                                If cDiahora = "H" And cTipo = "I" Then
                                   cInicio = p_ing
                                   cFin = choraini
                                   If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                                      F_falta = True
                                      cDescuenta = "S"
                                   End If
                                ElseIf cDiahora = "D" Then      'valida autoriza Dia Comp/ pierde
                                       If cDescuenta = "S" Then
                                          F_falta = True
                                       End If
                                Else
                                   F_falta = True
                                End If
                         ElseIf (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                                'Marco tarde  Autoriza H al ingreso
                                If cDiahora = "H" And cTipo = "I" Then
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
                                    If (m_sal > m_ing And m_sal < p_ing) And p_horas <> 2400 Then   '¿2400?
                                       '-salida Antes de Turno
                                       F_falta = True
                                    Else
                                       If c2dias = "S" And ((m_sal > p_sal And d_fentra = dfsale) Or (m_sal < p_sal And d_fentra < dfsale)) Then
                                          '-No completo turno
                                          F_falta = True
                                       End If
                                    End If
                                Else
                                    '-Marca DOBLE
                                    '  PI   PS            MI   MS     MI1  MS1
                                    '- 2000-0800      :   0740-1401   1930-0830 ok
                                    '-                :   0743-1352   1400-1848 X
                                    '-                :   1750-0820   1200-1400
                                    '                 :   0745-2001   0746
                                    '-                :   1741-0801   1742
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
                         '-Evalua MARCAS Salida
                         If F_ALTA Then
                         ElseIf (Val(m_sal) + Val(m_sal1) = 0) Then
                                '-NO marco  Sal   valida vs autoriza
                                If cDiahora = "H" And cTipo = "S" Then
                                   If (Val(p_sal) - Val(chorafin) > 0) Or (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                                      F_falta = True
                                      cDescuenta = "S"
                                   End If
                                ElseIf cDiahora = "D" Then      'valida autoriza Dia Comp/ pierde
                                       If cDescuenta = "S" Then
                                          F_falta = True
                                       End If
                                Else
                                    F_falta = True
                                End If
                          ElseIf Val(p_sal1) > 0 And p0_sal <> p0_ing1 Then    'p_horas<>"2400"
                                If Val(m_sal1) = 0 Then
                                    F_falta = True
                                    wwind = "@"
                                Else
                                    If Val(p_sal1) > Val(m_sal1) Then
                                         F_falta = True
                                         wwind = "#"
                                    End If
                                 End If
                         ElseIf Val(m_sal) > 0 And Val(p0_sal) > Val(m_sal) Then      'p_sal
                                'Marco antes de, Autoriza H
                                If cDiahora = "H" And cTipo = "S" Then     '====> And cTipo = "S"  ** TO  CHECK **
                                   cInicio = p_sal
                                   cFin = chorafin
                                   If Val(p_sal) - Val(chorafin) > 0 Then
                                      F_falta = True
                                   End If
                                ElseIf cDiahora = "X" Then
                                       '------cInicio = p_sal
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
                                       ElseIf Val(p_sal1) - Val(cFin) > 0 Then
                                               F_falta = True
                                           
                                       End If
                                End If
                         End If
                     
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
                                      If nmintarde >= 6 And nmintarde < 31 Then
                                         If nmintarde >= 6 And nmintarde <= 10 Then
                                            ndscto = 10
                                         End If
                                         If nmintarde >= 11 And nmintarde <= 20 Then
                                            ndscto = 20
                                         End If
                                         If nmintarde >= 21 And nmintarde < 31 Then
                                            ndscto = 30
                                         End If
                                         '-acumula tardanza  /acumula tm30
                                         nTtarda = nTtarda + nmintarde
                                         nTDscto = nTDscto + ndscto
                                      Else
                                        If nmintarde > 30 Then
                                            F_falta = True
                                        End If
                                        nmintarde = 0
                                      End If
                                   End If
                            End If
                      End If
                  Else
                     '-MARCAS en Dia NO PROGRAMADO      fecha  hora =>  hhmm
                     '- Retira marca final de turno Dia Anterior, luego debe MOSTRAR
                     '-If cdiaant = "N" Then
                     Sq0 = "SELECT * FROM CpMarca" & genaño & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & FechaB & "' ORDER BY fecha,hora"
                     Set rs_d = New Recordset
                     rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                     If Not rs_d.EOF Then
                        Do While Not rs_d.EOF
                           If (c2dias = "N" And p_Sale = "") Or (p_Sale <> "" And nTotMints(p_Sale, rs_d!hora) > 120) Then
                              cInicio = m_ing
                              cFin = rs_d!hora
                              If Trim(m_ing) = "" Or (Val(m_ing) > 0 And (Val(rs_d!hora) - Val(m_ing) < 5)) Then
                                 m_ing = rs_d!hora
                              ElseIf Trim(m_ing) <> "" And CalcMints(cInicio, cFin) > 60 Then
                                     m_sal = rs_d!hora
                              ElseIf Trim(m_sal) <> "" Or (Val(rs_d!hora) - Val(m_sal) < 5) Then
                                     m_ing1 = rs_d!hora
                              End If
                           End If
                           rs_d.MoveNext
                        Loop
                     End If
                     p_Sale = ""    'apaga efecto turno anterior N
                     rs_d.Close
                  End If
                  If F_falta Then
                     If cAuPrg <> "" And cAuPrg <> "FI" And cDiahora = "D" And cDescuenta = "S" Then
                        If cAuPrg = "MP" Then
                           ntHorasFMP = ntHorasFMP + nHrProg             'Acumula hrs Faltas MP
                           'nDMP = nDMP + 1
                           'If InStr("D_N", Left(Lbltur(z).Caption, 1)) > 0 Then
                           '   nDMP = nDMP + 1
                           'End If
                        Else
                           nThorasFj = nThorasFj + nHrProg               'acumula
                           'nDFJ = nDFJ + 1
                           'If InStr("D_N", Left(Lbltur(z).Caption, 1)) > 0 Then
                           '   nDFJ = nDFJ + 1
                           'End If
                        End If
                     Else
                        If cAuPrg = "" Or cAuPrg = "FI" Or cDiahora = "H" Then
                           '-Corrige Descuento
                           cDescuenta = "S"
                           nThorasFi = nThorasFi + nHrProg               'acumula
                           'nDFI = nDFI + 1
                           'If InStr("D_N", Left(Lbltur(z).Caption, 1)) > 0 Then
                           '   nDFI = nDFI + 1
                           'End If
                        End If
                     End If
                  Else
                     If p_ing <> "" Or cdiaant = "S" Then
                        nthorasT = nthorasT + nHrProg
                        nvalor = nHrProg - (ndscto / 60)
                        If cDiahora = "H" Then
                           If cDescuenta = "S" Then
                              If cAuPrg = "MP" Then
                                 cInicio = choraini
                                 cFin = chorafin
                                 nmotpersonal = CalcMints(cInicio, cFin)
                                 If nmotpersonal < 0 And c2dias = "S" Then
                                    nmotpersonal = nmotpersonal + 24
                                 End If
                                 If nmotpersonal Mod 60 > 30 Then
                                    nDsctoMP = Int(nmotpersonal / 60) + 1
                                 Else
                                    nDsctoMP = Int(nmotpersonal / 60)
                                 End If
                                 nThorasMP = nThorasMP + nDsctoMP
                                 nthorasT = nthorasT - nDsctoMP
                              Else
                                 notros = Val(Left(chorafin, 2)) - Val(Left(choraini, 2)) + (Val(Right(chorafin, 2)) - Val(Right(choraini, 2))) / 60
                                 If notros < 0 And c2dias = "S" Then
                                    notros = notros + 24
                                 End If
                                 nthorasT = nthorasT - notros
                              End If
                           Else
                              nComision = Val(Left(chorafin, 2)) - Val(Left(choraini, 2)) + (Val(Right(chorafin, 2)) - Val(Right(choraini, 2))) / 60
                              If nComision < 0 And c2dias = "S" Then
                                 nComision = nComision + 24
                              End If
                           End If
                        End If
                     End If
                  End If
               End If       '----X
               cNombDia = Left(WeekdayName(Weekday(primer + z)), 3)
            Else
               'Dias exceso -> lineas vacias
               cNombDia = "X."
               VerDia = "N"
            End If        '---Y
        
            rt1.AddNew
            rt1!tarjeta = cTarjeta
            rt1!nombdia = cNombDia
            rt1!fecha = primer + z
            If VerDia = "S" Then
               rt1!turno = cTurno
               rt1!horario = cvalor
               '-- Turno Original
               rt1!ping = "" + p0_ing
               rt1!psal = "" + p0_sal
               rt1!ping1 = "" + p0_ing1
               rt1!psal1 = "" + p0_sal1
               rt1!hrasprog = nHrProg
               rt1!HrProg = IIf(nHrProg = 0, "", cHorsMint(nHrProg))
               rt1!ming = "" + m_ing
               rt1!msal = "" + m_sal
               rt1!ming1 = "" + m_ing1
               rt1!msal1 = "" + m_sal1
               rt1!marca2 = CStr(IIf(Val(m_ing1) > 0, m_sal, ""))
               rt1!marca3 = CStr(IIf(Val(m_ing1) = 0, m_sal, m_ing1))
               rt1!marca4 = CStr(IIf(Val(m_sal1) > 0, m_sal1, ""))
               rt1!tardanza = 0 + IIf(F_falta, 0, nmintarde)
               rt1!DESCUENTO = 0 + IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", ndscto, 0))
               nDescuento = IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", ndscto, 0))
               rt1!autorizac = "" + cAuPrg
               rt1!descuenta = "" + cDescuenta
               rt1!diahora = "" + cDiahora
               rt1!tipo = "" + cTipo
               rt1!horaini = "" + choraini
               rt1!horafin = "" + chorafin
               rt1!HrIni = IIf(choraini = "", "", Left(choraini, 2) + ":" + Right(choraini, 2))
               rt1!HrFin = IIf(chorafin = "", "", Left(chorafin, 2) + ":" + Right(chorafin, 2))
               rt1!falta = IIf(F_falta, "S", "N")
               rt1!Comision = nComision
               rt1!motpersonal = nmotpersonal
               rt1!DescuentoMP = 0 + IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", nDsctoMP, 0))
               rt1!DatoMp = IIf(cAuPrg <> "" And cAuPrg <> "FI", IIf(nmotpersonal = 0, IIf(nComision > 0, cHorsMint(nComision), ""), cHorsMint(nmotpersonal)), "")
               rt1!HrasTrab = IIf(F_falta, IIf(cDescuenta <> "N", 0, nvalor), IIf(p_ing <> "" Or cdiaant = "S", nvalor - nDsctoMP - notros, 0))
               nHrTrab = IIf(F_falta, IIf(cDescuenta = "S", 0, nvalor), IIf(p_ing <> "" Or cdiaant = "S", nvalor - nDsctoMP - notros, 0))
               rt1!HrTrab = IIf(nHrTrab > 0, cHorsMint(nHrTrab), "")
               rt1!hrasfalt = IIf(F_falta, IIf(cDescuenta <> "N", (nValorH + (nValorM / 60)), 0), 0)
               nHrFalt = IIf(F_falta, IIf(cDescuenta = "S", nHrProg, 0), 0)
               c_Falta = "Falta"
               If nHrTrab = 0 And cDiahora = "D" Then
                  '-autorizado "D" faltar
                  c_Falta = ""
               End If
               rt1!Descta = IIf(nHrFalt > 0, IIf(primer + z < Date, c_Falta, "* P *"), IIf(nDescuento > 0, CStr(Round(nDescuento, 2)), ""))
               'old  rt1!Descta = IIf(nHrFalt > 0, IIf(primer + z < Date, "Falta", " *P* "), IIf(nDescuento > 0, cHorsMint(nDescuento), ""))
               rt1!Valido = IIf(p_ing <> "" Or cdiaant = "S", "S", "")
               '-Campo GRUPO: usado para guardar si fecha esta en periodo LACTANCIA
               rt1!grupo = IIf(nLacta = 0, "", IIf(nLacta = 1, "LACT", "LAC2"))
            End If
            rt1.Update
            
            If VerDia = "S" Then
               rRS.AddNew
               rRS!tarjeta = cTarjeta
               rRS!Dni = cDni
               rRS!trabajador = Left(Trim(r_d!apelpat) + " " + Trim(r_d!apelmat) + " " + Trim(r_d!nombres) + IIf(nLacta = 0, "", IIf(nLacta = 1, " <LAC>", " <LAC2>")), 45)
               rRS!dscccosto = wdsccosto
               rRS!dscUPR = wdscUpr
               rRS!nombdia = cNombDia
               rRS!fecha = primer + z
               If cTurno <> "" Then
                  rRS!turno = cTurno
                  rRS!horario = cvalor
                  '-- Turno Original
                  rRS!ping = "" + p0_ing
                  rRS!psal = "" + p0_sal
                  rRS!ping1 = "" + p0_ing1
                  rRS!psal1 = "" + p0_sal1
                  rRS!hrasprog = nHrProg
                  rRS!HrProg = IIf(nHrProg = 0, "", cHorsMint(nHrProg))
                  rRS!ming = "" + m_ing
                  rRS!msal = "" + m_sal
                  rRS!ming1 = "" + m_ing1
                  rRS!msal1 = "" + m_sal1
                  rRS!tardanza = 0 + IIf(F_falta, 0, nmintarde)
                  rRS!DESCUENTO = 0 + IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", ndscto, 0))
                  nDescuento = IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", ndscto, 0))
                  rRS!autorizac = "" + cAuPrg
                  rRS!descuenta = "" + cDescuenta
                  rRS!diahora = IIf(InStr("DH", cDiahora) > 0, cDiahora, "")
                  rRS!tipo = "" + cTipo
                  rRS!horaini = "" + choraini
                  rRS!horafin = "" + chorafin
                  rRS!HrIni = IIf(choraini = "", "", Left(choraini, 2) + ":" + Right(choraini, 2))
                  rRS!HrFin = IIf(chorafin = "", "", Left(chorafin, 2) + ":" + Right(chorafin, 2))
                  rRS!falta = IIf(F_falta, "S", "N")
                  rRS!Comision = nComision
                  rRS!motpersonal = nmotpersonal
                  rRS!DescuentoMP = 0 + IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", nDsctoMP, 0))
                  rRS!DatoMp = IIf(cAuPrg <> "" And cAuPrg <> "FI", IIf(nmotpersonal = 0, IIf(nComision > 0, cHorsMint(nComision), ""), cHorsMint(nmotpersonal)), "")
                  rRS!HrasTrab = IIf(F_falta, IIf(cDescuenta <> "N", 0, Round(nvalor, 2)), IIf(p_ing <> "" Or cdiaant = "S", Round(nvalor - nDsctoMP - notros, 2), 0))
                  nHrTrab = IIf(F_falta, IIf(cDescuenta = "S", 0, nvalor), IIf(p_ing <> "" Or cdiaant = "S", nvalor - nDsctoMP - notros, 0))
                  rRS!HrTrab = IIf(nHrTrab > 0, cHorsMint(nHrTrab), "")
                  rRS!hrasfalt = IIf(F_falta, IIf(cDescuenta <> "N", (nValorH + (nValorM / 60)), 0), 0)
                  nHrFalt = IIf(F_falta, IIf(cDescuenta = "S", nHrProg, 0), 0)
                  c_Falta = "Falta"
                  If nHrTrab = 0 And cDiahora = "D" Then
                     '-autorizado "D" faltar
                     c_Falta = ""
                  End If
                  rRS!Descta = IIf(nHrFalt > 0, IIf(primer + z < Date, c_Falta, "* P *"), IIf(nDescuento > 0, CStr(Round(nDescuento, 2)), ""))
                  '---rRS!Valido = IIf(p_ing <> "" Or cdiaant = "S", "S", "")
                  '-Campo GRUPO: usado para guardar si fecha esta en periodo LACTANCIA
                  '---rRS!grupo = IIf(nLacta = 1, "LACT", "")
               End If
               rRS.Update
            End If
            
            If c2dias = "N" And cdiaant = "S" Then
               '-No turno 2dias y marca corresponde turno Dia Anterior
               cdiaant = "N"
            End If
         Next i             '===> sgt Dia
         nvalor = 0
         If nTtarda > 60 Then
            nvalor = nTDscto / 60
            If nTDscto Mod 60 > 30 Then
               nvalor = Int(nTDscto / 60) + 1
            Else
               nvalor = Int(nTDscto / 60)
            End If
         Else
            nTDscto = 0
         End If
         nthorasT = nthorasT - nvalor
         chdscto = cHorsMint(nvalor)
         
         Set rt = New Recordset
         sq_l = "SELECT * FROM CpTmGeneral WHERE tarjeta='" & cTarjeta & "'"
         rt.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If rt.EOF Then
            wdato = wdato + 1
            '-Generales del servidor
            rt.AddNew
            rt!tarjeta = r_d!tarjeta
            rt!apelpat = IIf(IsEmpty(r_d!apelpat), "", r_d!apelpat)
            rt!apelmat = IIf(IsEmpty(r_d!apelmat), "", r_d!apelmat)
            rt!nombres = IIf(IsEmpty(r_d!nombres), "", r_d!nombres)
            ctrabajador = Left(Trim(r_d!apelpat) + " " + Trim(r_d!apelmat) + " " + Trim(r_d!nombres) + IIf(nLacta = 0, "", IIf(nLacta = 1, " <LAC>", " <LAC2>")), 45)
         
            rt!trabajador = ctrabajador
            rt!PLH = IIf(IsEmpty(r_d!PLH), "", r_d!PLH)
            rt!tipopla = IIf(IsEmpty(r_d!tipopla), "", r_d!tipopla)
            rt!cCosto = IIf(IsEmpty(r_d!cCosto), "", r_d!cCosto)
            rt!CodUpr = IIf(IsEmpty(r_d!CodUpr), "", r_d!CodUpr)
            rt!grpgen = IIf(IsNull(r_d!grpgen), "", Trim(r_d!grpgen))
            rt!actividad = IIf(IsNull(r_d!actividad), "", Trim(r_d!actividad))
            rt!cargo = IIf(IsNull(r_d!cargo), "", Trim(r_d!cargo))
            cCodigo = r_d!cCosto
            If Trim(cCodigo) <> "" Then
                Set rs = New Recordset
                rs.Open "SELECT * FROM Cpccosto WHERE ( codigo='" & Trim(cCodigo) & "')", Cnx, adOpenStatic, adLockReadOnly
                If Not rs.EOF Then
                    rt!dscccosto = "" + rs!descripcion
                End If
                rs.Close
            End If
            cCodigo = r_d!CodUpr
            If IsEmpty(cCodigo) Then
            Else
                Set rs = New Recordset
                rs.Open "SELECT * FROM Cpundprgrol WHERE ( codupr='" & cCodigo & "')", Cnx, adOpenStatic, adLockOptimistic
                If Not rs.EOF Then
                    rt!dscUPR = "" + rs!descripcion
                End If
                rs.Close
            End If
            cCodigo = r_d!actividad
            If Trim(cCodigo) <> "" Then
                Set rs = New Recordset
                rs.Open "SELECT * FROM CpActividad WHERE  codigo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
                If Not rs.EOF Then
                    rt!dscactividad = "" + Left(rs!descripcion, 35)
                End If
                rs.Close
            End If
            cCodigo = r_d!cargo
            If Trim(cCodigo) <> "" Then
                Set rs = New Recordset
                rs.Open "SELECT * FROM Cpcargos WHERE  codigo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
                If Not rs.EOF Then
                    rt!DscCargo = "" + Trim(rs!descripcion)
                End If
                rs.Close
            End If
         End If
         If IsDate(r_d!fecnac) Then
            If Month(r_d!fecnac) = j Then
               '-Onomastico en el Mes
               rt!fecnac = r_d!fecnac
               rt!Onomastico = "Onomastico : " + Left(Format(r_d!fecnac, "DD/MM/YYYY"), 5)
            End If
         End If
         rt!hraprog = nThorasP
         rt!tardanza = nTtarda
         rt!tm130 = nTDscto
         rt!hraMp = nThorasMP
         rt!hraFJ = nThorasFj
         rt!hraFI = nThorasFi
         rt!hraFmp = ntHorasFMP
         rt!hraTrab = nthorasT
         rt!horasp = cHorsMint(nThorasP)
         rt!htarda = cHorsMint(nTtarda / 60)
         rt!hdscto = chdscto
         rt!horasmp = cHorsMint(nThorasMP)
         rt!horasfj = cHorsMint(nThorasFj)
         rt!horasfi = cHorsMint(nThorasFi)
         rt!horasfmp = cHorsMint(ntHorasFMP)
         rt!horast = cHorsMint(nthorasT)
         rt.Update
         rt.Close
      End If
      r_s.Close
     
      Next j         '====== sgt  DIRECTORIO del Servidor en periodo consultado
      
      r_d.MoveNext    'Sgt Servidor
   Loop
   rt1.Close
   rRS.Close
   r_d.Close
   Cnx.CommitTrans
      Beep 500, 350
   Progreso1.Visible = False
   If wdato > 0 Then
      Cmdprint.Enabled = True
      Lblrec.Caption = "Son " & Trim(CStr(wdato)) & " servidores.."
   End If
   If ccadena <> "" Then
      MsgBox "No programó servidores.." + ccadena, vbCritical
   End If
   
End Sub
Private Sub mnuFicImpSelec_Click()
    '- Seleccionar la impresora a usar                               (23/Ene/00)
    '- La detección de errores es por si no hay impresora instalada
    On Error Resume Next
    With CommonDialog1
        .DialogTitle = "Seleccionar impresora"
        .Flags = cdlPDPrintSetup
        .ShowPrinter
    End With
    Err = 0
End Sub
Private Sub CmdPrint_Click()
   If wdato = 0 Then Exit Sub
   ceditor = "X"
   If Wccosto = "" Then
      ceditor = ""
   Else
      'valida derechos en Upr
      npos = InStr(wCadenaUpr, TxCodUpr.Text)
      ceditor = Right(Mid(wCadenaUpr, npos, 3), 1)
   End If
   If ceditor = "X" Then
      MsgBox "Acceso, No autorizado...", vbExclamation
      Exit Sub
   End If
   
      nTipRpt = 5.5
     
   wTituRpt = ""
   InfoCpRS.Left = 0: InfoCpRS.Top = 0
   InfoCpRS.Show
   Exit Sub
End Sub

Private Sub CmdExportaXLS_Click()
    wTpoQuery = "7.0"
    Call ExportTableToXLS_click
    Beep 500, 350
    MsgBox "Se genero C:\RPT\SumRecord.xls", vbInformation
    Exit Sub
End Sub

Private Sub cmdsalir_Click()
  Set rs = Nothing
  Set rRS = Nothing
  Set rt = Nothing
  Set rt1 = Nothing
  Set DD = Nothing
  
  Me.Hide
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  Set rRS = Nothing
  Set rt = Nothing
  Set rt1 = Nothing
  Set DD = Nothing
  
  Me.Hide
  Unload Me
End Sub


