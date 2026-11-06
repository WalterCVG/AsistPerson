VERSION 5.00
Begin VB.Form TardaInasis 
   Caption         =   "Tardanzas  / Inasistencias  "
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   DrawWidth       =   4
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "TardaInasis.frx":0000
   ScaleHeight     =   4155
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ChkSUMARIO 
      Alignment       =   1  'Right Justify
      Caption         =   "Sumario Mensualizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox TxDscCargo 
      BorderStyle     =   0  'None
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   27
      Top             =   2160
      Width           =   5175
   End
   Begin VB.TextBox TxCargo 
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
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.CheckBox Chk3ero 
      Alignment       =   1  'Right Justify
      Caption         =   "Sin 3eros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7680
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Txdetalle1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Txccosto 
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
      Left            =   120
      MaxLength       =   2
      TabIndex        =   22
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox CboCCosto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   1200
      Width           =   5775
   End
   Begin VB.TextBox Txcodupr 
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
      Left            =   0
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox CboUpr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Top             =   1680
      Width           =   5775
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
      Left            =   5640
      Picture         =   "TardaInasis.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Imprimir"
      Top             =   3360
      Width           =   765
   End
   Begin VB.CommandButton cmdsalir 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8880
      Picture         =   "TardaInasis.frx":80B30
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cerrar"
      Top             =   3360
      Width           =   765
   End
   Begin VB.ComboBox CboTipoPla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Txtipopla 
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
      Height          =   225
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   375
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
      Left            =   5760
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
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
      Left            =   5040
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
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
      Left            =   7800
      TabIndex        =   4
      Top             =   240
      Width           =   975
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
      Left            =   6120
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tardanzas"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1225
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Inasistencia"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdcarga 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8520
      Picture         =   "TardaInasis.frx":80C7A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ubicar Registros"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton CmdClear 
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
      Left            =   6960
      Picture         =   "TardaInasis.frx":80E55
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Otra Consulta"
      Top             =   3360
      Width           =   735
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
      Picture         =   "TardaInasis.frx":80F6F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad Organica :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "U.P.R. (asistenciales)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Planilla"
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
      Left            =   1680
      TabIndex        =   20
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Al :"
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
      Left            =   7320
      TabIndex        =   16
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas     Del :"
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
      Left            =   5040
      TabIndex        =   15
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "TardaInasis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim r_t As New ADODB.Recordset
Dim Sw, f_cierra As Boolean
Dim nson, nsonC, nLacta, nSonLacta, nposicion, nlen, dfi, dff, z, nmintarde As Integer
Dim c_ano, c_riterio, cDni, csexo, cAplica, sql, t_urnoD, c_turno, c_Costo, cdscccosto, cCodUpr, cDscUpr, cCargo, cDscCargo, cvalor, c_opc, cDActividad As String
Dim unafecha As Date

Private Sub Command1_Click()
    c_grupo = "03": c_secuencia = "033"
    Load AydAcerca
    AydAcerca.Left = 0: AydAcerca.Top = 0
    AydAcerca.Show
End Sub
Private Sub Form_Load()
    TardaInasis.Caption = "Tardanzas / Inasistencias - " & genaño
    Chk3ero.Value = 1      'No Incluye 3eros
    Option1(0).Value = False: Option1(1).Value = False
    Txfecini = Date - 1
    TxFecFin = Date - 1
    c_ano = Trim(Str(genaño))
    TxTipoPla = ""
    'LLena el Combo TipoPLA
    CboTipoPla.Clear
    CboTipoPla.AddItem "ASISTENCIAL"
    CboTipoPla.AddItem "ADMINISTRATIVO"
    CboTipoPla.ListIndex = 0
    CboTipoPla.Text = ""
    Set rs = New Recordset
    rs.Open "SELECT * FROM CpCcosto ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
    rs.MoveFirst
    While Not rs.EOF
     ExisteYa = False
     For i = 0 To CboCcosto.ListCount - 1
         If CboCcosto.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("Codigo") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        CboCcosto.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("Codigo")
     End If
     rs.MoveNext
    Wend
    rs.Close
    
End Sub

Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
        c_opc = "T"
      Case 1
        c_opc = "I"
   End Select
   Txfecini.SetFocus
   'Txfecha.Text = primer + Index
End Sub


Private Sub txfecini_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If IsDate(Txfecini) Then
        TxFecFin = Txfecini
        TxFecFin.SetFocus
        Exit Sub
     Else
        MsgBox "verifique fecha ..", vbExclamation
     End If
  End If
End Sub
Private Sub txfecFin_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If IsDate(TxFecFin) Then
        If Year(DateValue(Txfecini)) <> Year(DateValue(TxFecFin)) Then
           MsgBox "ERROR: Verifique fechas ...", vbCritical, "Periodos Diferentes"
           TxFecFin = ""
           Exit Sub
        End If
        c_ano = Year(DateValue(TxFecFin))
        If Val(c_ano) <> Val(genaño) Then
           If MsgBox("Fechas no corresponden al año activo.. ", vbOKCancel + vbQuestion, "Confirme") = vbCancel Then
              Txfecini = "": TxFecFin = ""
              Exit Sub
           End If
        End If
        Exit Sub
     Else
        MsgBox "verifique fecha ..", vbExclamation
     End If
  End If
End Sub
Private Sub TxFecFin_LostFocus()
   Call txfecFin_KeyPress(13)
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
     TxTipoPla = Left(CboTipoPla.List(j), 3)
  End If
End Sub
Private Sub Cbotipopla_LostFocus()
   Call CboTipoPla_KeyPress(13)
End Sub
Private Sub Cboccosto_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboCcosto.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cboccosto_LostFocus()
  If IsEmpty(CboCcosto.Text) Then
  Else
     CbocCosto_KeyPress (13)
  End If
End Sub
Private Sub CbocCosto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboCcosto.Text) Then
         CboUPR.Clear
      Else
         j = CboCcosto.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboCcosto.List(j), ":")
            nlen = Len(CboCcosto.List(j))
            Txccosto = Right(CboCcosto.List(j), nlen - nposicion)
            Txdetalle1 = Left(CboCcosto.List(j), nposicion - 1)
            '-Carga Upr del CCosto
            
            CboUPR.Clear
            sql = "SELECT     tab1.CCOSTO, tab1.CODUPR, tab2.Descripcion " & _
            "FROM         CPCCostoUPR AS tab1 INNER JOIN " & _
            "                 CpUndPrgRol AS tab2 ON tab1.CODUPR = tab2.CodUPR " & _
            "WHERE CCOSTO='" & Txccosto + "'  ORDER BY tab2.descripcion "
     
            Set r_s = New Recordset
            r_s.Open sql, Cnx, adOpenStatic, adLockReadOnly
            r_s.MoveFirst
            While Not r_s.EOF
                ExisteYa = False
                For i = 0 To CboUPR.ListCount - 1
                    If CboUPR.List(i) = r_s.Fields("Descripcion") & "  :" & r_s.Fields("CodUPR") Then
                        ExisteYa = True
                        Exit For
                    End If
                Next i
                If Not ExisteYa Then
                    CboUPR.AddItem r_s.Fields("Descripcion") & "  :" & r_s.Fields("CodUPR")
                End If
                r_s.MoveNext
            Wend
            r_s.Close
         End If
      End If
   End If
End Sub
Private Sub CboUPR_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboUPR.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboUPR_LostFocus()
  If IsEmpty(CboUPR.Text) Then
  Else
     j = CboUPR.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboUPR.List(j), ":")
        nlen = Len(CboUPR.List(j))
        TxCodUpR = Right(CboUPR.List(j), nlen - nposicion)
        Txdetalle1 = Left(CboUPR.List(j), nposicion - 1)
     End If
     Exit Sub
  End If
End Sub
Private Sub CboUPR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call CboUPR_LostFocus
   End If
End Sub
Private Sub TxCargo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(TxCargo) = "" Then
         Load Ayuda3
         Ayuda3.Tag = 2
         Ayuda3.Show 1
      End If
      Set rs = New Recordset
      sql = "SELECT * FROM CpCargos WHERE codigo='" & Trim(TxCargo) & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      TxDscCargo = ""
      If Not rs.EOF Then
         TxDscCargo = Trim(rs!descripcion)
         If Trim(Txdetalle1) = "" Then
           Txdetalle1 = Trim(rs!descripcion)
        Else
           Txdetalle1 = Trim(Txdetalle1) + " - " + Trim(rs!descripcion)
        End If
      Else
         MsgBox "Cargo no existe...", vbExclamation, "[ ERROR ]"
         TxCargo = ""
      End If
      rs.Close
      Exit Sub
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
      KeyAscii = 0
   End If
End Sub
Private Sub cmdcarga_Click()
      
   If ChkSUMARIO.Value = 1 Then
      Call GenSumario
      Exit Sub
   End If
   
   If Option1(0).Value = False And Option1(1).Value = False Then
      MsgBox "Indicar Tardanza ó Inasistencias..", vbExclamation
      Exit Sub
   End If
   If IsDate(Txfecini) And IsDate(TxFecFin) Then
   Else
      MsgBox "Fecha no son validas..", vbExclamation
      Exit Sub
   End If
   If Month(Txfecini) <> Month(TxFecFin) Then
      MsgBox "No corresponden al mismo MES..", vbExclamation
      Exit Sub
   End If
   If DateValue(Txfecini) > Date Or DateValue(TxFecFin) > Date Then
      MsgBox "Verifique fechas..", vbExclamation
      Exit Sub
   End If
   
   If DateValue(Txfecini) = Date And Option1(1).Value Then
      MsgBox "Fechas indebida, para establecer faltas...", vbExclamation
      Exit Sub
   End If
   f_cierra = IIf(DateValue(Txfecini) = Date, False, True)
   
   ''''On Error Resume Next
   '-Limpia TEMPORAL
  
   Cnx.Execute "DELETE  FROM CpTmRecord"
   sql = "SELECT * FROM CpTmRecord "
   Set rt = New Recordset
   rt.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   dfi = Day(Txfecini)
   dff = Day(TxFecFin)
   cmes = Right("00" + Trim(Str(Month(Txfecini))), 2)
   unafecha = Txfecini
   Call ObtenerPrimerUltimoDia(unafecha)
   '-carga  "TURNO" Directorio
   c_riterio = ""
   If Trim(TxTipoPla) = "ASI" Then
      c_riterio = c_riterio & " and Flag='0' "
   ElseIf Trim(TxTipoPla) = "ADM" Then
          c_riterio = c_riterio & " and Flag='1' "
   End If
   If Chk3ero.Value = 1 Then
        c_riterio = c_riterio + "AND (PLH is NULL OR Plh<>'3RO') "
   End If
   If Trim(Txccosto) <> "" Then
      c_riterio = c_riterio + " AND ccosto='" & Txccosto & "' "
   End If
   If Trim(TxCodUpR) <> "" Then
      '-Upr solo ASISTENCIALES
      c_riterio = c_riterio + "AND codupr='" & TxCodUpR & "' "
   End If
   If Trim(TxCargo) <> "" Then
      c_riterio = c_riterio + "AND cargo='" & TxCargo & "' "
   End If
   '-*
   sql = "SELECT * FROM CpDirectorio  WHERE ano ='" & c_ano & "' AND mes='" & cmes & "' " & c_riterio + " ORDER BY tarjeta "
   
   Set rs = New Recordset
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   nsonC = 0: nson = 0
    Dim OpenForms
    v_an = 0
   Do While Not rs.EOF
   
      '-dar control a windows y permitir multihilo
      v_an = v_an + 1
      If v_an Mod 10 = 0 Then     'if loop has repeated
         OpenForms = DoEvents 'Yield to operating system
      End If
   
      cmdcarga.Caption = rs!tarjeta
      cTarjeta = Trim(rs!tarjeta)
      cDni = ""
      '-Contenido del Directorio
      c_Costo = "": cDscCosto = "": cCodUpr = "": cDscUpr = "": cCargo = "": cDscCargo = ""
      c_Costo = IIf(IsNull(rs!CCosto), "", rs!CCosto)
      cCodUpr = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)
      cCargo = IIf(IsNull(rs!Cargo), "", rs!Cargo)
      c_actividad = "": cDActividad = ""
      cGrpGen = "": cGrpPer = ""
      sql = "SELECT * FROM Cpgeneral WHERE  (baja is NULL) and tarjeta='" & cTarjeta & "'"
      Set rs_d = New Recordset
      rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs_d.EOF Then
         MsgBox "Servidor de baja " + cTarjeta, vbExclamation, "[IMPORTANTE]"
         rs_d.Close
      Else     '<=+=+=+=+
      
         cDni = IIf(IsNull(rs_d!documento), "", Trim(rs_d!documento))
         csexo = IIf(IsNull(rs_d!sexo), "", rs_d!sexo)
         cGrpGen = "" + Trim(rs_d!grpgen)
         cGrpPer = "" + Trim(rs_d!grpPer)
         c_actividad = IIf(IsNull(rs_d!actividad), "", rs_d!actividad)
         cCargo = IIf(IsNull(rs_d!Cargo), "", rs_d!Cargo)
         '-Corrige contenidos caso Vacios
         c_Costo = IIf(c_Costo = "", rs_d!CCosto, c_Costo)
         cCodUpr = IIf(cCodUpr = "", rs_d!CodUpr, cCodUpr)
         cCargo = IIf(cCargo = "", rs_d!Cargo, cCargo)
         rs_d.Close
      
     ''   If nsonC < 5 Then
     ''      MsgBox (" serv " + cTarjeta + "[" + cplh + "]")
     ''   End If
        
      Set rs_d = New Recordset
      rs_d.Open "SELECT * FROM CpActividad WHERE codigo='" & c_actividad & "'", Cnx, adOpenStatic, adLockReadOnly
      If Not rs_d.EOF Then
         cDActividad = Trim(rs_d!descripcion)
      End If
      rs_d.Close
        
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
          If cCargo <> "" Then
             Set rs_d = New Recordset
             rs_d.Open "SELECT * FROM Cpcargos WHERE codigo='" & cCargo & "'", Cnx, adOpenStatic, adLockReadOnly
             If Not rs_d.EOF Then
                cDscCargo = Trim(rs_d!descripcion)
             End If
             rs_d.Close
          End If
      
      '-Afecto X LACTANCIA
      '===================
      cAplica = "N": nSonLacta = 0
      TxIniLac = "": TxFinLac = ""
      
      'If cTarjeta = "11080" Then
      '   MsgBox " INICIO del " + Trim(Str(dfi)) + " al:" + Trim(Str(dff))
      'End If
            
      cdiaant = "N"      'controla marcacion caso Cierra turno NOCHE (2dias  19:30-07:30)
      
      ''  If nsonC < 5 Then
      ''     MsgBox (" Inc Dias " + cTarjeta)
      ''  End If
      
      For i = dfi To dff
         z = i - 1
         t_urnoD = "X"     'turno del directorio DMTMXNHCP
         c_turno = Trim(IIf(IsNull(rs.Fields(z + 4)), "", rs.Fields(z + 4)))
         If Trim(c_turno) = "*" Or Trim(c_turno) = "L" Or Trim(c_turno) = "V" Or Trim(c_turno) = "R" Then
         Else
            fechaB = Format(primer + z, "yyyymmdd")
            dfentra = primer + z
            dfsale = primer + z
            p0_ing = "": p0_sal = "": p0_ing1 = "": p0_sal1 = ""    'HORARIO ORIGINAL
            p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": c2dias = "N"
            m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
            nValorH = 0: nValorM = 0: ndscto = 0: nmintarde = 0: F_falta = False: nmotpersonal = 0: nDsctoMP = 0:  nComision = 0
            cvalor = "": cAuPrg = "": cDiahora = "": Ctipo = "": choraini = "": chorafin = "": cDescuenta = ""
            cMantGuardia = "N"
            
            '-Afecto X LACTANCIA
            '=============================
            cAplica = "N": nSonLacta = 0
            TxIniLac = "": TxFinLac = ""
            If csexo = "F" Then
                '-Ubica lactancia al periodo
                Set rs_d = New Recordset
                SQ1 = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "' order by fecnac"
                rs_d.Open SQ1, Cnx, adOpenStatic, adLockReadOnly, adCmdText
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
            '-Tolerancia x LACTANCIA (I)ngreso/(S)alida
            '=======================
            nLacta = 0
            If cAplica <> "N" Then
               nLacta = IIf(nSonLacta = 1, 1, 2)
            End If
            'If cTarjeta = "01979" Then
            '   MsgBox "Aplica:" + cAplica + " [" + Trim(Str(nLacta)) + "] Dia:" + Format(primer + z, "dd/mm/yyyy") + " en " + Format(TxIniLac, "dd/mm/yyyy") + " - " + Format(TxFinLac, "dd/mm/yyyy")
            'End If
            
            '-HORARIO prog en dia -**
            Sq0 = "SELECT * FROM CpHor" & c_ano & " WHERE Tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
            Set rs_d = New Recordset
            rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If rs_d.EOF Then
               If Trim(cGrpGen) <> "" Then
                  rs_d.Close
                  Sq0 = "SELECT * FROM CpHor" & c_ano & " WHERE grupo='" & cGrpGen & "' AND fecha ='" & fechaB & "'"
                  Set rs_d = New Recordset
                  rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
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
                  
                    'If cTarjeta = "11080" And i = 24 Then
                    '   MsgBox ("HOR => " + p0_ing + "-" + p0_sal + " ; " + p0_ing1 + "-" + p0_sal1)
                    'End If
                  
                  p_horas = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
                  nValorH = Val(Left(p_horas, 2))
                  nValorM = Val(Right(p_horas, 2))
                  If InStr("NHCP", t_urnoD) > 0 Then
                     c2dias = "S"
                  End If
               End If
            End If
            rs_d.Close
            
            '-Ubica AUTORIZACIONES: Motivos Personales, Comisiones Serv, vaca
            Sq0 = "SELECT * FROM CpAutori" & c_ano & " WHERE (anulado is NULL) and tarjeta='" & Trim(cTarjeta) & "' and FechaIni<='" + fechaB + "' and FechaFin>='" + fechaB + "'"     '   and turno=''"
            Set rs_d = New Recordset
            rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            Do While Not rs_d.EOF
               If Not rs_d.EOF Then
                  cAuPrg = Trim(rs_d!Codigo)
                  cDiahora = rs_d!diahora                              'D/H
                  Ctipo = IIf(IsNull(rs_d!tipo), "", rs_d!tipo)        'I/S/O   caso horas
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
                  If Ctipo = "I" Then
                     m_ing = choraini
                  ElseIf Ctipo = "S" Then
                         m_sal = chorafin
                  End If
               End If
               rs_d.MoveNext
            Loop
            rs_d.Close
            
            If Val(p_horas) > 0 Then     '-------> (A)
            
               nsonC = nsonC + 1   '---servidor  programado
               '-Marcas en Dia Programado
               Sq0 = "SELECT * FROM CpMarca" & c_ano & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
               '-Marcas del dia, solo si Justifica "D"
               If cDiahora <> "D" Then
                  If Val(p_sal1) > 0 Or Val(p_sal) < Val(p_ing) Then
                     '-Ubica marcas 2dias Actual-siguiente caso: turno t/n, n/d
                     fecha1 = Format(primer + z + 1, "yyyymmdd")
                     Sq0 = "SELECT * FROM CpMarca" & c_ano & " WHERE tarjeta='" & cTarjeta & "' AND fecha >='" & fechaB & "' AND fecha <='" & fecha1 & "' ORDER BY fecha,hora"
                     cdiaant = "S"
                     p_Sale = p_sal
                     If Val(p_sal1) > 0 Then
                        p_Sale = p_sal1         'Corrige fin turno
                     End If
                  End If
               End If
               
               Set rs_d = New Recordset
               rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If Not rs_d.EOF Then
                  '-ASIGNA marcas basado en Horario    I - S  /I1 - S1
                  d_fentra = rs_d!fecha    '========> 1era marca del servidor para turno programado
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
                        
                     ElseIf Trim(m_ing) <> "" And (Val(rs_d!hora) - Val(m_ing) < 60) And c2dias = "S" And i = Day(dfsale) Then
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
                            '07001900-19000700   0701 0738x 0740    1352 0712
                            m_sal = rs_d!hora
                     ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) > 0 And Val(m_sal1) = 0 And (Val(rs_d!hora) - Val(p_sal1) >= 0) Then
                            '-sal Turno partido      1300-1900/1900-0700
                            m_sal1 = rs_d!hora
                     '-Check carga marcaciones sin valor para turno
                      ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then     'And Val(m_ing1) > 0
                           m_sal1 = rs_d!hora
                     End If
                     rs_d.MoveNext
                  Loop
               End If
               rs_d.Close
          
            End If            '<----- (A)
            
              'If cTarjeta = "11080" And i = 24 Then
              '   MsgBox ("Marcas => " + m_ing + "-" + m_sal + "; " + m_ing1 + "-" + m_sal1)
              'End If
            
            If Val(p_horas) > 0 Then    '------> (B)
               '-Evalua Marcas  INGRESO
               
               If (t_urnoD <> "N" And t_urnoD <> "C" And Not f_cierra) Or f_cierra Then     '------> (C)
               If cDiahora = "D" And cDescuenta = "S" Then
                  F_falta = True
                  
               ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And Val(p_ing1) > 0 And p_sal <> p_ing1 Then  '======>
                      'Caso Horario:2000-0800   marca 0805-2010   O marca 0749-1404 1930-0820 ó  0803-2015/-0824
                      F_falta = True
               
               ElseIf (Val(m_ing) + Val(m_ing1) = 0) Then
                  '-No marco Ing,   valida vs autoriza
                  If cDiahora = "H" And Ctipo = "I" Then
                     cInicio = p_ing
                     cFin = choraini
                     If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                        F_falta = True
                        cDescuenta = "S"
                     End If
                  ElseIf cDiahora = "D" Then
                       If cDescuenta = "S" Then
                          F_falta = True
                       End If
                  Else
                       F_falta = True
                  End If
               ElseIf (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                  'Marco tarde  Autoriza H
                  If cDiahora = "H" And Ctipo = "I" Then
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
                               '-trabajo Antes de Turno
                               F_falta = True
                            Else
                               If c2dias = "S" And ((m_sal > p_sal And d_fentra = dfsale) Or (m_sal < p_sal And d_fentra < dfsale)) Then
                                  '-No completo el turno
                                  F_falta = True
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
                               End If
                            End If
                        End If
                  
               End If
               End If       '<------ (C)
               
               If f_cierra Then     '------> (D)
               '-Evalua MARCAS Salida
               If F_ALTA Then
               ElseIf (Val(m_sal) + Val(m_sal1) = 0) Then
                  'No marco  Sal   valida vs autoriza
                  If cDiahora = "H" And Ctipo = "S" Then
                     If (Val(p_sal) - Val(chorafin) > 0) Or (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                        F_falta = True
                        cDescuenta = "S"
                     End If
                  ElseIf cDiahora = "D" Then
                     If cDescuenta = "S" Then
                        F_falta = True
                     End If
                  Else
                     F_falta = True
                  End If
               ElseIf Val(p_sal1) > 0 And p0_sal <> p0_ing1 Then    'p_horas<>"2400"
                        If Val(m_sal1) = 0 Then
                            F_falta = True
                        Else
                            If Val(p_sal1) > Val(m_sal1) Then
                                 F_falta = True
                            End If
                        End If
               ElseIf Val(m_sal) > 0 And Val(p0_sal) > Val(m_sal) Then        'p_sal
                  'Marco antes de, Autoriza H
                  If cDiahora = "H" And Ctipo = "S" Then    '====> And cTipo = "S"  ** TO CHECK **
                     cInicio = p_sal
                     cFin = chorafin
                     If Val(p_sal) - Val(chorafin) > 0 Then
                        F_falta = True
                     End If
                  ElseIf cDiahora = "" Then
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
                        
                     ElseIf Val(p_sal1) - Val(cFin) > 0 Then
                       F_falta = True
                                               
                     End If
                  End If
               End If
               End If       '<-------- (D)
               
               If F_falta Then
               Else
                  If cDiahora = "H" And Ctipo = "I" Then
                  ElseIf cDiahora <> "D" Then      '<===No Autorizaciones
                     If (Val(m_ing) + Val(m_ing1) > 0) Then
                        '-Hay marcas, aplica tolerancia, calcula tardanza
                        If cDiahora = "H" And Ctipo = "I" Then
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
                        If Ctipo = "I" Then
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
               
            End If      '-------> (B)
            
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
            
            Sw = False
            '---Servidor Tardanza / Ausente
            If Option1(0).Value = True And Not F_falta And (ndscto > 6 Or nDsctoMP > 0) Then
               '-Tardanza
               Sw = True
               nson = nson + 1
            ElseIf Option1(1).Value = True And F_falta Then
                   '-Falta
                   nson = nson + 1
                   Sw = True
            End If
            
              'If cTarjeta = "11080" And F_falta Then
              '   MsgBox IIf(Sw, "F", "*")
              'End If
            
            If Sw Then
               rt.AddNew
               rt!tarjeta = cTarjeta
               rt!dscactividad = cDActividad
               rt!nombdia = Left(WeekdayName(Weekday(primer + z)), 3)
               rt!fecha = primer + z
               rt!turno = c_turno
               rt!Horario = cvalor
               '-- Turno Original
               rt!ping = "" + p0_ing
               rt!psal = "" + p0_sal
               rt!ping1 = "" + p0_ing1
               rt!psal1 = "" + p0_sal1
               rt!hrasprog = 0 + (nValorH + nValorM / 60)
               rt!ming = "" + m_ing
               rt!msal = "" + m_sal
               rt!ming1 = "" + m_ing1
               rt!msal1 = "" + m_sal1
               rt!tardanza = 0 + IIf(F_falta, 0, nmintarde)
               rt!DESCUENTO = 0 + IIf(F_falta, 0, ndscto)
               rt!autorizac = "" + cAuPrg
               rt!descuenta = "" + cDescuenta
               rt!diahora = "" + cDiahora
               rt!tipo = "" + Ctipo
               rt!horaini = "" + choraini
               rt!horafin = "" + chorafin
               rt!falta = IIf(F_falta, "S", "N")
               rt!Comision = nComision
               rt!motpersonal = nmotpersonal
               rt!DescuentoMP = 0 + IIf(F_falta, 0, IIf(p_ing <> "" Or cdiaant = "S", nDsctoMP, 0))
               rt!HrasTrab = IIf(F_falta, IIf(cDescuenta <> "N", 0, nvalor), IIf(p_ing <> "" Or cdiaant = "S", nvalor - nDsctoMP, 0))
               rt!hrasfalt = IIf(F_falta, IIf(cDescuenta <> "N", (nValorH + (nValorM / 60)), 0), 0)
               rt!Valido = IIf(p_ing <> "" Or cdiaant = "S", "S", "")
               
                  rt!CCosto = c_Costo
                  rt!dscccosto = Left(cDscCosto, 50)
                  rt!CodUpr = cCodUpr
                  rt!dscUPR = Left(cDscUpr, 55)
               rt.Update
               
            End If
            If c2dias = "N" And cdiaant = "S" Then
               '-No turno 2dias y marca corresponde turno Dia Anterior
               cdiaant = "N"
            End If
         End If
      Next
      
   ''     If nsonC < 5 Then
   ''         MsgBox (" fin Dias " + cTarjeta)
   ''     End If
        
      End If    '=+=+=+=>
      rs.MoveNext
   Loop
   rs.Close
   rt.Close
   cmdcarga.Caption = ""
   
   If nson > 0 Then
      MsgBox "Servidores Programados (ubicados): " + Format(nsonC, "###0") + IIf(Option1(0).Value, " Tardanzas : ", " Ausentes : ") + Format(nson, "###0")
      Call CmdPrint_Click
   Else
      MsgBox "No Registra Información..", vbExclamation
   End If
   CmdClear.SetFocus
End Sub

'====================   S U M A R I O   ====================

Private Sub GenSumario()
   If IsDate(Txfecini) And IsDate(TxFecFin) Then
   Else
      MsgBox "Fecha no son validas..", vbExclamation
      Exit Sub
   End If
   If Year(Txfecini) <> Year(TxFecFin) Then
      MsgBox "Periodos diferentes...", vbExclamation, "[NO PROCEDE]"
      Exit Sub
   End If
   If DateValue(Txfecini) >= Date Or DateValue(TxFecFin) >= Date Then
      MsgBox "Verifique fechas..", vbExclamation
      Exit Sub
   End If
   
   ''''On Error Resume Next
  
   Cnx.BeginTrans
   On Error GoTo CREA
   Cnx.Execute "DELETE from  #TablaST "
   GoTo sigue
CREA:
   Cnx.Execute ("SELECT * Into #TablaST  FROM CpTmpFaltTard WHERE 1 = 2")
sigue:
Cnx.CommitTrans
   Sqt = "SELECT * FROM #TablaST "
   Set rt = New Recordset
   rt.Open Sqt, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   rt.Close
  
   wflg = "NO"
   NombTabla = "TablaSD"
   Call CompruebaTabla
   If wflg = "SI" Then
      Cnx.Execute "DROP TABLE TablaSD "
   End If
   
 sqd = "  CREATE table TablaSD " & _
   "(tarjeta varchar(5) NULL, DNI varchar(10) NULL, Nombdia varchar(3) NULL, Fecha smalldatetime NULL, " & _
   "turno varchar(2) NULL, horario varchar(4) NULL, ping varchar(4) NULL, psal varchar(4) NULL, Ping1 varchar(4) NULL, psal1 varchar(4) NULL, " & _
   "hrasprog float NULL, ming varchar(4) NULL, msal varchar(4) NULL, ming1 varchar(4) NULL, msal1 varchar(4) NULL, tardanza float NULL, descuento float NULL," & _
   "Falta Varchar(1) NULL, ccosto varchar(2) NULL, dscCCosto varchar(50) NULL, codupr varchar(2) NULL,dscupr varchar(55) NULL )"
   ')  ON [PRIMARY]   TEXTIMAGE_ON [Primary]
   Cnx.Execute (sqd)
   
   sqd = "SELECT * FROM TablaSD "
   Set rt1 = New Recordset
   rt1.Open sqd, Cnx, adOpenStatic, adLockOptimistic, adCmdText
  
   fecha1 = Format(Txfecini, "yyyymmdd")
   msi = Month(Txfecini)
   diai = Day(Txfecini)
   fecha2 = Format(TxFecFin, "yyyymmdd")
   msf = Month(TxFecFin)
   diaf = Day(TxFecFin)
   cano = Trim(Str(Year(TxFecFin)))
   
   c_riterio = " BAJA IS NULL "
   If Trim(TxTipoPla) <> "" Then
      c_riterio = " and Tipopla='" + TxTipoPla + "' "
   End If
      
   If Chk3ero.Value = 1 Then
      c_riterio = c_riterio + "AND (PLH is NULL OR Plh<>'3RO') "
   End If
   If Trim(Txccosto) <> "" Then
      c_riterio = c_riterio + " AND ccosto='" & Txccosto & "' "
   End If
   If Trim(TxCodUpR) <> "" Then
      '-Upr solo ASISTENCIALES
      c_riterio = c_riterio + "AND codupr='" & TxCodUpR & "' "
   End If
   If Trim(TxCargo) <> "" Then
      c_riterio = c_riterio + "AND cargo='" & TxCargo & "' "
   End If
   
   Set r_d = New ADODB.Recordset
   sqG = "SELECT * FROM CpGeneral WHERE " & c_riterio & " ORDER BY ccosto,codUpr,apelpat,apelmat"
   r_d.Open sqG, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If r_d.EOF Then
      r_d.Close
      Exit Sub
   End If
   
   r_d.MoveFirst
   Dim OpenForms
   v_an = 0
   nson = 0
   Do While Not r_d.EOF         '=====> Servidores
      v_an = v_an + 1
      If v_an Mod 10 = 0 Then     'if loop has repeated
         OpenForms = DoEvents 'Yield to operating system
      End If
         
      cTarjeta = r_d!tarjeta
      cObjeto = r_d!tarjeta
      csexo = IIf(IsNull(r_d!sexo), "", Trim(r_d!sexo))
      cDni = IIf(IsNull(r_d!documento), "", Trim(r_d!documento))
      cnombres = Left(Trim(r_d!apelpat) + " " + Trim(r_d!apelmat) + " " + r_d!nombres, 45)
      cCargo = IIf(IsNull(r_d!Cargo), "", Trim(r_d!Cargo))
      ccCosto = IIf(IsNull(r_d!CCosto), "", Trim(r_d!CCosto))
      cCodUpr = IIf(IsNull(r_d!CodUpr), "", Trim(r_d!CodUpr))
      cDscCosto = "": cDscUpr = "": cDscCargo = ""
      Set rs = New Recordset
      rs.Open "SELECT * FROM Cpccosto WHERE codigo='" & ccCosto & "'", Cnx, adOpenStatic, adLockReadOnly
      If Not rs.EOF Then
         cDscCosto = Left(rs!descripcion, 50)
      End If
      rs.Close
      If cCodUpr <> "" Then
         Set rs = New Recordset
         rs.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs.EOF Then
            cDscUpr = Left(rs!descripcion, 55)
         End If
         rs.Close
      End If
      
      cCodigo = r_d!Cargo
      If Trim(cCodigo) <> "" Then
         Set rs = New Recordset
         rs.Open "SELECT * FROM Cpcargos WHERE  codigo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
         If Not rs.EOF Then
            cDscCargo = "" + Trim(rs!descripcion)
         End If
         rs.Close
      End If
      
      '- Totales acumulado  por servidor en el periodo
      nThorasP = 0:  nTtarda = 0: nTDscto = 0
      ntfaltas = 0: nttarde = 0:  ntHorast = 0
     
      '-Establece Afecto X LACTANCIA
      '=============================
      cAplica = "N": nSonLacta = 0
      TxIniLac = "": TxFinLac = ""
      
      For j = msi To msf
         '- Totales acumulado  por servidor en MES
         nfaltas = 0: ntarde = 0
         cms = Right("00" + Trim(Str(j)), 2)
         nThorasP = 0:  nTtarda = 0: nTDscto = 0
         ntHorast = 0
         cmdcarga.Caption = cTarjeta + ":" + cms
      ''''dff0 = Day(ultimo)  ultimo dia calendario en cada mes
        
         '-procesa DIRECTORIO del mes"
         sqd = "SELECT * FROM CpDirectorio  WHERE  tarjeta='" & cTarjeta & "' AND ano ='" & genaño & "' AND mes='" & cms & "'"
         Set r_s = New Recordset
         r_s.Open sqd, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not r_s.EOF Then
            ' If cTarjeta = "11080" Then
            '    MsgBox " INICIO del " + Trim(Str(dfi)) + " al:" + Trim(Str(dff))
            ' End If
            
            cdiaant = "N"      'controla marcacion caso Cierra turno NOCHE (2dias  19:30-07:30)
        
            '  If nsonC < 5 Then
            '     MsgBox (" Inc Dias " + cTarjeta)
            '  End If
            p_Sale = ""
            
            unafecha = "01/" + cms + "/" + cano
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
            For i = dfi To dff
               z = i - 1
               If i <= dff0 Then      '---Y  Marcas hasta Dia ultimo de mes dff0
                  t_urnoD = "X"     'turno del directorio DMTMXNHCP
                  c_turno = Trim(IIf(IsNull(r_s.Fields(z + 4)), "", r_s.Fields(z + 4)))
                  If Trim(c_turno) = "*" Or Trim(c_turno) = "L" Or Trim(c_turno) = "V" Or Trim(c_turno) = "R" Then
                  Else     '=.=.=>  (1)
                     fechaB = Format(primer + z, "yyyymmdd")
                     dfentra = primer + z
                     dfsale = primer + z
                     p0_ing = "": p0_sal = "": p0_ing1 = "": p0_sal1 = ""    'HORARIO ORIGINAL
                     p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = "": p_horas = "": c2dias = "N"
                     m_ing = "": m_ing1 = "": m_sal = "": m_sal1 = ""
                     nValorH = 0: nValorM = 0: ndscto = 0: nmintarde = 0: F_falta = False: nmotpersonal = 0: nDsctoMP = 0:  nComision = 0
                     cvalor = "": cAuPrg = "": cDiahora = "": Ctipo = "": choraini = "": chorafin = "": cDescuenta = ""
                     cMantGuardia = "N"
            
                     '-Afecto X LACTANCIA
                     '=============================
                     cAplica = "N": nSonLacta = 0
                     TxIniLac = "": TxFinLac = ""
                     If csexo = "F" Then
                        '-Ubica lactancia al periodo
                        Set rs = New Recordset
                        Sq0 = "SELECT * FROM CpLactancia WHERE DNI ='" & cDni & "' order by fecnac"
                        rs.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                        Do While Not rs.EOF
                           dftope = DateAdd("YYYY", 1, rs!fecnac)   'Tope Lactancia
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
                     'If cTarjeta = "01979" Then
                     '   MsgBox "Aplica:" + cAplica + " [" + Trim(Str(nLacta)) + "] Dia:" + Format(primer + z, "dd/mm/yyyy") + " en " + Format(TxIniLac, "dd/mm/yyyy") + " - " + Format(TxFinLac, "dd/mm/yyyy")
                     'End If
            
                     '-HORARIO prog en dia -**
                     Sq0 = "SELECT * FROM CpHor" & c_ano & " WHERE Tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "'"
                     Set rs_d = New Recordset
                     rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                     If rs_d.EOF Then
                        If Trim(cGrpGen) <> "" Then
                           rs_d.Close
                           Sq0 = "SELECT * FROM CpHor" & c_ano & " WHERE grupo='" & cGrpGen & "' AND fecha ='" & fechaB & "'"
                           Set rs_d = New Recordset
                           rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
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
                  
                           'If cTarjeta = "11080" And i = 24 Then
                           '   MsgBox ("HOR => " + p0_ing + "-" + p0_sal + " ; " + p0_ing1 + "-" + p0_sal1)
                           'End If
                  
                           p_horas = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
                           nValorH = Val(Left(p_horas, 2))
                           nValorM = Val(Right(p_horas, 2))
                           If InStr("NHCP", t_urnoD) > 0 Then
                              c2dias = "S"
                           End If
                        End If
                     End If
                     rs_d.Close
            
                     '-Ubica AUTORIZACIONES: Motivos Personales, Comisiones Serv, vaca
                     Sq0 = "SELECT * FROM CpAutori" & c_ano & " WHERE (anulado is NULL) and tarjeta='" & Trim(cTarjeta) & "' and FechaIni<='" + fechaB + "' and FechaFin>='" + fechaB + "'"     '   and turno=''"
                     Set rs_d = New Recordset
                     rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                     Do While Not rs_d.EOF
                        If Not rs_d.EOF Then
                            cAuPrg = Trim(rs_d!Codigo)
                            cDiahora = rs_d!diahora                              'D/H
                            Ctipo = IIf(IsNull(rs_d!tipo), "", rs_d!tipo)        'I/S/O   caso horas
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
                            If Ctipo = "I" Then
                                m_ing = choraini
                            ElseIf Ctipo = "S" Then
                                    m_sal = chorafin
                            End If
                        End If
                        rs_d.MoveNext
                     Loop
                     rs_d.Close
            
                     If Val(p_horas) > 0 Then     '-------> (A)
                        '-Marcas en Dia Programado
                        Sq0 = "SELECT * FROM CpMarca" & c_ano & " WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & fechaB & "' ORDER BY fecha,hora"
                        '-Marcas del dia, solo si Justifica "D"
                        If cDiahora <> "D" Then
                            If Val(p_sal1) > 0 Or Val(p_sal) < Val(p_ing) Then
                                '-Ubica marcas 2dias Actual-siguiente caso: turno t/n, n/d
                                fecha1 = Format(primer + z + 1, "yyyymmdd")
                                Sq0 = "SELECT * FROM CpMarca" & c_ano & " WHERE tarjeta='" & cTarjeta & "' AND fecha >='" & fechaB & "' AND fecha <='" & fecha1 & "' ORDER BY fecha,hora"
                                cdiaant = "S"
                                p_Sale = p_sal
                                If Val(p_sal1) > 0 Then
                                    p_Sale = p_sal1         'Corrige fin turno
                                End If
                            End If
                        End If
               
                        Set rs_d = New Recordset
                        rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                        If Not rs_d.EOF Then
                           '-ASIGNA marcas basado en Horario    I - S  /I1 - S1
                           d_fentra = rs_d!fecha    '========> 1era marca del servidor para turno programado
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
                            
                              ElseIf Trim(m_ing) <> "" And (Val(rs_d!hora) - Val(m_ing) < 60) And c2dias = "S" And i = Day(dfsale) Then
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
                                     '07001900-19000700   0701 0738x 0740    1352 0712
                                     m_sal = rs_d!hora
                              ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) > 0 And Val(m_sal1) = 0 And (Val(rs_d!hora) - Val(p_sal1) >= 0) Then
                                     '-sal Turno partido      1300-1900/1900-0700
                                     m_sal1 = rs_d!hora
                                    '-Check carga marcaciones sin valor para turno
                              ElseIf Day(dfentra + 1) = Day(dfsale) And Val(p0_sal1) = 0 And Val(m_sal1) = 0 And Val(m_ing) > 0 And Val(m_sal) > 0 Then     'And Val(m_ing1) > 0
                                     m_sal1 = rs_d!hora
                              End If
                              rs_d.MoveNext
                           Loop
                        End If
                        rs_d.Close
                     End If            '<----- (A)
            
                     'If cTarjeta = "11080" And i = 24 Then
                     '   MsgBox ("Marcas => " + m_ing + "-" + m_sal + "; " + m_ing1 + "-" + m_sal1)
                     'End If
            
                     If Val(p_horas) > 0 Then    '=====> (B)
                        '-Evalua Marcas  INGRESO
                        If cDiahora = "D" And cDescuenta = "S" Then       '------> (C)
                              F_falta = True
                  
                        ElseIf c2dias = "S" And (m_ing < m_sal) And (m_ing1 < m_sal1) And Val(p_ing1) > 0 And p_sal <> p_ing1 Then  '======>
                               'Caso Horario:2000-0800   marca 0805-2010   O marca 0749-1404 1930-0820 ó  0803-2015/-0824
                               F_falta = True
               
                        ElseIf (Val(m_ing) + Val(m_ing1) = 0) Then
                               '-No marco Ing,   valida vs autoriza
                               If cDiahora = "H" And Ctipo = "I" Then
                                  cInicio = p_ing
                                  cFin = choraini
                                  If CalcMints(cInicio, cFin) > 30 Or (Val(m_sal) > 0 And Val(p_sal) - Val(m_sal) > 0) Then
                                     F_falta = True
                                     cDescuenta = "S"
                                  End If
                               ElseIf cDiahora = "D" Then
                                      If cDescuenta = "S" Then
                                         F_falta = True
                                      End If
                               Else
                                  F_falta = True
                               End If
                        ElseIf (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                               'Marco tarde  Autoriza H al ingreso
                               If cDiahora = "H" And Ctipo = "I" Then
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
                                     '-trabajo Antes de Turno
                                     F_falta = True
                                  Else
                                     If c2dias = "S" And ((m_sal > p_sal And d_fentra = dfsale) Or (m_sal < p_sal And d_fentra < dfsale)) Then
                                        '-No completo el turno
                                        F_falta = True
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
                                    End If
                                 End If
                               End If
                  
                        End If      '<------ (C)
                        
               
                        '-Evalua MARCAS Salida
                        If F_ALTA Then       '------> (D)
                        ElseIf (Val(m_sal) + Val(m_sal1) = 0) Then
                               'No marco  Sal   valida vs autoriza
                               If cDiahora = "H" And Ctipo = "S" Then
                                  If (Val(p_sal) - Val(chorafin) > 0) Or (Val(m_ing) > 0 And CalcMTard(p_ing, m_ing) > 30) Then
                                     F_falta = True
                                     cDescuenta = "S"
                                  End If
                               ElseIf cDiahora = "D" Then
                                      If cDescuenta = "S" Then
                                         F_falta = True
                                      End If
                               Else
                                  F_falta = True
                               End If
                        ElseIf Val(p_sal1) > 0 And p0_sal <> p0_ing1 Then    'p_horas<>"2400"
                               If Val(m_sal1) = 0 Then
                                  F_falta = True
                               Else
                                  If Val(p_sal1) > Val(m_sal1) Then
                                     F_falta = True
                                  End If
                               End If
                        ElseIf Val(m_sal) > 0 And Val(p0_sal) > Val(m_sal) Then        'p_sal
                              'Marco antes de, Autoriza H
                              If cDiahora = "H" And Ctipo = "S" Then    '====> And cTipo = "S"  ** TO CHECK **
                                 cInicio = p_sal
                                 cFin = chorafin
                                 If Val(p_sal) - Val(chorafin) > 0 Then
                                    F_falta = True
                                 End If
                              ElseIf cDiahora = "" Then
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
                                     ElseIf Val(p_sal1) - Val(cFin) > 0 Then
                                            F_falta = True
                                     End If
                              End If
                        End If       '<-------- (D)
               
                        If F_falta Then
                        Else
                           If cDiahora = "H" And Ctipo = "I" Then
                           ElseIf cDiahora <> "D" Then      '<===No Autorizaciones
                                  If (Val(m_ing) + Val(m_ing1) > 0) Then
                                     '-Hay marcas, aplica tolerancia, calcula tardanza
                                     If cDiahora = "H" And Ctipo = "I" Then
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
                                     End If
                                  End If
                           End If
                        End If
               
                     End If      '<====== (B)
            
                     If F_falta Then
                        nvalor = nValorH + (nValorM / 60)
                     Else
                        nvalor = nValorH + ((nValorM - ndscto) / 60)
                     End If
                  End If   '<=.=.=  (1)
            
                  Sw = False
                  '---Servidor Tardanza / Ausente
                  If F_falta Then
                     '-falta
                     Sw = True
                     nfaltas = nfaltas + 1
                  ElseIf ndscto > 6 Then
                         '-tarde
                         Sw = True
                         ntarde = ntarde + 1
                  End If
            
                  'If cTarjeta = "11080" And F_falta Then
                  '   MsgBox IIf(Sw, "F", "*")
                  'End If
                  If Sw Then
                     'detalle dia falta/tarde
                     rt1.AddNew
                     rt1!tarjeta = cTarjeta
                     rt1!Dni = cDni
                     rt1!nombdia = Left(WeekdayName(Weekday(primer + z)), 3)
                     rt1!fecha = primer + z
                     rt1!turno = c_turno
                     rt1!Horario = cvalor
                     '-- Turno Original
                     rt1!ping = "" + p0_ing
                     rt1!psal = "" + p0_sal
                     rt1!ping1 = "" + p0_ing1
                     rt1!psal1 = "" + p0_sal1
                     rt1!hrasprog = 0 + (nValorH + nValorM / 60)
                     rt1!ming = "" + m_ing
                     rt1!msal = "" + m_sal
                     rt1!ming1 = "" + m_ing1
                     rt1!msal1 = "" + m_sal1
                     rt1!tardanza = 0 + IIf(F_falta, 0, nmintarde)
                     rt1!DESCUENTO = 0 + IIf(F_falta, 0, ndscto)
                     rt1!falta = IIf(F_falta, "S", "")
                     rt1!CCosto = c_Costo
                     rt1!dscccosto = Left(cDscCosto, 50)
                     rt1!CodUpr = cCodUpr
                     rt1!dscUPR = Left(cDscUpr, 55)
                     rt1.Update
                  End If
                  If c2dias = "N" And cdiaant = "S" Then
                     '-No turno 2dias y marca corresponde turno Dia Anterior
                     cdiaant = "N"
                  End If
               End If
            Next i      'cambio dia
      
            '     If nsonC < 5 Then
            '         MsgBox (" fin Dias " + cTarjeta)
            '     End If
         End If
         If nfaltas + ntarde > 0 Then
            Sqt = "SELECT * FROM #TablaST where DNI='" + cDni + "'"
            Set rt = New Recordset
            rt.Open Sqt, Cnx, adOpenStatic, adLockOptimistic, adCmdText
            If rt.EOF Then
               rt.AddNew
               rt!Dni = cDni
               rt!servidor = cnombres
               rt!Cargo = cDscCargo
               nson = nson + 1
            End If
            ntfaltas = ntfaltas + nfaltas
            nttarde = nttarde + ntarde
            rt.Fields(3 + (j - 1) * 2) = nfaltas
            rt.Fields(3 + (j - 1) * 2 + 1) = ntarde
            rt.Update
            rt.Close
         End If
         
         r_s.Close
      Next j    'cambio Mes  graba totales
      If ntfaltas + nttarde > 0 Then
         Sqt = "SELECT * FROM #TablaST where DNI='" + cDni + "'"
         Set rt = New Recordset
         rt.Open Sqt, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If Not rt.EOF Then
            rt!Totfaltas = ntfaltas
            rt!TotTarda = nttarde
            rt.Update
         End If
         rt.Close
      End If
      r_d.MoveNext    'otro servidor
   Loop
   r_d.Close
   rt1.Close
   cmdcarga.Caption = ""
   
   If nson > 0 Then
      wTpoQuery = "4.3"
      wNameXLS = "C:\RPT\SUMARIOTot.xls"
      Call Exportar_ADO_Excel
      
      wTpoQuery = "5.2"
      wNameXLS = "C:\RPT\SUMARIODet.xls"
      Call Exportar_ADO_Excel
      
      
      Beep 500, 350
      MsgBox "Generó C:\RPT\SUMARIOTot.xls", vbInformation, "[IMPORTANTE]"
      ''''Call CmdPrint_Click

      sqd = "drop table TablaSD "
      Cnx.Execute (sqd)

   Else
      MsgBox "No Registra Información..", vbExclamation
   End If
   CmdClear.SetFocus
End Sub

Private Sub CmdPrint_Click()
   If nson = 0 Then
      MsgBox "No registra resultados...", vbExclamation
      Exit Sub
   End If
   If ChkSUMARIO.Value = 1 Then
      nTipRpt = 4.3
      InfoCpRS.Left = 0: InfoCpRS.Top = 0
      InfoCpRS.Refresh
      InfoCpRS.Show
      Exit Sub
      
   Else
       If Option1(0).Value = False And Option1(1).Value = False Then
            MsgBox "Seleccione tipo de reporte", vbExclamation
            Exit Sub
       End If
       If Option1(0).Value = True Then
          '-Tardanza
          nTipRpt = 8
       ElseIf Option1(1).Value = True Then
              '-Inasistencia
              nTipRpt = 9
       End If
   End If
   InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
   InfoCpSQL.Refresh
   InfoCpSQL.Show
   Exit Sub
End Sub

Private Sub CmdClear_Click()
  Chk3ero.Value = 1      'No Incluye 3eros
  Option1(0).Value = False: Option1(1).Value = False
  Txfecini = Date - 1: TxFecFin = Date - 1: TxIniLac = "": TxFinLac = ""
  CboTipoPla.Text = "": TxTipoPla = ""
  CboCcosto.Text = "": Txccosto = "": Txdetalle1 = ""
  CboUPR.Clear: TxCodUpR = "": TxCargo = "": TxDscCargo = ""
  c_turno = "": ChkSUMARIO.Value = 1
  
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


