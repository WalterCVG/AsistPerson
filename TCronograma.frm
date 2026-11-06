VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TCronograma 
   Caption         =   "Cronograma"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "TCronograma.frx":0000
   ScaleHeight     =   5835
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   0
      Picture         =   "TCronograma.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   0
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   7920
      TabIndex        =   11
      Top             =   2760
      Width           =   1335
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   7215
      Begin VB.TextBox TxCierreMes 
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
         Left            =   6000
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   131006465
         CurrentDate     =   41609
         MinDate         =   41609
      End
      Begin VB.ComboBox CboMes 
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Txmes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         MaxLength       =   2
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Txano 
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CIERRE DE MES (cambios/Excepciones)"
         Height          =   375
         Left            =   4200
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "MES :"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha limite Programación UPR"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "AÑO :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   8280
      Picture         =   "TCronograma.frx":80EA8
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cerrar"
      Top             =   5040
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5106
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
      Caption         =   "Contenida de Tabla"
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
End
Attribute VB_Name = "TCronograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean
Dim unafecha As Date
Private Sub Command1_Click()
     c_grupo = "01": c_secuencia = "01A"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub

Private Sub Form_Load()
    Set DD = New Recordset
    With DD
      .Fields.Append "ano", adChar, 4
      .Fields.Append "mes", adChar, 2
      .Fields.Append "fecha", adDate
      .Fields.Append "cierreMes", adDate
      .Open
    End With
    Set DataGrid1.DataSource = DD
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         '.Caption = "ano"
          .Width = Me.TextWidth(String(6, "S"))
         Case 1
          '.Caption = "mes"
          .Width = Me.TextWidth(String(4, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM CpCronograma  order by ano,mes"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!ano = rs!ano
       DD!mes = "" + rs!mes
       If IsDate(rs!fecha) Then DD!fecha = rs!fecha
       If IsDate(rs!cierremes) Then DD!cierremes = rs!cierremes
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
    
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
    desabilitatextos
    Limpiacampos
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Txano = DataGrid1.Columns(0)
   Txmes = DataGrid1.Columns(1)
   If IsDate(DataGrid1.Columns(2)) Then DTPicker1.Value = DataGrid1.Columns(2)
   If IsDate(DataGrid1.Columns(3)) Then TxCierreMes.Text = DataGrid1.Columns(3)
   If Trim(Txmes) <> "" Then
       For i = 0 To CboMes.ListCount - 1
         npos = InStr(CboMes.List(i), ":")
         nlen = Len(CboMes.List(i))
         dato = Trim(Right(CboMes.List(i), nlen - npos))
         If dato = Trim(Txmes) Then
            CboMes.ListIndex = i
            Exit For
         End If
       Next i
   End If
   CmdEditar.Enabled = True
   CmdEliminar.Enabled = True
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
  With DD
    If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
       .Sort = .Fields(ColIndex).[Name] & " Desc"
    Else
       .Sort = .Fields(ColIndex).[Name] & " Asc"
    End If
  End With
  Call Limpiacampos
End Sub
Private Sub cmdNuevo_Click()
    If wrights = "CON" Then
       MsgBox "Usuario No Autorizado", vbExclamation
       Exit Sub
    End If
  If Cmdnuevo.Caption = "&Nuevo" Then
    Limpiacampos
    habilitatextos
    Txano.Locked = False
    CboMes.Enabled = True
    Txano.SetFocus
    Txano = genaño
    DTPicker1.Value = Date
    CmdEliminar.Enabled = False
    Sw = True
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
  Else
    'Grabar
    If DTPicker1.Value > DateValue(TxCierreMes) Or Txmes = "" Then
       MsgBox "Verifique Periodo..", vbExclamation
       Exit Sub
    End If
    On Error GoTo err_trans
    Set rs = New Recordset
    sql = "SELECT * FROM CpCronograma WHERE ano='" & Txano & "' and mes='" & Txmes & "'"
    rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    If Sw Then
       If rs.EOF Then
          rs.AddNew
          rs!ano = Txano
          rs!mes = Txmes
       End If
    End If
    rs!fecha = DTPicker1.Value
    rs!cierremes = TxCierreMes
    rs!usuario = wUsuario
    rs!fecRegistro = Date
    rs.Update
    rs.Close
    If Sw Then
       DD.AddNew
       DD!ano = Txano
       DD!mes = Txmes
    End If
    DD!fecha = DTPicker1.Value
    DD!cierremes = TxCierreMes
    DD.Update
    desabilitatextos
    Limpiacampos
    CmdEliminar.Enabled = True
    Cmdnuevo.Caption = "&Nuevo"
    CmdEditar.Caption = "&Editar"
    DataGrid1.Refresh
    CmdEditar.Enabled = False
    CmdEliminar.Enabled = False
    DataGrid1.SetFocus
    Exit Sub
err_trans:
    MsgBox Err.Description
  End If
End Sub
Private Sub Cmdeditar_Click()
   If wrights = "CON" Then
       MsgBox "Usuario No Autorizado", vbExclamation
       Exit Sub
   End If
   If CmdEditar.Caption = "&Editar" Then
      habilitatextos
      DTPicker1.SetFocus
      CmdEliminar.Enabled = False
      Sw = False
      Cmdnuevo.Caption = "&Grabar"
      CmdEditar.Caption = "&Cancelar"
   Else                   'Cancelar
      desabilitatextos
      Limpiacampos
      CmdEliminar.Enabled = True
      Cmdnuevo.Caption = "&Nuevo"
      CmdEditar.Caption = "&Editar"
      CmdEditar.Enabled = False
      CmdEliminar.Enabled = False
      DataGrid1.SetFocus
  End If
End Sub
Private Sub CmdEliminar_Click()
    If wrights = "CON" Then
       MsgBox "Usuario No Autorizado", vbExclamation
       Exit Sub
    End If
    Set rs = New Recordset
    sql = "SELECT * FROM CpDirectorio WHERE ano='" & Trim(Txano) & "' and mes='" & Txmes & "'"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then
    Else
       MsgBox "Hay Directorio con dicho Periodo", vbExclamation
       rs.Close
       Exit Sub
    End If
    rs.Close
    If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
      Cnx.Execute "DELETE FROM CpCronograma WHERE ano='" & Txano & "' and mes='" & Txmes & "'"
      DD.Delete
      'DataGrid1.Refresh
      Limpiacampos
    End If
End Sub
Private Sub cmdsalir_Click()
   Set r_d = Nothing
   Set DD = Nothing
   Set rs = Nothing
   Unload Me
   Me.Hide
End Sub
Sub habilitatextos()
   DTPicker1.Enabled = True
   TxCierreMes.Enabled = True
End Sub
Sub desabilitatextos()
   Txano.Locked = True
   Txmes.Locked = True
   CboMes.Enabled = False
   DTPicker1.Enabled = False
   TxCierreMes.Enabled = False
End Sub
Sub Limpiacampos()
   Txano = "":  Txmes = "": CboMes.Text = "": TxCierreMes = "": DTPicker1.Value = Date
End Sub




Private Sub Txano_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Val(Txano) >= 2013 Then
      CboMes.SetFocus
      Exit Sub
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
      KeyAscii = 0
   End If
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
            
            unafecha = "01/" + Txmes + "/" + Trim(Txano)
            Call ObtenerPrimerUltimoDia(unafecha)
          
         End If
      End If
      Exit Sub
   End If
End Sub
Private Sub DTPicker1_Change()
  If DTPicker1.Value > ultimo Then
     If MsgBox("Edita cronograma de mes concluido ?  ", vbQuestion + vbYesNo + vbDefaultButton2, "[IMPORTANTE ]") = vbYes Then
        Exit Sub
     End If
     desabilitatextos
     Limpiacampos
     
     CmdEliminar.Enabled = True
     Cmdnuevo.Caption = "&Nuevo"
     CmdEditar.Caption = "&Editar"
     CmdEditar.Enabled = False
     CmdEliminar.Enabled = False
     DataGrid1.SetFocus
     Sw = False
     
  End If
End Sub

Private Sub TxCierreMes_KeyPress(KeyAscii As Integer)
   If KeyPress = 13 Then
      If DTPicker1.Value > DateValue(TxCierreMes) Then
         MsgBox " Verifique fecha cierre de mes...", vkcritical, "[IMPORTANTE]"
         Exit Sub
         TxCierreMes = ""
      End If
   End If
End Sub
