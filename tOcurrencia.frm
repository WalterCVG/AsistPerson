VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tOcurrencia 
   Caption         =   "Tabla Ocurrencias"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "tOcurrencia.frx":0000
   ScaleHeight     =   7305
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txbuscar 
      Height          =   285
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   17
      Top             =   3960
      Width           =   3975
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   9240
      Picture         =   "tOcurrencia.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cerrar"
      Top             =   6480
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   8520
      TabIndex        =   11
      Top             =   4200
      Width           =   1575
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   840
      TabIndex        =   1
      Top             =   4440
      Width           =   7335
      Begin VB.TextBox TxturnoDirectorio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         MaxLength       =   2
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CboTurno 
         Height          =   315
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Mantiene Directorio"
         Height          =   375
         Left            =   5040
         TabIndex        =   19
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Mantiene Guardia"
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxDescuento 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox TxAbrev 
         Height          =   285
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Txdescripcion 
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   7
         Top             =   840
         Width           =   5295
      End
      Begin VB.TextBox Txcodigo 
         Height          =   285
         Left            =   720
         MaxLength       =   4
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Turno en Directorio"
         Height          =   375
         Left            =   4200
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Descuento (falta)"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Abreviatura"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5953
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
      Caption         =   "Tabla de contenido"
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Ubica Descripcion"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   3960
      Width           =   1335
   End
End
Attribute VB_Name = "tOcurrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean
Private Sub Form_Load()
    CboTurno.Clear
    CboTurno.AddItem "Vacaciones:V"
    CboTurno.AddItem "Licencia  :L"
    CboTurno.AddItem "Descanso  :R"
    CboTurno.ListIndex = 0
    CboTurno.Text = ""
    Set DD = New Recordset
    With DD
     .Fields.Append "codigo", adVarChar, 4
     .Fields.Append "Abrev", adVarChar, 10
     .Fields.Append "descripcion", adVarChar, 20
     .Fields.Append "descuento", adVarChar, 1
     .Fields.Append "MantGuardia", adVarChar, 1
     .Fields.Append "MantDirectorio", adVarChar, 1
     .Fields.Append "TurnoDirectorio", adVarChar, 2
     .Open
    End With
    Set DataGrid1.DataSource = DD
    'Configura columnas
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         '.Caption = "Codigo"
          .Width = Me.TextWidth(String(10, "S"))
         Case 1
          '.Caption = "Abrev"
          .Width = Me.TextWidth(String(15, "S"))
         Case 2
          '.Caption = "Descripcion"
          .Width = Me.TextWidth(String(30, "S"))
         Case Else
          .Visible = False
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM CpOcurrencias "
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!codigo = rs!codigo
       DD!abrev = IIf(IsNull(rs!abrev), "", rs!abrev)
       DD!descripcion = rs!descripcion
       DD!DESCUENTO = IIf(IsNull(rs!DESCUENTO), "", rs!DESCUENTO)
       DD!MantGuardia = IIf(IsNull(rs!MantGuardia), "", rs!MantGuardia)
       DD!MantDirectorio = IIf(IsNull(rs!MantDirectorio), "", rs!MantDirectorio)
       DD!turnodirectorio = IIf(IsNull(rs!turnodirectorio), "", Trim(rs!turnodirectorio))
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Call Limpiacampos
   Txcodigo = DataGrid1.Columns(0)
   TxAbrev = DataGrid1.Columns(1)
   Txdescripcion = DataGrid1.Columns(2)
   TxDescuento = DataGrid1.Columns(3)
   Check1.Value = IIf(DataGrid1.Columns(4) = "S", 1, 0)
   Check2.Value = IIf(DataGrid1.Columns(5) = "S", 1, 0)
   TxturnoDirectorio = DataGrid1.Columns(6)
   If Trim(TxturnoDirectorio) <> "" Then
       For i = 0 To CboTurno.ListCount - 1
         npos = InStr(CboTurno.List(i), ":")
         nlen = Len(CboTurno.List(i))
         dato = Trim(Right(CboTurno.List(i), nlen - npos))
         If dato = Trim(TxturnoDirectorio) Then
            CboTurno.ListIndex = i
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
Private Sub Check1_Click()
  If Check1.Value = 1 Then
  Else
  End If
End Sub
Private Sub Check12Click()
  If Check2.Value = 1 Then
  Else
  End If
End Sub
Private Sub cmdNuevo_Click()
    If wrights = "CON" Then
       MsgBox "Usuario No Autorizado", vbExclamation
       Exit Sub
    End If
  If Cmdnuevo.Caption = "&Nuevo" Then
    Limpiacampos
    habilitatextos
    Txcodigo.SetFocus
    CmdEliminar.Enabled = False
    Sw = True
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
  Else                  'Grabar
    On Error GoTo err_trans
    Cnx.BeginTrans
    Dim sql As String
    Set rs = New Recordset
    sql = "SELECT * FROM CpOcurrencias WHERE codigo='" + Trim(Txcodigo) + "'"
    rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF Then
       rs.AddNew
       rs!codigo = Txcodigo
    End If
    rs!abrev = UCase(TxAbrev)
    rs!descripcion = Txdescripcion
    rs!DESCUENTO = TxDescuento
    rs!MantGuardia = IIf(Check1.Value = 1, "S", "")
    rs!MantDirectorio = IIf(Check2.Value = 1, "S", "")
    rs!turnodirectorio = UCase(TxturnoDirectorio)
    rs!usuario = wUsuario
    rs!fecha = Date
    rs.Update
    Cnx.CommitTrans
    rs.Close
    If Sw Then
       DD.AddNew
       DD!codigo = Txcodigo.Text
    End If
    DD!abrev = IIf(IsNull(TxAbrev), "", TxAbrev)
    DD!descripcion = UCase(Txdescripcion)
    DD!DESCUENTO = IIf(IsNull(TxDescuento), "", TxDescuento)
    DD!MantGuardia = IIf(Check1.Value = 1, "S", "")
    DD!MantDirectorio = IIf(Check2.Value = 1, "S", "")
    DD!turnodirectorio = UCase(TxturnoDirectorio)
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
    Txcodigo.Locked = True
    TxAbrev.SetFocus
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
   Set r_d = New ADODB.Recordset
   sql = "SELECT * FROM CpAutori" & genaño & "  WHERE (anulado is NULL) and (codigo='" & Txcodigo.Text & "')"
   r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If r_d.EOF Then
   Else
      MsgBox "Hay Autorizaciones con dicho Codigo", vbExclamation
      r_d.Close
      Exit Sub
   End If
   r_d.Close
   If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
      Cnx.Execute "DELETE FROM CpOcurrencias WHERE codigo='" & Txcodigo.Text & "'"
      DD.Delete
      'DataGrid1.Refresh
      Limpiacampos
   End If
End Sub
Private Sub cmdsalir_Click()
   Unload Me
   Me.Hide
End Sub
Sub habilitatextos()
   Txcodigo.Locked = False
   TxAbrev.Locked = False
   Txdescripcion.Locked = False
   TxDescuento.Locked = False
   Check1.Enabled = True: Check2.Enabled = True
   CboTurno.Enabled = True
End Sub
Sub desabilitatextos()
   Txcodigo.Locked = True
   TxAbrev.Locked = True
   Txdescripcion.Locked = True
   TxDescuento.Locked = True
   Check1.Enabled = False: Check2.Enabled = True
   CboTurno.Enabled = False
End Sub
Sub Limpiacampos()
   Txcodigo = "": TxAbrev = "":  Txdescripcion = "":     TxDescuento = ""
   Check1.Value = 0: Check2.Value = 0
   CboTurno.Text = "": TxturnoDirectorio = ""
End Sub

Private Sub TxBuscar_Change()
    Dim cBuscar As String
    cpuntero = DataGrid1.Bookmark
    If Txbuscar = "" Then Exit Sub
    cBuscar = "descripcion like '*" & Txbuscar & "*'"
    DD.MoveFirst
    DD.Find (cBuscar)
    If DD.EOF Then
       Txbuscar.BackColor = &HFF&
    Else
       Txbuscar.BackColor = &H80000005
       DataGrid1.Refresh
    End If
End Sub

Private Sub Txcodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Txcodigo = UCase(Txcodigo)
      TxAbrev.SetFocus
      Exit Sub
   End If
End Sub
Private Sub Txabrev_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxAbrev = UCase(TxAbrev)
      TxDescuento.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Cboturno_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboTurno.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cboturno_LostFocus()
   If IsEmpty(CboTurno.Text) Then
   Else
      j = CboTurno.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboTurno.List(j), ":")
         nlen = Len(CboTurno.List(j))
         TxturnoDirectorio = Right(CboTurno.List(j), nlen - nposicion)
      End If
      Exit Sub
   End If
End Sub
Private Sub Cboturno_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsEmpty(CboTurno.Text) Then
      Else
         j = CboTurno.ListIndex
         If j <> -1 Then
            nposicion = InStr(CboTurno.List(j), ":")
            nlen = Len(CboTurno.List(j))
            TxturnoDirectorio = Right(CboTurno.List(j), nlen - nposicion)
         End If
      End If
      Exit Sub
   End If
End Sub


Private Sub Txdescuento_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxDescuento = UCase(TxDescuento)
      Check1.SetFocus
      Exit Sub
   End If
End Sub
