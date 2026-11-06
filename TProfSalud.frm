VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tProfSalud 
   Caption         =   "Tablas Profesionales de Salud (SuSalud)"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "TProfSalud.frx":0000
   ScaleHeight     =   8610
   ScaleMode       =   0  'User
   ScaleWidth      =   10358.06
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   8880
      Picture         =   "TProfSalud.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Cerrar"
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox Txbuscar 
      Height          =   285
      Left            =   5280
      MaxLength       =   35
      TabIndex        =   19
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   840
      TabIndex        =   12
      Top             =   5040
      Width           =   6735
      Begin VB.TextBox Txtipdoc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxCodigo 
         Height          =   285
         Left            =   3840
         MaxLength       =   6
         ScrollBars      =   1  'Horizontal
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxObservacion 
         Height          =   285
         Left            =   2400
         MaxLength       =   40
         TabIndex        =   7
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox Txtelefono 
         Height          =   285
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   6
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox CboTipDoc 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox TxNumero 
         Height          =   285
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TxTipo 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Txdescripcion 
         Height          =   285
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label8 
         Caption         =   "CÓDIGO :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Observación :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Numero:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "D.O.I."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Apellidos y Nombres"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   8400
      TabIndex        =   8
      Top             =   5160
      Width           =   1575
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   6588
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
      Caption         =   "Tablas de contenido"
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
      Caption         =   "Ubica Apellidios y Nombres"
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   4200
      Width           =   2295
   End
End
Attribute VB_Name = "tProfSalud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean
Private Sub Form_Load()
    Set rs = New Recordset      'carga Data
    sql = "SELECT * from TablasAux   where grupo='DOC' ORDER BY descripcion"
    rs.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
    While Not rs.EOF
         ExisteYa = False
         For i = 0 To CboTipDoc.ListCount - 1
              If CboTipDoc.List(i) = rs.Fields("descripcion") & " :" & rs.Fields("codigo") Then
                 ExisteYa = True
                 Exit For
              End If
         Next i
         If Not ExisteYa Then
             CboTipDoc.AddItem rs.Fields("descripcion") & " :" & rs.Fields("codigo")
         End If
         rs.MoveNext
    Wend
    rs.Close

    Set DD = New Recordset
    With DD
     .Fields.Append "Tipo", adVarChar, 6
     .Fields.Append "Codigo", adVarChar, 6
     .Fields.Append "descripcion", adVarChar, 60
     .Fields.Append "TipDoc", adVarChar, 2
     .Fields.Append "Numero", adVarChar, 10
     .Fields.Append "Telefono", adVarChar, 9
     .Fields.Append "Observacion", adVarChar, 40
     .Open
    End With
    Set DataGrid1.DataSource = DD
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         .Caption = "Grupo"
          .Width = Me.TextWidth(String(7, "S"))
         Case 1
         .Caption = "Codigo"
          .Width = Me.TextWidth(String(6, "S"))
         Case 2
          .Caption = "Descripcion"
          .Width = Me.TextWidth(String(40, "S"))
         Case 3
          .Caption = "Tip.Doc"
          .Width = Me.TextWidth(String(6, "S"))
         Case 4
          .Caption = "Numero"
          .Width = Me.TextWidth(String(10, "S"))
         Case Else
          .Visible = False
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM ProfSalud "
    rs.Open sql, Cnx0, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!tipo = rs!tipo
       DD!Codigo = rs!Codigo
       DD!descripcion = Trim(rs!descripcion)
       DD!tipdoc = IIf(IsNull(rs!tipdoc), "", Trim(rs!tipdoc))
       DD!numero = IIf(IsNull(rs!numero), "", Trim(rs!numero))
       DD!telefono = IIf(IsNull(rs!telefono), "", Trim(rs!telefono))
       DD!observacion = IIf(IsNull(rs!observacion), "", Trim(rs!observacion))
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
    desabilitatextos
End Sub
Private Sub DataGrid1_Click()
   Txbuscar = ""
   Txbuscar.BackColor = &H80000005
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   TxTipo = DataGrid1.Columns(0)
   TxCodigo = DataGrid1.Columns(1)
   Txdescripcion = DataGrid1.Columns(2)
   Txtipdoc = DataGrid1.Columns(3)
      If Trim(Txtipdoc) <> "" Then
      'Ubicar CODIGO  en combo1
      For i = 0 To CboTipDoc.ListCount - 1
         npos = InStr(CboTipDoc.List(i), ":")
         nlen = Len(CboTipDoc.List(i))
         dato = Right(CboTipDoc.List(i), nlen - npos)
         If dato = Trim(Txtipdoc) Then
            Exit For
         End If
      Next i
      CboTipDoc.ListIndex = i
   Else
      Txtipdoc = ""
   End If
   TxNumero = DataGrid1.Columns(4)
   Txtelefono = DataGrid1.Columns(5)
   TxObservacion = Trim(DataGrid1.Columns(6))
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
  Call desabilitatextos
  Call Limpiacampos
End Sub
Private Sub txBuscar_Change()
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
Private Sub cmdNuevo_Click()
    If wrights = "CON" Then
       MsgBox "Usuario No Autorizado", vbExclamation
       Exit Sub
    End If
  If Cmdnuevo.Caption = "&Nuevo" Then
    Limpiacampos
    habilitatextos
    TxTipo.SetFocus
    CmdEliminar.Enabled = False
    Sw = True
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
  Else                  'Grabar
    On Error GoTo err_trans
    Cnx0.BeginTrans
    Dim sql As String
    If Sw Then
       sql = "INSERT INTO ProfSalud(tipo,codigo,tipdoc,numero,descripcion,telefono,observacion,usuario,fecha,estacion) "
       sql = sql + " VALUES('" + UCase(TxTipo) + "','" + UCase(TxCodigo) + "','" + Txtipdoc + "','" + TxNumero + "','" + Trim(Txdescripcion) + "','" + wUsuario + "','" + Format(Date, "yyyymmdd") + "','" + w_clienName + "')"
    Else
       sql = "UPDATE ProfSalud SET tipdoc='" + Txtipdoc + "',numero='" + TxNumero + "',descripcion='" + UCase(Trim(Txdescripcion)) + "',usuario='" + wUsuario + "',fecha='" + Format(Date, "yyyymmdd") + "',estacion='" + w_clienName + "'"
       sql = sql + " WHERE Tipo='" + UCase(TxTipo) + "' and  codigo='" + UCase(TxCodigo) + "'"
    End If
    Cnx0.Execute sql
    Cnx0.CommitTrans
    If Sw Then
       DD.AddNew
       DD!tipo = UCase(TxTipo)
       DD!Codigo = UCase(TxCodigo)
    End If
    DD!tipdoc = Trim(Txtipdoc)
    DD!numero = Trim(TxNumero)
    DD!descripcion = UCase(Trim(Txdescripcion))
    DD!telefono = Trim(Txtelefono)
    DD!observacion = UCase(Trim(TxObservacion))
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
    TxTipo.Locked = True
    TxCodigo.Locked = True
    CboTipDoc.SetFocus
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
  If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
     Cnx0.Execute "DELETE FROM ProfSalud WHERE tipo='" & TxTipo & "' and codigo='" & TxCodigo & "'"
     DD.Delete
     Limpiacampos
  End If
End Sub
Private Sub cmdsalir_Click()
  Unload Me
  Me.Hide
  Set rs = Nothing
End Sub
Sub habilitatextos()
  TxTipo.Locked = False
  TxCodigo.Locked = False
  Txdescripcion.Locked = False
  CboTipDoc.Enabled = True
  TxNumero.Locked = False
  Txtelefono.Locked = False
  TxObservacion.Locked = False
End Sub
Sub desabilitatextos()
  TxTipo.Locked = True
  TxCodigo.Locked = True
  Txdescripcion.Locked = True
  CboTipDoc.Enabled = False
  TxNumero.Locked = True
  Txtelefono.Locked = True
  TxObservacion.Locked = True
End Sub
Sub Limpiacampos()
  TxTipo = "":   TxCodigo = ""
  CboTipDoc.Text = "": Txtipdoc = "": TxNumero = ""
  Txtelefono = "": Txdescripcion = "": TxObservacion = ""
End Sub

Private Sub TxTipo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     TxTipo = UCase(TxTipo)
     TxCodigo.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Txcodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     TxCodigo = UCase(TxCodigo)
     CboTipDoc.SetFocus
     Exit Sub
   End If
End Sub
Private Sub CboTipDoc_Click()
  If Trim(CboTipDoc.Text) = "" Then
     Txtipdoc = ""
     TxNumero = ""
  Else
     'Extrae codigo
     j = CboTipDoc.ListIndex
     nposicion = InStr(CboTipDoc.List(j), ":")
     nlen = Len(CboTipDoc.List(j))
     Txtipdoc = Right(CboTipDoc.List(j), nlen - nposicion)
     TxNumero.SetFocus
  End If
End Sub
Private Sub CboTipDoc_Change()
   Call CboTipDoc_Click
End Sub

Private Sub Txnumero_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Exit Sub
   End If
   If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not KeyAscii = 8 Then
       KeyAscii = 0
   End If
End Sub
Private Sub Txdescripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txdescripcion = UCase(Txdescripcion)
     Exit Sub
   End If
End Sub
Private Sub Txdescripcion_LostFocus()
   Txdescripcion_KeyPress (13)
End Sub
Private Sub TxObservacion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     TxObservacion = UCase(TxObservacion)
     Cmdnuevo.SetFocus
     Exit Sub
   End If
End Sub






