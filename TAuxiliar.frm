VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tAuxiliar 
   Caption         =   "Tablas Auxiliares"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "TAuxiliar.frx":0000
   ScaleHeight     =   7395
   ScaleMode       =   0  'User
   ScaleWidth      =   8572.184
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   7320
      Picture         =   "TAuxiliar.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cerrar"
      Top             =   6720
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   480
      TabIndex        =   5
      Top             =   4560
      Width           =   5895
      Begin VB.TextBox TxGrupo 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Txcodigo 
         Height          =   285
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Txdescripcion 
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   6600
      TabIndex        =   1
      Top             =   4440
      Width           =   1575
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3975
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7011
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
End
Attribute VB_Name = "tAuxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean
Private Sub Form_Load()
    Set DD = New Recordset
    With DD
     .Fields.Append "grupo", adVarChar, 3
     .Fields.Append "Codigo", adVarChar, 7
     .Fields.Append "descripcion", adVarChar, 30
     .Open
    End With
    Set DataGrid1.DataSource = DD
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         .Caption = "Grupo"
          .Width = Me.TextWidth(String(4, "S"))
         Case 1
         .Caption = "Codigo"
          .Width = Me.TextWidth(String(6, "S"))
         Case 2
          .Caption = "Descripcion"
          .Width = Me.TextWidth(String(30, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM CpTablasAux "
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!grupo = rs!grupo
       DD!codigo = rs!codigo
       DD!descripcion = rs!descripcion
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
    desabilitatextos
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   TxGrupo.Text = DataGrid1.Columns(0)
   Txcodigo.Text = DataGrid1.Columns(1)
   Txdescripcion.Text = DataGrid1.Columns(2)
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
Private Sub cmdNuevo_Click()
    If wrights = "CON" Then
       MsgBox "Usuario No Autorizado", vbExclamation
       Exit Sub
    End If
  If Cmdnuevo.Caption = "&Nuevo" Then
    Limpiacampos
    habilitatextos
    TxGrupo.SetFocus
    CmdEliminar.Enabled = False
    Sw = True
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
  Else                  'Grabar
    On Error GoTo err_trans
    Cnx.BeginTrans
    Dim sql As String
    If Sw Then
       sql = "INSERT INTO CpTablasAux(Grupo,codigo,descripcion,usuario,fecha) "
       sql = sql + " VALUES('" + UCase(TxGrupo) + "','" + UCase(Txcodigo) + "','" + Txdescripcion + "','" + wUsuario + "','" + Format(Date, "yyyymmdd") + "')"
    Else
       sql = "UPDATE CpTablasAux SET descripcion='" + UCase(Txdescripcion.Text) + "',usuario='" + wUsuario + "',fecha='" + Format(Date, "yyyymmdd") + "'"
       sql = sql + " WHERE grupo='" + UCase(TxGrupo) + "' and  codigo='" + UCase(Txcodigo) + "'"
    End If
    Cnx.Execute sql
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!grupo = UCase(TxGrupo)
       DD!codigo = UCase(Txcodigo)
    End If
    DD!descripcion = UCase(Txdescripcion)
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
    TxGrupo.Locked = True
    Txcodigo.Locked = True
    Txdescripcion.SetFocus
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
     Cnx.Execute "DELETE FROM CpTablasAux WHERE grupo='" & TxGrupo.Text & "' and codigo='" & Txcodigo.Text & "'"
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
  TxGrupo.Locked = False
  Txcodigo.Locked = False
  Txdescripcion.Locked = False
End Sub
Sub desabilitatextos()
  TxGrupo.Locked = True
  Txcodigo.Locked = True
  Txdescripcion.Locked = True
End Sub
Sub Limpiacampos()
  TxGrupo.Text = ""
  Txcodigo.Text = ""
  Txdescripcion.Text = ""
End Sub




Private Sub Txgrupo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     TxGrupo.Text = UCase(TxGrupo.Text)
     Txcodigo.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Txcodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txcodigo.Text = UCase(Txcodigo.Text)
     Txdescripcion.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Txdescripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txdescripcion.Text = UCase(Txdescripcion.Text)
     Cmdnuevo.SetFocus
     Exit Sub
   End If
End Sub
