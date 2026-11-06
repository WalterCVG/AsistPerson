VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tCargos 
   Caption         =   "Cargos"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "tCargos.frx":0000
   ScaleHeight     =   7950
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txBuscar 
      Height          =   285
      Left            =   4200
      TabIndex        =   12
      Top             =   4320
      Width           =   4935
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   9960
      Picture         =   "tCargos.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cerrar"
      Top             =   7200
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   1080
      TabIndex        =   4
      Top             =   5160
      Width           =   8055
      Begin VB.TextBox Txcodigo 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Txdescripcion 
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         Top             =   960
         Width           =   6015
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   9480
      TabIndex        =   0
      Top             =   4800
      Width           =   1575
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   480
      TabIndex        =   10
      Top             =   480
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   6165
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion a Ubicar :"
      ForeColor       =   &H80000001&
      Height          =   195
      Left            =   1320
      TabIndex        =   11
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   8355
   End
End
Attribute VB_Name = "tCargos"
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
      .Fields.Append "Codigo", adChar, 4
      .Fields.Append "descripcion", adChar, 50
      .Open
    End With
    Set DataGrid1.DataSource = DD
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         '.Caption = "Codigo"
          .Width = Me.TextWidth(String(6, "S"))
         Case 1
          '.Caption = "descripcion"
          .Width = Me.TextWidth(String(50, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM CpCargos"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!codigo = rs!codigo
       DD!descripcion = "" + rs!descripcion
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
    desabilitatextos
    Limpiacampos
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Txbuscar.Text = ""
   Txcodigo.Text = DataGrid1.Columns(0)
   Txdescripcion.Text = DataGrid1.Columns(1)
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
    Txcodigo.Locked = False
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
    If Sw Then
       sql = "INSERT INTO CpCargos(Codigo,descripcion,usuario,fecha) "
       sql = sql + " VALUES('" + Txcodigo.Text + "','" + Trim(Txdescripcion.Text) + "','" + wUsuario + "','" + Format(Date, "dd/mm/yyyy") + "')"
    Else
       sql = "UPDATE CpCargos SET descripcion='" + Trim(Txdescripcion.Text) + "',usuario='" + wUsuario + "',fecha='" + Format(Date, "dd/mm/yyyy") + "'"
       sql = sql + " WHERE Codigo='" + Txcodigo.Text & "'"
    End If
    Cnx.Execute sql
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!codigo = Txcodigo.Text
    End If
    DD!descripcion = Trim(Txdescripcion.Text)
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
    Set rs = New Recordset
    sql = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and (cargo='" & Trim(Txcodigo.Text) & "')"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If r_d.EOF Then
    Else
       MsgBox "Hay Registros con dicho Cargo", vbExclamation
       rs.Close
       Exit Sub
    End If
    rs.Close
    If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
      Cnx.Execute "DELETE FROM CpCargos WHERE codigo='" & Txcodigo.Text & "'"
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
    Txdescripcion.Locked = False
End Sub
Sub desabilitatextos()
   Txcodigo.Locked = True
   Txdescripcion.Locked = True
End Sub
Sub Limpiacampos()
  Txbuscar.Text = ""
  Txcodigo.Text = "":  Txdescripcion.Text = ""
End Sub


Private Sub Txcodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Val(Txcodigo.Text) >= 0 Then
     Txdescripcion.SetFocus
     Exit Sub
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
       KeyAscii = 0
   End If
End Sub
Private Sub Txdescripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txdescripcion.Text = Trim(Txdescripcion.Text)
     Cmdnuevo.SetFocus
     Exit Sub
   End If
End Sub
Private Sub TxBuscar_Change()
    Dim cBuscar As String
    cpuntero = DataGrid1.Bookmark
    If Txbuscar.Text = "" Then Exit Sub
    cBuscar = "descripcion like '*" & Txbuscar.Text & "*'"
    DD.MoveFirst
    DD.Find (cBuscar)
    If DD.EOF Then
       Txbuscar.BackColor = &HFF&
    Else
       Txbuscar.BackColor = &H80000005
       DataGrid1.Refresh
    End If
End Sub
