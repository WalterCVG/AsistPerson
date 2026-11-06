VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form TCptablas 
   Caption         =   "Tablas CP"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "TCptablas.frx":0000
   ScaleHeight     =   8790
   ScaleWidth      =   9975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   8400
      Picture         =   "TCptablas.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cerrar"
      Top             =   7920
      Width           =   735
   End
   Begin VB.TextBox txBuscar 
      Height          =   285
      Left            =   3840
      MaxLength       =   30
      TabIndex        =   16
      Top             =   4920
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   7920
      TabIndex        =   7
      Top             =   5640
      Width           =   1575
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   840
      TabIndex        =   0
      Top             =   5640
      Width           =   6375
      Begin VB.TextBox Txbloqueo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6000
         MaxLength       =   1
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bloqueado"
         Height          =   255
         Left            =   4920
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Txcampo 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox TxValor 
         Height          =   285
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   13
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox Txdescripcion 
         Height          =   285
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1080
         Width           =   3735
      End
      Begin VB.TextBox Txcodigo 
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Txtabla 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Campo"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Valor"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   960
      TabIndex        =   11
      Top             =   360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7223
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
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1500.095
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Denominación a Ubicar :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1560
      TabIndex        =   17
      Top             =   4920
      Width           =   1755
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   7515
   End
End
Attribute VB_Name = "TCptablas"
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
     .Fields.Append "TABLA", adVarChar, 2
     .Fields.Append "codigo", adVarChar, 2
     .Fields.Append "descripcion", adVarChar, 30
     .Fields.Append "VALOR", adVarChar, 30
     .Fields.Append "CAMPO", adVarChar, 10
     .Fields.Append "Bloqueo", adVarChar, 1
     .Open
    End With
    Set DataGrid1.DataSource = DD
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         '.Caption = "TABLA"
          .Width = Me.TextWidth(String(5, "S"))
         Case 1
         '.Caption = "codigo"
          .Width = Me.TextWidth(String(5, "S"))
         Case 2
          '.Caption = "Descripcion"
          .Width = Me.TextWidth(String(30, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    
    Set rs = New Recordset
    sql = "SELECT * FROM cpTablas ORDER BY tabla,codigo"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!tabla = rs!tabla
       DD!codigo = "" + rs!codigo
       DD!descripcion = rs!descripcion
       DD!Valor = IIf(IsNull(rs!Valor), "", rs!Valor)
       DD!campo = IIf(IsNull(rs!campo), "", rs!campo)
       DD!Bloqueo = IIf(IsNull(rs!Bloqueo), "", rs!Bloqueo)
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
    desabilitatextos
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Txtabla = DataGrid1.Columns(0)
   Txcodigo = DataGrid1.Columns(1)
   Txdescripcion = DataGrid1.Columns(2)
   TxValor = DataGrid1.Columns(3)
   Txcampo = DataGrid1.Columns(4)
   Txbloqueo = DataGrid1.Columns(5)
   Check1.Value = IIf(Txbloqueo = "S", 1, 0)
   CmdEditar.Enabled = True
   CmdEliminar.Enabled = True
End Sub
Private Sub cmdNuevo_Click()
    If wrights = "CON" Then
       MsgBox "Usuario No Autorizado", vbExclamation
       Exit Sub
    End If
  If Cmdnuevo.Caption = "&Nuevo" Then
    Limpiacampos
    habilitatextos
    Txtabla.SetFocus
    CmdEliminar.Enabled = False
    Sw = True
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
  Else                  'Grabar
    On Error GoTo err_trans
    Cnx.BeginTrans
    Dim sql As String
    Txbloqueo = IIf(Check1.Value = 1, "S", "N")
    'cfecha = Format(Date, "dd/mm/yyyy")
    If Sw Then
       sql = "INSERT INTO CPTablas(tabla,codigo,descripcion,valor,campo,bloqueo,usuario,fecha) "
       sql = sql + " values('" + Txtabla + "','" + Txcodigo + "','" + Txdescripcion + "','" + TxValor + "','" + Txcampo + "','" + Txbloqueo + "','" + wUsuario + "','" + Format(Date, "yyyymmdd") + "')"
    Else
       sql = "UPDATE CpTablas SET descripcion='" + Txdescripcion + "',valor='" + TxValor + "',campo='" + Txcampo + "',bloqueo='" + Txbloqueo + "',usuario='" + wUsuario + "',fecha='" + Format(Date, "yyyymmdd") + "'"
       sql = sql + " WHERE tabla='" + Txtabla + "' and  codigo='" + Txcodigo + "'"
    End If
    Cnx.Execute sql
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!tabla = Txtabla
       DD!codigo = Txcodigo
    End If
    DD!descripcion = IIf(IsNull(Txdescripcion), "", Txdescripcion)
    DD!Valor = IIf(IsNull(TxValor), "", TxValor)
    DD!campo = IIf(IsNull(Txcampo), "", Txcampo)
    DD!Bloqueo = IIf(IsNull(Txbloqueo), "N", Txbloqueo)
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
    Txtabla.Locked = True
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
     Cnx.Execute "DELETE FROM CpTablas WHERE tabla='" & Txtabla & "' and codigo='" & Txcodigo & "'"
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
  Txtabla.Locked = False
  Txcodigo.Locked = False
  Txdescripcion.Locked = False
  TxValor.Locked = False
  Txcampo.Locked = False
  Check1.Enabled = True
End Sub
Sub desabilitatextos()
  Txtabla.Locked = True
  Txcodigo.Locked = True
  Txdescripcion.Locked = True
  TxValor.Locked = True
  Txcampo.Locked = True
  Check1.Enabled = False
End Sub
Sub Limpiacampos()
  Txtabla = "":    Txcodigo = ""
  Txdescripcion = "": TxValor = "": Txcampo = "": Check1.Value = 0: Txbloqueo = ""
End Sub

Private Sub TxBuscar_Change()
     Dim cBuscar As String
     cpuntero = DataGrid1.Bookmark
     If Txbuscar = "" Then Exit Sub
     cBuscar = "Descripcion like '" & Trim(Txbuscar) & "*'"
     DD.MoveFirst
     DD.Find (cBuscar)
     If DD.EOF Then
        Txbuscar.BackColor = &HFF&
     Else
        Txbuscar.BackColor = &H80000005
        DataGrid1.Refresh
     End If
End Sub

Private Sub Txtabla_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txcodigo.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Txcodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txdescripcion.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Txdescripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     TxValor.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Txvalor_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txcampo.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Txcampo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Cmdnuevo.SetFocus
     Exit Sub
   End If
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


