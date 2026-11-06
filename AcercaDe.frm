VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form AcercaDe 
   Caption         =   "Tabla de Contenidos"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "AcercaDe.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   9360
      TabIndex        =   7
      Top             =   5280
      Width           =   1455
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Width           =   8895
      Begin VB.TextBox TxDescripcion 
         Height          =   2055
         Left            =   1200
         MaxLength       =   1500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   960
         Width           =   7335
      End
      Begin VB.TextBox Txsecuencia 
         Height          =   285
         Left            =   2760
         MaxLength       =   3
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox TxSumilla 
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   35
         TabIndex        =   4
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox Txgrupo 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   3
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Contenido"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Secuencia"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Sumilla"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3720
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Grupo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   9720
      Picture         =   "AcercaDe.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cerrar"
      Top             =   7560
      Width           =   615
   End
   Begin VB.TextBox txBuscar 
      Height          =   285
      Left            =   3600
      MaxLength       =   35
      TabIndex        =   0
      Top             =   3360
      Width           =   4335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
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
      Caption         =   "Contenida Acerca de"
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
      Caption         =   "Sumilla a Ubicar :"
      ForeColor       =   &H80000001&
      Height          =   195
      Left            =   2160
      TabIndex        =   12
      Top             =   3360
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   6195
   End
End
Attribute VB_Name = "AcercaDe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim sql As String
Dim Sw As Boolean
Private Sub Form_Load()
    Set DD = New Recordset
    With DD
      .Fields.Append "Grupo", adVarChar, 2
      .Fields.Append "Secuencia", adVarChar, 3
      .Fields.Append "Sumilla", adVarChar, 35
      .Fields.Append "descripcion", adVarChar, 1500
      .Open
    End With
    Set DataGrid1.DataSource = DD
    'Configura columnas
    For i = 0 To DataGrid1.Columns.Count - 1
       With DataGrid1.Columns(i)
         Select Case i
         Case 0
          .Caption = "Grupo"
          .Width = Me.TextWidth(String(5, "S"))
         Case 1
          .Caption = "Secuencia"
          .Width = Me.TextWidth(String(6, "S"))
         Case 2
          .Caption = "Sumilla"
          .Width = Me.TextWidth(String(35, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
       End With
    Next
    
    Set rs = New Recordset
    sql = "SELECT * FROM CpAyudas ORDER BY secuencia"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!grupo = rs!grupo
       DD!secuencia = rs!secuencia
       DD!sumilla = rs!sumilla
       DD!descripcion = rs!descripcion
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
End Sub
Private Sub DataGrid1_Click()
   Txgrupo.Text = DataGrid1.Columns(0)
   Txsecuencia.Text = DataGrid1.Columns(1)
   TxSumilla.Text = DataGrid1.Columns(2)
   Txdescripcion.Text = DataGrid1.Columns(3)
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
    Txgrupo.SetFocus
    CmdEliminar.Enabled = False
    Sw = True
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
  Else                  'Grabar
    On Error GoTo err_trans
    If Trim(Txgrupo.Text) = "" Then
       MsgBox ("Favor Verifique Informacion")
       Exit Sub
    End If
    Cnx.BeginTrans
    Dim sql As String
    If Sw Then
       sql = "INSERT INTO CpAyudas(grupo,secuencia,sumilla,descripcion) "
       sql = sql + " VALUES('" + Txgrupo.Text + "','" + Txsecuencia.Text + "','" + TxSumilla.Text + "','" + Txdescripcion.Text + "')"
    Else
       sql = "UPDATE CpAyudas SET sumilla='" + TxSumilla.Text + "' , descripcion='" + Txdescripcion.Text + "'"
       sql = sql + " WHERE grupo='" + Txgrupo.Text + "' AND secuencia='" + Txsecuencia.Text + "'"
    End If
    Cnx.Execute sql
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!grupo = Txgrupo.Text
       DD!secuencia = Txsecuencia.Text
    End If
    DD!sumilla = TxSumilla.Text
    DD!descripcion = Txdescripcion.Text
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
    Txgrupo.Locked = True
    Txsecuencia.Locked = True
    TxSumilla.SetFocus
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
       Cnx.Execute "DELETE FROM CpAyudas WHERE grupo='" & Txgrupo.Text & "' AND secuencia='" & Txsecuencia.Text & "'"
       DD.Delete
       Limpiacampos
    End If
End Sub
Private Sub cmdsalir_Click()
  Unload Me
  Me.Hide
End Sub
Sub habilitatextos()
  Txgrupo.Locked = False
  Txsecuencia.Locked = False
  TxSumilla.Locked = False
  Txdescripcion.Locked = False
End Sub
Sub desabilitatextos()
  Txgrupo.Locked = True
  Txsecuencia.Locked = True
  TxSumilla.Locked = True
  Txdescripcion.Locked = True
End Sub
Sub Limpiacampos()
  Txgrupo.Text = "": Txsecuencia.Text = ""
  TxSumilla.Text = "": Txdescripcion.Text = ""
End Sub
Private Sub TxBuscar_Change()
    Dim cBuscar As String
    cpuntero = DataGrid1.Bookmark
    If Txbuscar.Text = "" Then Exit Sub
    cBuscar = "sumilla like '%" & Txbuscar.Text & "%'"
    DD.MoveFirst
    DD.Find (cBuscar)
    If DD.EOF Then
       Txbuscar.BackColor = &HFF&
    Else
       Txbuscar.BackColor = &H80000005
       DataGrid1.Refresh
    End If
End Sub



Private Sub Txgrupo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txgrupo.Text = Right("00" + Trim(Txgrupo.Text), 2)
     Txsecuencia.SetFocus
     Exit Sub
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
       KeyAscii = 0
   End If
End Sub
Private Sub Txsecuencia_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     'Txsecuencia.Text = Right("000" + Trim(Txsecuencia.Text), 3)
     TxSumilla.SetFocus
     Exit Sub
   End If
   'If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
   '    KeyAscii = 0
   'End If
End Sub
Private Sub Txsumilla_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txdescripcion.SetFocus
     Exit Sub
   End If
End Sub



