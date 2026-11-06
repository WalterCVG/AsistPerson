VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form TUndPrgRol 
   Caption         =   "Servicio - U.P.R."
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "TUndPrgRol.frx":0000
   ScaleHeight     =   7815
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   9120
      TabIndex        =   11
      Top             =   4440
      Width           =   1575
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
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
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
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
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   960
      TabIndex        =   8
      Top             =   4440
      Width           =   8055
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Miembros tienen derecho a Horas Complementarias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   5415
      End
      Begin VB.TextBox TxServicioSuSalud 
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
         Left            =   1920
         MaxLength       =   65
         TabIndex        =   5
         Top             =   1320
         Width           =   5895
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "OFICINA  -   Jefaturas  de Servicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
      Begin VB.TextBox Txdescripcion 
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
         Left            =   1920
         MaxLength       =   55
         TabIndex        =   4
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox Txcodupr 
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
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Servicio SUSALUD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
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
      Height          =   495
      Left            =   9720
      Picture         =   "TUndPrgRol.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Cerrar"
      Top             =   7200
      Width           =   615
   End
   Begin VB.TextBox txBuscar 
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
      Left            =   4800
      MaxLength       =   20
      TabIndex        =   1
      Top             =   3960
      Width           =   4455
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
      Left            =   0
      Picture         =   "TUndPrgRol.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   960
      TabIndex        =   15
      Top             =   240
      Width           =   9615
      _ExtentX        =   16960
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
      Caption         =   "Contenido de tabla"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina - Jefaturas de Servicio :  la integran los jefes de UPR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6960
      Width           =   6015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contenido  a ubicar en descripción :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   240
      Left            =   1560
      TabIndex        =   16
      Top             =   3960
      Width           =   3180
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   8355
   End
End
Attribute VB_Name = "TUndPrgRol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean
Private Sub Command1_Click()
     c_grupo = "01": c_secuencia = "013"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub
Private Sub Form_Load()
    Set DD = New Recordset
    With DD
     .Fields.Append "codupr", adVarChar, 2
     .Fields.Append "descripcion", adVarChar, 55
     .Fields.Append "UPRJefaturas", adVarChar, 1
     .Fields.Append "ServicioSusalud", adVarChar, 65
     .Fields.Append "HrComplementaria", adVarChar, 1
     .Open
    End With
    Set DataGrid1.DataSource = DD
    'Configura columnas
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         .Caption = "Codigo"
          .Width = Me.TextWidth(String(6, "S"))
         Case 1
          .Caption = "Descripcion"
          .Width = Me.TextWidth(String(55, "S"))
         Case 2
          .Caption = "UPRJefaturas"
          .Width = Me.TextWidth(String(10, "S"))
         Case 4
          .Caption = "Hr.Compl."
          .Width = Me.TextWidth(String(10, "S"))
         Case Else
          .Visible = False
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    
    Set rs = New Recordset
    sql = "SELECT * FROM CpUndPrgRol order by Descripcion"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!CodUpr = rs!CodUpr
       DD!descripcion = rs!descripcion
       DD!UPRJefaturas = IIf(IsNull(rs!UPRJefaturas), "", Trim(rs!UPRJefaturas))
       DD!servicioSusalud = IIf(IsNull(rs!servicioSusalud), "", Trim(rs!servicioSusalud))
       DD!hrcomplementaria = IIf(IsNull(rs!hrcomplementaria), "", Trim(rs!hrcomplementaria))
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
    DD.MoveFirst
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Txcodupr = DataGrid1.Columns(0)
   Txdescripcion = DataGrid1.Columns(1)
   Check1.Value = IIf(DataGrid1.Columns(2) = "S", 1, 0)
   TxServicioSuSalud = DataGrid1.Columns(3)
   Check2.Value = IIf(DataGrid1.Columns(4) = "S", 1, 0)
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
  Dim cHRcompl  As String
  If Cmdnuevo.Caption = "&Nuevo" Then
    Limpiacampos
    habilitatextos
    Txcodupr.SetFocus
    CmdEliminar.Enabled = False
    Sw = True
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
  Else                  'Grabar
    On Error GoTo err_trans
    If Trim(Txcodupr) = "" Then
       MsgBox ("Favor Verifique Informacion")
       Exit Sub
    End If
    cUPRjefaturas = IIf(Check1.Value = 0, "N", "S")
    cHRcompl = IIf(Check2.Value = 0, "N", "S")
    Cnx.BeginTrans
    Dim sql As String
    If Sw Then
       sql = "INSERT INTO CpUndPrgRol(codupr,descripcion,UPRJefaturas,servicioSusalud,HRComplementaria,usuario, fecha) "
       sql = sql + " VALUES('" + Txcodupr + "','" + Trim(Txdescripcion) + "','" + cUPRjefaturas + "','" + Trim(TxServicioSuSalud) + "','" + cHRcompl + "','" + wUsuario + "','" + Format(Date, "yyyy-mm-dd") + "')"
    Else
       sql = "UPDATE CpundPrgRol SET descripcion='" + Trim(Txdescripcion) + "',UPRjefaturas='" + cUPRjefaturas + "',ServicioSusalud='" + Trim(TxServicioSuSalud) + "',HRcomplementaria='" + cHRcompl + "',usuario='" + wUsuario + "',fecha='" + Format(Date, "yyyy-mm-dd") + "'"
       sql = sql + " WHERE codupr='" + Txcodupr + "'"
    End If
    Cnx.Execute sql
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!CodUpr = Txcodupr
    End If
    DD!descripcion = Trim(Txdescripcion)
    DD!UPRJefaturas = cUPRjefaturas
    DD!servicioSusalud = Trim(TxServicioSuSalud)
    DD!hrcomplementaria = cHRcompl
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
    Txcodupr.Locked = True
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
    'Set R_d = New ADODB.Recordset
    'sql = "SELECT * FROM OGCregi" & genaño & "'  WHERE (servorigen='" & Txcodigo.Text & "')"
    'R_d.Open sql, Cnx, adOpenStatic, adlockreadonly, adCmdText
    'If R_d.EOF  Then
    'Else
    '   MsgBox "Hay Gestiones con dicha Servicio", vbExclamation
    '   R_d.Close
    '   Exit Sub
    'End If
    'R_d.Close
    
    If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
       Cnx.Execute "DELETE FROM CpUndPrgRol WHERE codupr='" & Txcodupr.Text & "'"
       DD.Delete
       Limpiacampos
    End If
End Sub

Private Sub cmdsalir_Click()
  Unload Me
  Me.Hide
End Sub
Sub habilitatextos()
  Txcodupr.Locked = False
  Txdescripcion.Locked = False
  TxServicioSuSalud.Locked = False
  Check1.Enabled = True: Check2.Enabled = True
End Sub
Sub desabilitatextos()
  Txcodupr.Locked = True
  Txdescripcion.Locked = True
  TxServicioSuSalud.Locked = True
  Check1.Enabled = False: Check2.Enabled = False
End Sub
Sub Limpiacampos()
  Txcodupr = ""
  Txdescripcion = "": Check1.Value = 0: Check2.Value = 0
  cUPRjefaturas = "": TxServicioSuSalud = ""
End Sub
Private Sub TxBuscar_Change()
    Dim cBuscar As String
    cpuntero = DataGrid1.Bookmark
    If Txbuscar = "" Then Exit Sub
    cBuscar = "Descripcion like '%" & Txbuscar & "%'"
    DD.MoveFirst
    DD.Find (cBuscar)
    If DD.EOF Then
       Txbuscar.BackColor = &HFF&
    Else
       Txbuscar.BackColor = &H80000005
       DataGrid1.Refresh
    End If
End Sub
Private Sub Txcodupr_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txdescripcion.SetFocus
     Exit Sub
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
       KeyAscii = 0
   End If
End Sub
Private Sub Txdescripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Txdescripcion = UCase(Txdescripcion)
      Exit Sub
   End If
End Sub
Private Sub TxservicioSusalud_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxServicioSuSalud = UCase(TxServicioSuSalud)
      Cmdnuevo.SetFocus
     Exit Sub
   End If
End Sub



