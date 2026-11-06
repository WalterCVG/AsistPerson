VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tTarifas 
   Caption         =   "Tarifas"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "tTarifas.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txBuscar 
      Height          =   285
      Left            =   4920
      MaxLength       =   25
      TabIndex        =   10
      Top             =   4680
      Width           =   3975
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   735
      Left            =   10320
      Picture         =   "tTarifas.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cerrar"
      Top             =   7680
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   1440
      TabIndex        =   15
      Top             =   5640
      Width           =   7695
      Begin VB.TextBox TxBaseLegal 
         Height          =   285
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox TxNocheFer 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   8
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TxDiaFer 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox TxNocheOrd 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Txorden 
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Txdescripcion 
         Height          =   285
         Left            =   3840
         MaxLength       =   25
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox TxNiveles 
         Height          =   285
         Left            =   1200
         MaxLength       =   21
         TabIndex        =   4
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox TxDiaOrd 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Base Legal"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Noche Feriado"
         Height          =   255
         Left            =   2520
         TabIndex        =   23
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Dia Feriado"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Noche Ordinario"
         Height          =   375
         Left            =   2520
         TabIndex        =   21
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Orden"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion Cargo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Niveles"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Dia Ordinario"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   9600
      TabIndex        =   0
      Top             =   5280
      Width           =   1575
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4095
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   11295
      _ExtentX        =   19923
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label LblDS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LblDS"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   9720
      TabIndex        =   25
      Top             =   8640
      Width           =   2055
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion a Ubicar :"
      ForeColor       =   &H80000001&
      Height          =   195
      Left            =   2520
      TabIndex        =   20
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   7515
   End
End
Attribute VB_Name = "tTarifas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean
Private Sub Form_Load()
     LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
    Set DD = New Recordset
    With DD
      .Fields.Append "Orden", adInteger
      .Fields.Append "descripcion", adChar, 25
      .Fields.Append "Niveles", adChar, 21
      .Fields.Append "DiaOrd", adDouble
      .Fields.Append "NocheOrd", adDouble
      .Fields.Append "DiaFer", adDouble
      .Fields.Append "NocheFer", adDouble
      .Open
    End With
    Set DataGrid1.DataSource = DD
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         '.Caption = "Orden"
          .Width = Me.TextWidth(String(6, "S"))
         Case 1
          '.Caption = "descripcion"
          .Width = Me.TextWidth(String(25, "S"))
         Case 2
          '.Caption = "descripcion"
          .Width = Me.TextWidth(String(21, "S"))
         Case 3
          '.Caption = "descripcion"
          .Width = Me.TextWidth(String(13, "S"))
         End Select
         '.NumberFormat = "##,###.99"
         .Locked = True
     End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM CpTarifas WHERE (baja is NULL)"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!orden = rs!orden
       DD!descripcion = "" + rs!descripcion
       DD!Niveles = "" + rs!Niveles
       DD!DiaOrd = rs!DiaOrd
       DD!NocheOrd = rs!NocheOrd
       DD!DiaFer = rs!DiaFer
       DD!NocheFer = rs!NocheFer
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
   Txbuscar = ""
   Txorden = DataGrid1.Columns(0)
   Txdescripcion = DataGrid1.Columns(1)
   TxNiveles = DataGrid1.Columns(2)
   TxDiaOrd = DataGrid1.Columns(3)
   TxNocheOrd = DataGrid1.Columns(4)
   TxDiaFer = DataGrid1.Columns(5)
   TxNocheFer = DataGrid1.Columns(6)
   
   Set rs = New Recordset
   sql = "SELECT * FROM CpTarifas WHERE (baja is NULL) and Orden =" & Val(Txorden)
   rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If Not rs.EOF Then
      TxBaseLegal = IIf(IsNull(rs!BaseLegal), "", UCase(Trim(rs!BaseLegal)))
   End If
   CmdEditar.Enabled = True
   CmdEliminar.Enabled = True
   Sw = False
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
    Txorden.Locked = False
    Txorden.SetFocus
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
    sql = "SELECT * FROM CpTarifas WHERE (baja is NULL) and Orden =" & Val(Txorden)
    rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    If Sw Then
       If Not rs.EOF Then
          MsgBox "Registro ya Existe", vbExclamation
          Exit Sub
       End If
       rs.AddNew
       rs!orden = Txorden
       rs!Registro = wUsuario
       rs!fecRegistro = Date
    Else
       rs!usrCambia = wUsuario
       rs!fecCambia = Date
       rs!horCambia = Format(Time, "HH:mm:ss")
    End If
    rs!descripcion = "" + Trim(UCase(Txdescripcion))
    rs!BaseLegal = "" + Trim(UCase(TxBaseLegal))
    rs!Niveles = "" + Trim(UCase(TxNiveles))
    rs!DiaOrd = 0 + Val(TxDiaOrd)
    rs!NocheOrd = 0 + Val(TxNocheOrd)
    rs!DiaFer = 0 + Val(TxDiaFer)
    rs!NocheFer = 0 + Val(TxNocheFer)
    rs!Registro = wUsuario
    rs!fecRegistro = Date
    rs.Update
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!orden = Txorden
    End If
    DD!descripcion = Trim(Txdescripcion)
    DD!Niveles = Trim(TxNiveles)
    DD!DiaOrd = Val(TxDiaOrd)
    DD!NocheOrd = Val(TxNocheOrd)
    DD!DiaFer = Val(TxDiaFer)
    DD!NocheFer = Val(TxNocheFer)
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
    If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
       sql = "UPDATE Cptarifas SET baja='X',usrbaja='" + wUsuario + "',fecBaja='" + Format(Date, "yyyy-mm-dd") + "',horbaja='" + Format(Time, "HH:mm:ss") + "' WHERE Orden=" & Val(Txorden)
       Cnx.Execute sql
       DD.Delete
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
    TxBaseLegal.Locked = False
    TxNiveles.Locked = False
    TxDiaOrd.Locked = False
    TxNocheOrd.Locked = False
    TxDiaFer.Locked = False
    TxNocheFer.Locked = False
End Sub
Sub desabilitatextos()
   Txorden.Locked = True
   Txdescripcion.Locked = True
   TxBaseLegal.Locked = True
   TxNiveles.Locked = True
   TxDiaOrd.Locked = True
   TxNocheOrd.Locked = True
   TxDiaFer.Locked = True
   TxNocheFer.Locked = True
End Sub
Sub Limpiacampos()
  Txbuscar = ""
  Txorden = "":  Txdescripcion = "": TxBaseLegal = "": TxNiveles = ""
  TxDiaOrd = "": TxNocheOrd = "": TxDiaFer = "": TxNocheFer = ""
End Sub
Private Sub Txorden_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Val(Txorden) >= 0 Then
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
     TxNiveles.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Txbaselegal_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     TxBaseLegal = UCase(TxBaseLegal)
     Exit Sub
   End If
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

