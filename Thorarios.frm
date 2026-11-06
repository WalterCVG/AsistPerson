VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tHorarios 
   Caption         =   "Tabla Horarios"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Thorarios.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   11145
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxBuscar 
      Height          =   285
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   28
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txBuscar1 
      Height          =   285
      Left            =   4320
      TabIndex        =   27
      Top             =   4320
      Width           =   4935
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   9960
      Picture         =   "Thorarios.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Cerrar"
      Top             =   7560
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   720
      TabIndex        =   21
      Top             =   5040
      Width           =   8535
      Begin VB.ComboBox CboTurno 
         Height          =   315
         Left            =   6720
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "No Activo"
         Height          =   195
         Left            =   5880
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Asistencial"
         Height          =   195
         Left            =   6840
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TxHoras 
         Height          =   285
         Left            =   4920
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox TxTurno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6120
         MaxLength       =   2
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxSalida1 
         Height          =   285
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox TxIngreso1 
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox TxSalida 
         Height          =   285
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox TxCodigo 
         Height          =   315
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox TxIngreso 
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox TxDescripcion 
         Height          =   285
         Left            =   3960
         MaxLength       =   30
         TabIndex        =   4
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HHMM"
         Height          =   255
         Left            =   4920
         TabIndex        =   32
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HHMM"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "HHMM"
         Height          =   255
         Left            =   1080
         TabIndex        =   30
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Salida 1"
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   17
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Ingreso 1"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Horas"
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Turno"
         Height          =   255
         Left            =   6120
         TabIndex        =   19
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Salida"
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Ingreso"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   9480
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   24
      Top             =   480
      Width           =   10695
      _ExtentX        =   18865
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
   Begin VB.Label LblDS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LblDS"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   8760
      TabIndex        =   33
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Horario"
      Height          =   255
      Left            =   1320
      TabIndex        =   29
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Denominación:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3240
      TabIndex        =   26
      Top             =   4320
      Width           =   1065
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
Attribute VB_Name = "tHorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean
Dim casistencial, cbloqueo As String
Dim nposicion As Integer
Private Sub Form_Load()
    LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
    Set rs = New Recordset
    rs.Open "SELECT * FROM cptablas WHERE tabla='30' and codigo <>'..' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
    While Not rs.EOF
        ExisteYa = False
        For i = 0 To CboTurno.ListCount - 1
            If CboTurno.List(i) = rs.Fields("Descripcion") & " :" & rs.Fields("codigo") Then
               ExisteYa = True
               Exit For
            End If
        Next i
        If Not ExisteYa Then
           CboTurno.AddItem rs.Fields("Descripcion") & " :" & rs.Fields("codigo")
        End If
        rs.MoveNext
    Wend
    rs.Close
    Set DD = New Recordset
    With DD
      .Fields.Append "CODIGO", adChar, 4
      .Fields.Append "Descripcion", adChar, 30
      .Fields.Append "Asistencial", adChar, 1
      .Fields.Append "ingreso", adChar, 4
      .Fields.Append "Salida", adChar, 4
      .Fields.Append "Ingreso1", adChar, 4
      .Fields.Append "Salida1", adChar, 4
      .Fields.Append "HORAS", adChar, 4
      .Fields.Append "Turno", adChar, 2
      .Fields.Append "Bloqueo", adChar, 1
      .Open
    End With
    Set DataGrid1.DataSource = DD
    'Configura columnas
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
          '.Caption = "CODIGO"
          .Width = Me.TextWidth(String(7, "S"))
         Case 1
          '.Caption = "descripcion"
          .Width = Me.TextWidth(String(30, "S"))
         Case 2
         '.Caption = "asistencial"
         .Width = Me.TextWidth(String(9, "S"))
         Case 3
         '.Caption = "Ingreso"
          .Width = Me.TextWidth(String(6, "S"))
         Case 4
          '.Caption = "salida"
          .Width = Me.TextWidth(String(6, "S"))
         Case 5
          '.Caption = "ingreso1"
          .Width = Me.TextWidth(String(6, "S"))
         Case 6
          '.Caption = "Salida1"
          .Width = Me.TextWidth(String(6, "S"))
         Case 7
          '.Caption = "Horas"
          .Width = Me.TextWidth(String(6, "S"))
        End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM CpHorarioNew ORDER BY DESCRIPCION"        'CpHORARIOS"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!codigo = rs!codigo
       DD!descripcion = IIf(IsNull(rs!descripcion), "", rs!descripcion)
       DD!asistencial = IIf(IsNull(rs!asistencial), "", rs!asistencial)
       DD!ingreso = IIf(IsNull(rs!ingreso), "", rs!ingreso)
       DD!salida = IIf(IsNull(rs!salida), "", rs!salida)
       DD!Ingreso1 = IIf(IsNull(rs!Ingreso1), "", rs!Ingreso1)
       DD!salida1 = IIf(IsNull(rs!salida1), "", rs!salida1)
       DD!Horas = IIf(IsNull(rs!Horas), "", rs!Horas)
       DD!turno = IIf(IsNull(rs!turno), "", rs!turno)
       DD!Bloqueo = IIf(IsNull(rs!Bloqueo), "", rs!Bloqueo)
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
    desabilitatextos
    Limpiacampos
    DD.MoveFirst
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Txcodigo.Enabled = False
   Txbuscar.Text = "": Txbuscar1.Text = ""
   Txcodigo.Text = DataGrid1.Columns(0)
   Txdescripcion.Text = DataGrid1.Columns(1)
   Check1.Value = IIf(DataGrid1.Columns(2) = "S", 1, 0)
   TxIngreso.Text = DataGrid1.Columns(3)
   TxSalida.Text = DataGrid1.Columns(4)
   TxIngreso1.Text = DataGrid1.Columns(5)
   TxSalida1.Text = DataGrid1.Columns(6)
   TxHoras.Text = DataGrid1.Columns(7)
   Txturno.Text = DataGrid1.Columns(8)
   Check2.Value = IIf(DataGrid1.Columns(9) = "S", 1, 0)
   If Trim(Txturno.Text) <> "" Then
      For i = 0 To CboTurno.ListCount - 1
         npos = InStr(CboTurno.List(i), ":")
         nlen = Len(CboTurno.List(i))
         dato = Trim(Right(CboTurno.List(i), nlen - npos))
         If dato = Trim(Txturno.Text) Then
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
Private Sub cmdNuevo_Click()
  If wrights = "CON" Then
     MsgBox "Usuario No Autorizado", vbExclamation
     Exit Sub
  End If
  If Cmdnuevo.Caption = "&Nuevo" Then
    Limpiacampos
    habilitatextos
    Txcodigo.Enabled = True
    Sw = True
    CmdEliminar.Enabled = False
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
  Else                  'Grabar
    If Val(TxHoras.Text) < 100 Or Val(TxHoras.Text) > 3000 Or Len(TxHoras) > 4 Or Left(TxHoras, 2) > 30 Or Right(TxHoras, 2) > 59 Then
       MsgBox "Verifique formato Horas", vbExclamation
       Exit Sub
    End If
  
    On Error GoTo err_trans
    If Trim(Txcodigo.Text) = "" Then Exit Sub
    If Val(TxIngreso.Text) = 0 Or Val(TxSalida.Text) = 0 Or Val(TxHoras.Text) = 0 Then
       MsgBox "Verifique horario..", vbExclamation
       Exit Sub
    End If
    If Trim(Txturno.Text) = "" Then
       MsgBox "Indique Turno..", vbExclamation
       Exit Sub
    End If
    Cnx.BeginTrans
    Dim sql As String
    casistencial = IIf(Check1.Value = 1, "S", "N")
    cbloqueo = IIf(Check2.Value = 1, "S", "N")
    If Sw Then
       sql = "INSERT INTO CpHorarioNew(CODIGO,descripcion,asistencial,ingreso,salida,Ingreso1,Salida1,horas,Turno,bloqueo,usuario,fecha) " & _
       "VALUES('" + Txcodigo.Text + "','" + Txdescripcion.Text + "','" + casistencial + "','" + TxIngreso.Text + "','" + TxSalida.Text + "','" + TxIngreso1.Text + "','" + TxSalida1.Text + "','" + TxHoras.Text + "','" + Txturno.Text + "','" + cbloqueo + "','" + wUsuario + "','" + Format(Date, "yyyy-mm-dd") + "')"
    Else
       sql = "UPDATE CpHorarioNew SET Descripcion='" + UCase(Txdescripcion.Text) + "',asistencial='" + casistencial + "',Ingreso='" + TxIngreso.Text + "',Salida='" + TxSalida.Text + "',Ingreso1='" + TxIngreso1.Text + "',Salida1='" + TxSalida1.Text + "',Horas='" + TxHoras.Text + "',Turno='" + Txturno.Text + "',bloqueo='" + cbloqueo + "',usuario='" + wUsuario + "',fecha='" + Format(Date, "yyyy-mm-dd") + "'"
       sql = sql + " WHERE CODIGO='" + Txcodigo.Text & "'"
    End If
    Cnx.Execute sql
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!codigo = Txcodigo.Text
    End If
    DD!descripcion = UCase(Txdescripcion.Text)
    DD!asistencial = casistencial
    DD!ingreso = TxIngreso.Text
    DD!salida = TxSalida.Text
    DD!Ingreso1 = TxIngreso1.Text
    DD!salida1 = TxSalida1.Text
    DD!Horas = TxHoras.Text
    DD!turno = Txturno.Text
    DD!Bloqueo = cbloqueo
    DD.Update
    desabilitatextos
    Limpiacampos
    Txcodigo.Enabled = False
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
    Set r_d = New ADODB.Recordset
    sql = "SELECT * FROM CpHor" & genaño & "  WHERE (HORARIO='" & Txcodigo.Text & "')"
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If r_d.EOF Then
    Else
       MsgBox "Hay Programaciones con dicho HORARIO", vbExclamation
       r_d.Close
       Exit Sub
    End If
    r_d.Close
    If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
       Cnx.Execute "DELETE FROM CpHorarioNew WHERE CODIGO='" & Txcodigo.Text & "'"
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
    Check1.Enabled = True
    TxIngreso.Locked = False
    TxSalida.Locked = False
    TxIngreso1.Locked = False
    TxSalida1.Locked = False
    TxHoras.Locked = False
    Txturno.Locked = False
    CboTurno.Enabled = True
    Check2.Enabled = True
End Sub
Sub desabilitatextos()
    Txdescripcion.Locked = True
    Check1.Enabled = False
    TxIngreso.Locked = True
    TxSalida.Locked = True
    TxIngreso1.Locked = True
    TxSalida1.Locked = True
    TxHoras.Locked = True
    Txturno.Locked = True
    CboTurno.Enabled = False
    Check2.Enabled = False
End Sub
Sub Limpiacampos()
  Txcodigo.Enabled = False
  Txbuscar.Text = "": Txbuscar1.Text = ""
  Check1.Value = 0: casistencial = ""
  Txcodigo.Text = "": Txdescripcion.Text = ""
  TxIngreso.Text = "":  TxSalida.Text = ""
  TxIngreso1.Text = "": TxSalida1.Text = ""
  TxHoras.Text = "":  Txturno.Text = "": CboTurno.Text = ""
  Check2.Value = 0: cbloqueo = ""
End Sub
Private Sub Txcodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Trim(Txcodigo.Text) <> "" Then
      Set rs = New Recordset
      sql = "SELECT * FROM CpHor" & genaño & "  WHERE (HORARIO='" & Txcodigo.Text & "')"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         rs.Close
         Set rs = New Recordset
         sql = "SELECT * FROM CpHorarioNew  WHERE (codigo='" & Txcodigo.Text & "')"
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then
            MsgBox "<Error> Ya existe..", vbExclamation
            rs.Close
            Exit Sub
         End If
      Else
         MsgBox "<Error> Ya existe..", vbExclamation
         rs.Close
         Exit Sub
      End If
      rs.Close
      Txdescripcion.SetFocus
      Exit Sub
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
       KeyAscii = 0
   End If
End Sub

Private Sub Txcodigo_LostFocus()
   If Trim(Txcodigo.Text) <> "" Then
      Set rs = New Recordset
      sql = "SELECT * FROM CpHor" & genaño & "  WHERE (HORARIO='" & Txcodigo.Text & "')"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         rs.Close
         Set rs = New Recordset
         sql = "SELECT * FROM CpHorarioNew  WHERE (codigo='" & Txcodigo.Text & "')"
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then
            MsgBox "<Error> Ya existe..", vbExclamation
            rs.Close
            Exit Sub
         End If
      Else
         MsgBox "<Error> Ya existe..", vbExclamation
         rs.Close
         Exit Sub
      End If
      rs.Close
      Txdescripcion.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Txdescripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Txdescripcion.Text = UCase(Txdescripcion.Text)
      Check1.SetFocus
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
        Txturno.Text = Right(CboTurno.List(j), nlen - nposicion)
     End If
     Exit Sub
  End If
End Sub
Private Sub TxBuscar_Change()
    Dim cBuscar As String
    cpuntero = DataGrid1.Bookmark
    If Txbuscar.Text = "" Then Exit Sub
    cBuscar = "CODIGO like '%" & Txbuscar.Text & "%'"
    DD.MoveFirst
    DD.Find (cBuscar)
    If DD.EOF Then
       Txbuscar.BackColor = &HFF&
    Else
       Txbuscar.BackColor = &H80000005
       DataGrid1.Refresh
    End If
End Sub
Private Sub txBuscar1_Change()
    Dim cBuscar As String
    cpuntero = DataGrid1.Bookmark
    If Txbuscar1.Text = "" Then Exit Sub
    cBuscar = "DESCRIPCION like '%" & Txbuscar1.Text & "%'"
    DD.MoveFirst
    DD.Find (cBuscar1)
    If DD.EOF Then
       Txbuscar1.BackColor = &HFF&
    Else
       Txbuscar1.BackColor = &H80000005
       DataGrid1.Refresh
    End If
End Sub


