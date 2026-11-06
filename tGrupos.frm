VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tGrupos 
   Caption         =   "Grupos de Trabajadores"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "tGrupos.frx":0000
   ScaleHeight     =   8970
   ScaleWidth      =   12060
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txbuscar1 
      Height          =   285
      Left            =   5160
      MaxLength       =   35
      TabIndex        =   18
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox Txbuscar 
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   10080
      Picture         =   "tGrupos.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cerrar"
      Top             =   8280
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   9600
      TabIndex        =   10
      Top             =   5520
      Width           =   1575
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
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
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   1080
      TabIndex        =   1
      Top             =   5520
      Width           =   8055
      Begin VB.TextBox Txsabado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7560
         MaxLength       =   1
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox ChkSabado 
         Caption         =   "Sabados"
         Height          =   195
         Left            =   6480
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox Cbohorario 
         Height          =   315
         Left            =   1440
         TabIndex        =   22
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox Txhorario 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Txtrabajador 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   35
         TabIndex        =   6
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox Txtarjeta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Txgeneral 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         MaxLength       =   1
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox ChkGeneral 
         Caption         =   "Generico"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5760
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox Txdescripcion 
         Height          =   285
         Left            =   1200
         MaxLength       =   35
         TabIndex        =   7
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox Txgrupo 
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Horario"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Grupo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   9015
      _ExtentX        =   15901
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ubica desc"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Ubica Grupo"
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   4800
      Width           =   975
   End
End
Attribute VB_Name = "tGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean
Private Sub Form_Load()
    Set rs = New Recordset
    rs.Open "SELECT * FROM CpHorarioNew  ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
    rs.MoveFirst
    While Not rs.EOF
      For i = 0 To CboHorario.ListCount - 1
          If CboHorario.List(i) = Trim(rs.Fields("Descripcion")) & " (" & rs.Fields("horas") & ")  =>" & rs.Fields("codigo") Then
             ExisteYa = True
             Exit For
          End If
      Next i
      If Not ExisteYa Then
         CboHorario.AddItem Trim(rs.Fields("Descripcion")) & " (" & rs.Fields("horas") & ")  =>" & rs.Fields("codigo")
      End If
      rs.MoveNext
    Wend
    rs.Close
    ChkGeneral.Value = 0
    ChkSabado.Value = 0
    Set DD = New Recordset
    With DD
     .Fields.Append "grupo", adVarChar, 4
     .Fields.Append "general", adVarChar, 1
     .Fields.Append "horario", adVarChar, 4
     .Fields.Append "sabado", adVarChar, 1
     .Fields.Append "descripcion", adVarChar, 35
     .Fields.Append "tarjeta", adVarChar, 5
     .Open
    End With
    Set DataGrid1.DataSource = DD
    'Configura columnas
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         '.Caption = "Codigo"
          .Width = Me.TextWidth(String(6, "S"))
         Case 1
          '.Caption = "general"
          .Width = Me.TextWidth(String(6, "S"))
         Case 2
          '.Caption = "horario"
          .Width = Me.TextWidth(String(6, "S"))
         Case 3
          '.Caption = "sabado"
          .Width = Me.TextWidth(String(6, "S"))
         Case 4
          '.Caption = "Descripcion"
          .Width = Me.TextWidth(String(35, "S"))
         Case 5
          '.Caption = "Tarjeta"
          .Width = Me.TextWidth(String(8, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM CPgrupos "
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
      rs.MoveFirst
      Do While Not rs.EOF
         DD.AddNew
         DD!grupo = rs!grupo
         DD!General = IIf(IsNull(rs!General), "N", rs!General)
         DD!Horario = IIf(IsNull(rs!Horario), "", rs!Horario)
         DD!sabado = IIf(IsNull(rs!sabado), "N", rs!sabado)
         DD!descripcion = IIf(IsNull(rs!descripcion), "", rs!descripcion)
         DD!tarjeta = IIf(IsNull(rs!tarjeta), "", rs!tarjeta)
         DD.Update
         rs.MoveNext
      Loop
    End If
    rs.Close
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Sw = False
   Call Limpiacampos
   Txgrupo.Text = DataGrid1.Columns(0)
   Txgeneral.Text = DataGrid1.Columns(1)
   TxHorario.Text = DataGrid1.Columns(2)
   Txsabado.Text = DataGrid1.Columns(3)
   Txdescripcion.Text = DataGrid1.Columns(4)
   Txtarjeta.Text = DataGrid1.Columns(5)
   If Txgeneral = "S" Then
      ChkGeneral.Value = 1
   Else
      ChkGeneral.Value = 0
   End If
   If Txsabado = "S" Then
      ChkSabado.Value = 1
   Else
      ChkSabado.Value = 0
   End If
   If Txtarjeta.Text <> "" Then
      Set rs = New Recordset
      sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and tarjeta = '" & Txtarjeta.Text & "'"
      rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "No Ubicado", vbExclamation
         Txtarjeta.Text = ""
      Else
         TxTrabajador.Text = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
      End If
      rs.Close
   End If
   If Trim(TxHorario.Text) <> "" Then
      For i = 0 To CboHorario.ListCount - 1
         npos = InStr(CboHorario.List(i), ">")
         nlen = Len(CboHorario.List(i))
         dato = Trim(Right(CboHorario.List(i), nlen - npos))
         If dato = Trim(TxHorario.Text) Then
            CboHorario.ListIndex = i
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
    Txgrupo.SetFocus
    CmdEliminar.Enabled = False
    Sw = True
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
  Else                  'Grabar
    On Error GoTo err_trans
    Txgeneral.Text = IIf(ChkGeneral.Value = 1, "S", "N")
    Txsabado.Text = IIf(ChkSabado.Value = 1, "S", "N")
    If Txgeneral.Text = "S" Then
       Txtarjeta.Text = ""
    End If
    Cnx.BeginTrans
    Dim sql As String
    If Sw Then
       sql = "INSERT INTO Cpgrupos(grupo,general,horario,sabado,descripcion,tarjeta,usuario,Fecha) "
       sql = sql + " VALUES('" + Txgrupo.Text + "','" + Txgeneral.Text + "','" + TxHorario.Text + "','" + Txsabado.Text + "','" + Txdescripcion.Text + "','" + Txtarjeta.Text + "','" + wUsuario + "','" + Format(Date, "yyyy-mm-dd") + "')"
    Else
       sql = "UPDATE CPgrupos SET general='" + Txgeneral.Text + "', horario='" + TxHorario.Text + "', sabado='" + Txsabado.Text + "', descripcion='" + Txdescripcion.Text + "', tarjeta='" + Txtarjeta.Text + "',usuario='" + wUsuario + "',fecha='" + Format(Date, "yyyy-mm-dd") + "'"
       sql = sql + " WHERE grupo='" + Txgrupo.Text + "'"
    End If
    Cnx.Execute sql
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!grupo = Txgrupo.Text
    End If
    DD!General = Txgeneral.Text
    DD!Horario = TxHorario.Text
    DD!sabado = Txsabado.Text
    DD!descripcion = Txdescripcion.Text
    DD!tarjeta = Txtarjeta.Text
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
   sql = "SELECT * FROM CpGeneral  WHERE  (baja is NULL) and (grpgen='" & Trim(Txgrupo.Text) & "')"
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If rs.EOF Then
   Else
      MsgBox "Hay Tarjetas con dicho Grupo generico ", vbExclamation
      rs.Close
      Exit Sub
   End If
   rs.Close
   If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
      Cnx.Execute "DELETE FROM CPgrupos WHERE grupo='" & Txgrupo.Text & "'"
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
   Txgrupo.Locked = False
   ChkGeneral.Enabled = True
   TxHorario.Locked = False
   ChkSabado.Enabled = True
   Txdescripcion.Locked = False
   Txtarjeta.Locked = False
End Sub
Sub desabilitatextos()
   Txgrupo.Locked = True
   ChkGeneral.Enabled = False
   TxHorario.Locked = True
   ChkSabado.Enabled = False
   Txdescripcion.Locked = True
   Txtarjeta.Locked = True
End Sub
Sub Limpiacampos()
   Txgrupo.Text = "": ChkGeneral.Value = 0: Txgeneral.Text = ""
   Txdescripcion.Text = "": TxHorario.Text = "": CboHorario.Text = ""
   ChkSabado.Value = 0: Txsabado.Text = ""
    Txtarjeta.Text = "": TxTrabajador.Text = ""
End Sub
Private Sub TxBUscar_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Limpiacampos
     Dim cBuscar As String
     cpuntero = DataGrid1.Bookmark
     If Txbuscar.Text = "" Then Exit Sub
     cBuscar = "grupo like '" & Txbuscar.Text & "*'"
     DD.MoveFirst
     DD.Find (cBuscar)
     If DD.EOF Then
        Txbuscar.BackColor = &HFF&
     Else
        Txbuscar.BackColor = &H80000005
        Txbuscar.Text = ""
        DataGrid1.Refresh
        DataGrid1_Click
     End If
     'Limpiacampos
   End If
End Sub
Private Sub Txbuscar1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Limpiacampos
     Dim cBuscar As String
     cpuntero = DataGrid1.Bookmark
     If Txbuscar1.Text = "" Then Exit Sub
     cBuscar = "Descripcion like '" & Trim(Txbuscar1.Text) & "*'"
     DD.MoveFirst
     DD.Find (cBuscar)
     If DD.EOF Then
        Txbuscar1.BackColor = &HFF&
     Else
        Txbuscar1.BackColor = &H80000005
        DataGrid1.Refresh
        DataGrid1_Click
        Txbuscar1.Text = ""
     End If
     'Limpiacampos
   End If
End Sub


Private Sub Txgrupo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Set rs = New Recordset
      sql = "SELECT * FROM Cpgrupos  WHERE (grupo='" & Txgrupo.Text & "')"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
      Else
         MsgBox "Horario Ya existe..", vbExclamation
      End If
      rs.Close
      Exit Sub
   End If
End Sub
Private Sub Txdescripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Txdescripcion.Text = Txdescripcion.Text
      Cmdnuevo.SetFocus
      Exit Sub
   End If
End Sub


Private Sub Cbohorario_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboHorario.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbohorario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If IsEmpty(CboHorario.Text) Then
     Else
        j = CboHorario.ListIndex
        If j <> -1 Then
           nposicion = InStr(CboHorario.List(j), ">")
           nlen = Len(CboHorario.List(j))
           TxHorario.Text = Right(CboHorario.List(j), nlen - nposicion)
        End If
        Exit Sub
     End If
   End If
End Sub
Private Sub Cbohorario_LostFocus()
   If IsEmpty(CboHorario.Text) Then
   Else
      j = CboHorario.ListIndex
      If j <> -1 Then
         nposicion = InStr(CboHorario.List(j), ">")
         nlen = Len(CboHorario.List(j))
         TxHorario.Text = Right(CboHorario.List(j), nlen - nposicion)
      End If
      Exit Sub
   End If
End Sub
Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(Txtarjeta.Text) = "" Or (Trim(Txtarjeta.Text) <> "" And EsTexto(Txtarjeta.Text)) Then
         If Trim(Txtarjeta.Text) <> "" And EsTexto(Txtarjeta.Text) Then
            Ayuda1.TxUbiAP = Trim(Txtarjeta.Text)
         End If
         Load Ayuda1
         Ayuda1.Tag = 0
         Ayuda1.Show 1
      Else
         Set rt = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and tarjeta = '" & Txtarjeta.Text & "'"
         rt.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rt.EOF Then
            MsgBox "No Ubicado", vbExclamation
            Txtarjeta.Text = ""
         Else
            TxTrabajador.Text = Left(Trim(rt!apelpat) + " " + Trim(rt!apelmat) + " " + Trim(rt!nombres), 35)
         End If
         rt.Close
      End If
      Exit Sub
   End If
End Sub
