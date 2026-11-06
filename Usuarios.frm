VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Usuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   Caption         =   "Mantenimiento de Usuarios"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "Usuarios.frx":0000
   ScaleHeight     =   7425
   ScaleMode       =   0  'User
   ScaleWidth      =   10730.11
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
      Left            =   9360
      Picture         =   "Usuarios.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4560
      TabIndex        =   24
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton Cmdsalir 
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
      Left            =   9480
      Picture         =   "Usuarios.frx":80EA8
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cerrar"
      Top             =   6840
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   360
      TabIndex        =   22
      Top             =   240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5318
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
      Caption         =   "USUARIOS"
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
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
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
      Left            =   9000
      TabIndex        =   14
      Top             =   4560
      Width           =   1575
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
         TabIndex        =   19
         Top             =   1440
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
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
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
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[DETALLES ]"
      ForeColor       =   &H000000C0&
      Height          =   3015
      Left            =   840
      TabIndex        =   0
      Top             =   3840
      Width           =   7695
      Begin VB.CheckBox ChkBAJA 
         Alignment       =   1  'Right Justify
         Caption         =   "BAJA"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox ChkDerEspec 
         Alignment       =   1  'Right Justify
         Caption         =   "Derechos Especiales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox Txcambia 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6000
         MaxLength       =   2
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CboCambia 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdAccesoS 
         Caption         =   "ACCESOS"
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
         Left            =   4200
         TabIndex        =   13
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Txacceso1 
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
         Left            =   6000
         TabIndex        =   31
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Txacceso2 
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
         Left            =   6000
         TabIndex        =   30
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Txacceso3 
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
         Left            =   6000
         TabIndex        =   29
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Txacceso4 
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
         Left            =   6000
         TabIndex        =   28
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Txacceso5 
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
         Left            =   6000
         TabIndex        =   27
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Txacceso6 
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
         Left            =   6000
         TabIndex        =   26
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Txacceso7 
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
         Left            =   6000
         TabIndex        =   25
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox TxOKPwd 
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
         IMEMode         =   3  'DISABLE
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox CboModulo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox TxModulo 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   960
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox TxUsuario 
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
         Left            =   960
         MaxLength       =   10
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox Cboderechos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Txnombre 
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
         Left            =   960
         MaxLength       =   40
         TabIndex        =   8
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox Txderechos 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   960
         MaxLength       =   3
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Txclave 
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
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Cambia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Modulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Derechos"
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
         Left            =   120
         TabIndex        =   2
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DD As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim Sw As Boolean
Dim cDetaDel As String
Dim marcador As Variant

Private Sub Command1_Click()
     c_grupo = "05": c_secuencia = "053"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub
Private Sub Form_Load()
'   Txclave.PasswordChar = "&"
'   TxOKPwd.PasswordChar = "&"
'  If wUsuario = "AMASTER" Then
'     Txclave.PasswordChar = ""
'  End If
'  If wacceso5 = "XXXXXXXXXXXXXXX" Then
'     Txclave.ForeColor = &H80000002
'     TxOKPwd.ForeColor = &H80000002
'  End If
  Set DD = New Recordset
   With DD
     .Fields.Append "usuario", adChar, 10
     .Fields.Append "Clave", adChar, 10
     .Fields.Append "cambia", adChar, 2
     .Fields.Append "nombre", adChar, 40
     .Fields.Append "modulo", adChar, 3
     .Fields.Append "derechos", adChar, 3
     .Fields.Append "DerEspec", adChar, 1
     .Fields.Append "BAJA", adChar, 1
     .Fields.Append "NickDel", adChar, 10
     .Fields.Append "fechaDel", adDate
     .Open
  End With
  Set DataGrid1.DataSource = DD
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
        Select Case i
         Case 0
          .Caption = "Codigo"
          .Width = Me.TextWidth(String(10, "S"))
         Case 1
          .Caption = "CLave"
          .Width = Me.TextWidth(String(10, "S"))
         Case 2
          .Caption = "Cambia"
          .Width = Me.TextWidth(String(6, "S"))
         Case 3
          .Caption = "Nombre"
          .Width = Me.TextWidth(String(30, "S"))
     '    Case 4
     '     .Caption = "modulo"
     '     .Width = Me.TextWidth(String(6, "S"))
     '    Case 5
     '     .Caption = "Derechos"
     '     .Width = Me.TextWidth(String(8, "S"))
     '    Case 6
     '     .Caption = "Derechos"
     '     .Width = Me.TextWidth(String(8, "S"))
         Case 7
          .Caption = "BAJA"
          .Width = Me.TextWidth(String(5, "S"))
         Case Else
          .Visible = False
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM CPusuarios "
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!usuario = rs!usuario
       DD!Clave = rs!Clave
       DD!cambia = IIf(IsNull(rs!cambia), "", rs!cambia)
       DD!nombre = Trim(rs!nombre)
       DD!modulo = IIf(IsNull(rs!modulo), "", rs!modulo)
       DD!derechos = IIf(IsNull(rs!derechos), "", rs!derechos)
       DD!DerEspec = IIf(IsNull(rs!DerEspec), "", rs!DerEspec)
       DD!BAJA = IIf(IsNull(rs!BAJA), "", rs!BAJA)
       DD!nickdel = IIf(IsNull(rs!nickdel), "", rs!nickdel)
       If IsDate(rs!fechadel) Then DD!fechadel = rs!fechadel
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
    'LLena el Combo Modulo
    CboModulo.Clear
    CboModulo.AddItem "SISTEMAS ": CboModulo.ItemData(CboModulo.NewIndex) = 1
    CboModulo.AddItem "OPERADOR": CboModulo.ItemData(CboModulo.NewIndex) = 2
    CboModulo.ListIndex = 0
    'LLena el Combo derechos
    Cboderechos.Clear
    Cboderechos.AddItem "CONSULTAS": Cboderechos.ItemData(Cboderechos.NewIndex) = 1
    Cboderechos.AddItem "EDICION": Cboderechos.ItemData(Cboderechos.NewIndex) = 2
    Cboderechos.AddItem "MASTER": Cboderechos.ItemData(Cboderechos.NewIndex) = 3
    Cboderechos.ListIndex = 0
    CboCambia.Clear
    CboCambia.AddItem "SI": CboCambia.ItemData(CboCambia.NewIndex) = 1
    CboCambia.AddItem "NO": Cboderechos.ItemData(CboCambia.NewIndex) = 2
    CboCambia.ListIndex = 0
    Call Limpiacampos
End Sub
Private Sub CmdAccesos_Click()
   If Trim(Txusuario.Text) = "" Then Exit Sub
   Load UsuarioAcceso
   UsuarioAcceso.Left = 1500: UsuarioAcceso.Top = 1300
   UsuarioAcceso.Refresh
   UsuarioAcceso.Show
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Txusuario = DataGrid1.Columns(0)
   Txclave = DataGrid1.Columns(1):
   TxOKPwd = DataGrid1.Columns(1)
   Txcambia = DataGrid1.Columns(2)
   Txnombre = DataGrid1.Columns(3)
   TxModulo = DataGrid1.Columns(4)
   Txderechos = DataGrid1.Columns(5)
   ChkDerEspec.Value = IIf(DataGrid1.Columns(6) = "S", 1, 0)
   ChkBAJA.Value = IIf(DataGrid1.Columns(7) = "S", 1, 0)
   cDetaDel = ""
   If DataGrid1.Columns(7) = "S" Then
      cDetaDel = " el " + Format(DataGrid1.Columns(9), "dd/mm/yyyy") + " por " + DataGrid1.Columns(8) + " "
   End If
   Frame1.ForeColor = IIf(DataGrid1.Columns(7) = "S", &HC0&, &H80000012)
   Frame1.Caption = IIf(DataGrid1.Columns(7) = "S", "[ USUARIO DE BAJA" + cDetaDel + "]", "[ DETALLES ]")
   '&H80000012&     negro
   '&H000000C0&   rojo
   
   Select Case Txcambia.Text
    Case "SI"
     CboCambia.ListIndex = 0
    Case "NO"
     CboCambia.ListIndex = 1
   End Select
   Select Case TxModulo.Text
    Case "SIS"
     CboModulo.ListIndex = 0
    Case "OPE"
     CboModulo.ListIndex = 1
   End Select
   Select Case Txderechos.Text
    Case "CON"
     Cboderechos.ListIndex = 0
    Case "EDI"
     Cboderechos.ListIndex = 1
    Case "MAS"
     Cboderechos.ListIndex = 2
   End Select
   Dim rs As Recordset
   Set rs = New ADODB.Recordset
   sql = "Select * From CPUsuarios where  usuario='" & Usuarios.Txusuario.Text & "'"
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If rs.EOF Then
      Txacceso1 = "000000000000000"
      Txacceso2 = "000000000000000"
      Txacceso3 = "000000000000000"
      Txacceso4 = "000000000000000"
      Txacceso5 = "000000000000000"
      Txacceso6 = "000000000000000"
      Txacceso7 = "000000000000000"
   Else
      Txacceso1 = IIf(IsNull(rs!acceso1), "" + "000000000000000", Left(rs!acceso1 + "000000000000000", 15))
      Txacceso2 = IIf(IsNull(rs!acceso2), "" + "000000000000000", Left(rs!acceso2 + "000000000000000", 15))
      Txacceso3 = IIf(IsNull(rs!acceso3), "" + "000000000000000", Left(rs!acceso3 + "000000000000000", 15))
      Txacceso4 = IIf(IsNull(rs!acceso4), "" + "000000000000000", Left(rs!acceso4 + "000000000000000", 15))
      Txacceso5 = IIf(IsNull(rs!acceso5), "" + "000000000000000", Left(rs!acceso5 + "000000000000000", 15))
      Txacceso6 = IIf(IsNull(rs!acceso6), "" + "000000000000000", Left(rs!acceso6 + "000000000000000", 15))
      Txacceso7 = IIf(IsNull(rs!acceso7), "" + "000000000000000", Left(rs!acceso7 + "000000000000000", 15))
   End If
   Encclave = ""
   WVdesencripta (rs!Clave)
   rs.Close
   Txclave = Encclave
   TxOKPwd = Encclave
   Cmdnuevo.Caption = "&Nuevo"
   CmdEditar.Caption = "&Editar"
   CmdEditar.Enabled = True
   CmdEliminar.Enabled = True
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
  With DD      ' Adodc1.Recordset
    If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
       .Sort = .Fields(ColIndex).[Name] & " Desc"
    Else
       .Sort = .Fields(ColIndex).[Name] & " Asc"
    End If
  End With
  Call Limpiacampos
End Sub

Private Sub cmdNuevo_Click()
  If wacceso Or wrights = "MAS" Then
  Else
     MsgBox "Usuario No Autorizado", vbExclamation
     Exit Sub
   End If
  If Cmdnuevo.Caption = "&Nuevo" Then
    Limpiacampos
    habilitatextos
    Txusuario.SetFocus
    CmdEliminar.Enabled = False
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
    Exit Sub
  End If
  '-Grabar
  If Txclave = UCase(TxOKPwd) Then
     Encclave = ""
     WVencripta (Txclave)
     Text1 = Encclave
  Else
     MsgBox "Claves No Coinciden..  ", vbExclamation
     TxOKPwd = ""
     Exit Sub
  End If
  If TxModulo = "" Then
     MsgBox "Defina modulo...  ", vbExclamation
     Exit Sub
  End If
  If Txderechos = "" Then
     MsgBox "Defina derechos...  ", vbExclamation
     Exit Sub
  End If
'''  On Error GoTo err_trans
    
    Cnx.BeginTrans
    Dim sql As String
    
      sql = "SELECT * FROM CPusuarios  WHERE usuario='" + Txusuario + "'"
      Set rs = New Recordset
      rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      If rs.EOF Then
         rs.AddNew
         rs!usuario = Txusuario
         rs!nick = wUsuario
         rs!fecha = Date
         Sw = True
      End If
      cbaja = IIf(IsNull(rs!BAJA), "", Trim(rs!BAJA))
      rs!Clave = Text1
      rs!cambia = Txcambia
      rs!nombre = Txnombre
      rs!modulo = TxModulo
      rs!derechos = Txderechos
      rs!DerEspec = IIf(ChkDerEspec.Value = 1, "S", "")
      rs!acceso1 = Txacceso1
      rs!acceso2 = Txacceso2
      rs!acceso3 = Txacceso3
      rs!acceso4 = Txacceso4
      rs!acceso5 = Txacceso5
      rs!acceso6 = Txacceso6
      rs!acceso7 = Txacceso7
      rs!nickmdf = wUsuario
      rs!fechamdf = Date
      If ChkBAJA.Value = 1 And cbaja <> "S" Then
         '-Da BAJA, actualiza detalles baja
         rs!BAJA = "S"
         rs!nickdel = wUsuario
         rs!fechadel = Date
      End If
      If ChkBAJA.Value = 0 Then
         '-Reversa BAJA, deja fecha detalles baja anterior
         rs!BAJA = ""
         'rs!nickdel = wUsuario
         'rs!fechadel = Date
      End If
      
      
      rs.Update
      rs.Close
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!usuario = Txusuario
    End If
    DD!Clave = Text1
    DD!cambia = Txcambia
    DD!nombre = Trim(Txnombre)
    DD!modulo = TxModulo
    DD!derechos = Txderechos
    DD!DerEspec = IIf(ChkDerEspec.Value = 1, "S", "")
    DD!BAJA = IIf(ChkBAJA.Value = 1, "S", "")
    If ChkBAJA.Value = 1 Then
       DD!nickdel = wUsuario
       DD!fechadel = Date
    End If
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
'''err_trans:
   ''' MsgBox Err.Description
  
End Sub
Private Sub Cmdeditar_Click()
  If wacceso Or wrights = "MAS" Then
  Else
     MsgBox "Usuario No Autorizado", vbExclamation
     Exit Sub
  End If
  If CmdEditar.Caption = "&Editar" Then
    habilitatextos
    Txusuario.Locked = True
    Txclave.SetFocus
    CmdEliminar.Enabled = False
    Sw = False
    Cmdnuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
  Else                   'Cancelar
    desabilitatextos
    CmdEliminar.Enabled = True
    Cmdnuevo.Caption = "&Nuevo"
    CmdEditar.Caption = "&Editar"
    CmdEditar.Enabled = False
    CmdEliminar.Enabled = False
    Limpiacampos
    DataGrid1.SetFocus
  End If
End Sub
Private Sub CmdEliminar_Click()
  If wacceso Or wrights = "MAS" Then
  Else
     MsgBox "Usuario No Autorizado", vbExclamation
     Exit Sub
  End If
  If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
     Set rs = New Recordset
     sql = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and (usuario ='" + Txusuario.Text + "')"
     rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
     If Not rs.EOF Then
        MsgBox "Usuario tiene Registros. ", vbExclamation
        rs.Close
        Exit Sub
     End If
     rs.Close
     'Borra en tabla
     Cnx.Execute "DELETE FROM cpusuarios WHERE Usuario='" & Txusuario.Text & "'"
     'Borra en Temporal
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
  Txusuario.Locked = False
  Txclave.Locked = False
  Txnombre.Locked = False
  CboCambia.Enabled = True
  CboModulo.Enabled = True
  Cboderechos.Enabled = True
  ChkDerEspec.Enabled = True
  ChkBAJA.Enabled = True
End Sub
Sub desabilitatextos()
  Txusuario.Locked = True
  Txclave.Locked = True
  Txnombre.Locked = True
  CboCambia.Enabled = False
  CboModulo.Enabled = False
  Cboderechos.Enabled = False
  ChkDerEspec.Enabled = False
  ChkBAJA.Enabled = False
End Sub
Sub Limpiacampos()
    Frame1.ForeColor = &H80000012
    Frame1.Caption = "[ DETALLES ]"
    Txusuario.Text = ""
    Txclave.Text = ""
    TxOKPwd.Text = ""
    Text1.Text = ""
    Txnombre.Text = ""
    Txcambia.Text = "":    CboCambia.Text = ""
    CboModulo.Text = "":    TxModulo.Text = ""
    Cboderechos.Text = "":    Txderechos.Text = ""
    ChkDerEspec.Value = 0
    ChkBAJA.Value = 0
    Txacceso1.Text = "000000000000000"
    Txacceso2.Text = "000000000000000"
    Txacceso3.Text = "000000000000000"
    Txacceso4.Text = "000000000000000"
    Txacceso5.Text = "000000000000000"
    Txacceso6.Text = "000000000000000"
    Txacceso7.Text = "000000000000000"
    
End Sub

Private Sub Txclave_LostFocus()
     Txclave.Text = UCase(Txclave.Text)
End Sub
Private Sub Txusuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txusuario.Text = UCase(Txusuario.Text)
     Txclave.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Txclave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txclave.Text = UCase(Txclave.Text)
     Encclave = ""
     WVencripta (Txclave.Text)
     Text1.Text = Encclave
     TxOKPwd = ""
     TxOKPwd.SetFocus
     Exit Sub
   End If
End Sub
Private Sub TxOkPwd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     If Txclave.Text = UCase(TxOKPwd) Then
        Txnombre.SetFocus
        Exit Sub
     Else
       MsgBox "Claves No Coinciden..  ", vbExclamation
       TxOKPwd = ""
     End If
   End If
End Sub
Private Sub Txnombre_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txnombre.Text = UCase(Txnombre.Text)
     CboModulo.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Cbocambia_Click()
   Txcambia = Left(CboCambia.Text, 2)
End Sub
Private Sub Cbomodulo_Click()
   TxModulo.Text = Left(CboModulo.Text, 3)
End Sub
Private Sub Cboderechos_Click()
   Txderechos.Text = Left(Cboderechos.Text, 3)
End Sub




