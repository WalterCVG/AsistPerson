VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tCadena 
   Caption         =   "Tabla Cadena"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "tCadena.frx":0000
   ScaleHeight     =   8415
   ScaleMode       =   0  'User
   ScaleWidth      =   11057.53
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxPeriodo 
      Height          =   285
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   29
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton CmdWorkers 
      Caption         =   "Trabajadores"
      Height          =   855
      Left            =   3960
      Picture         =   "tCadena.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Incorporar"
      Top             =   7440
      Width           =   1095
   End
   Begin VB.TextBox txBuscar 
      Height          =   285
      Left            =   4800
      TabIndex        =   5
      Top             =   4440
      Width           =   4935
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   9720
      Picture         =   "tCadena.frx":806A8
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cerrar"
      Top             =   7680
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   600
      TabIndex        =   2
      Top             =   5040
      Width           =   8055
      Begin VB.TextBox Txfinalidad 
         Height          =   285
         Left            =   5280
         MaxLength       =   7
         TabIndex        =   38
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Txmeta 
         Height          =   285
         Left            =   4800
         MaxLength       =   5
         TabIndex        =   34
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Txgrupo_func 
         Height          =   285
         Left            =   4320
         MaxLength       =   4
         TabIndex        =   33
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox TxDivision_func 
         Height          =   285
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   32
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TxFuncion 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   30
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox TxAno 
         Height          =   285
         Left            =   7320
         MaxLength       =   4
         TabIndex        =   13
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Txact_ai_obra 
         Height          =   285
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txProd_pry 
         Height          =   285
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Txprograma 
         Height          =   285
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Txsec_func 
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Txejecutora 
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TxPliego 
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Txcadena 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   52
         TabIndex        =   15
         Top             =   1200
         Width           =   5535
      End
      Begin VB.TextBox Txdescripcion 
         Height          =   285
         Left            =   1320
         MaxLength       =   150
         TabIndex        =   16
         Top             =   1800
         Width           =   6615
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Finalid"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5400
         TabIndex        =   39
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Mta"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4800
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "G.f"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4440
         TabIndex        =   36
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label13 
         Caption         =   "D.f."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4080
         TabIndex        =   35
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label12 
         Caption         =   "Func"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3600
         TabIndex        =   31
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "AÑO"
         Height          =   255
         Left            =   6960
         TabIndex        =   26
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Ac_obra"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Pr_pry"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prog"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "S_Fun"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Ejec"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Pliego"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   9000
      TabIndex        =   1
      Top             =   5160
      Width           =   1575
      Begin VB.CommandButton CmdEliminar 
         Caption         =   "Eli&minar"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   10575
      _ExtentX        =   18653
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      Height          =   255
      Left            =   480
      TabIndex        =   28
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Denominacion a ubicar :"
      ForeColor       =   &H80000001&
      Height          =   195
      Left            =   1920
      TabIndex        =   12
      Top             =   4440
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   8355
   End
End
Attribute VB_Name = "tCadena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean

Private Sub CmdWorkers_Click()
  If Trim(Txcadena.Text) = "" Then Exit Sub
  Load Workers
  Workers.Left = 1860: Workers.Top = 2325
  Workers.Refresh
  Workers.Show
End Sub

Private Sub Form_Load()
    TxPeriodo.Text = genaño
    Set DD = New Recordset
    With DD
     .Fields.Append "ano", adVarChar, 4
     .Fields.Append "cadena", adVarChar, 52
     .Fields.Append "Pliego", adVarChar, 6
     .Fields.Append "Ejecutora", adVarChar, 3
     .Fields.Append "Sec_func", adVarChar, 4
     .Fields.Append "Programa", adVarChar, 4
     .Fields.Append "Prod_pry", adVarChar, 7
     .Fields.Append "Act_ai_obra", adVarChar, 7
     .Fields.Append "Funcion", adVarChar, 2
     .Fields.Append "Division_Func", adVarChar, 3
     .Fields.Append "Grupo_Func", adVarChar, 4
     .Fields.Append "Meta", adVarChar, 5
     .Fields.Append "Finalidad", adVarChar, 7
     .Fields.Append "descripcion", adVarChar, 150
     .Open
    End With
    Set DataGrid1.DataSource = DD
    'Configura columnas
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
          .Caption = "AÑO"
          .Width = Me.TextWidth(String(4, "S"))
         Case 1
          .Caption = "Cad. Hep"
          .Width = Me.TextWidth(String(52, "S"))
         Case 2
         '.Caption =
          .Width = Me.TextWidth(String(5, "S"))
         Case 3
         '.Caption =
          .Width = Me.TextWidth(String(3, "S"))
         Case 13
          .Caption = "Descripcion"
          .Width = Me.TextWidth(String(150, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    
    Set rs = New Recordset
    sql = "SELECT * FROM Cpcadena where ano='" & TxPeriodo.Text + "'"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!ano = rs!ano
       DD!cadena = rs!cadena
       DD!Pliego = rs!Pliego
       DD!Ejecutora = rs!Ejecutora
       DD!Sec_func = rs!Sec_func
       DD!Programa = rs!Programa
       DD!Prod_pry = rs!Prod_pry
       DD!Act_ai_obra = rs!Act_ai_obra
       DD!Funcion = rs!Funcion
       DD!Division_Func = rs!Division_Func
       DD!Grupo_Func = rs!Grupo_Func
       DD!Meta = rs!Meta
       DD!Finalidad = rs!Finalidad
       DD!descripcion = rs!descripcion
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Txano = DataGrid1.Columns(0)
   Txcadena = DataGrid1.Columns(1)
   TxPliego = DataGrid1.Columns(2)
   Txejecutora = DataGrid1.Columns(3)
   Txsec_func = DataGrid1.Columns(4)
   Txprograma = DataGrid1.Columns(5)
   txProd_pry = DataGrid1.Columns(6)
   Txact_ai_obra = DataGrid1.Columns(7)
   TxFuncion = DataGrid1.Columns(8)
   TxDivision_func = DataGrid1.Columns(9)
   Txgrupo_func = DataGrid1.Columns(10)
   Txmeta = DataGrid1.Columns(11)
   Txfinalidad = DataGrid1.Columns(12)
   Txdescripcion = DataGrid1.Columns(13)
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
  Limpiacampos
End Sub
Private Sub cmdNuevo_Click()
    If wrights = "CON" Then
       MsgBox "Usuario No Autorizado", vbExclamation
       Exit Sub
    End If
  If cmdNuevo.Caption = "&Nuevo" Then
    Limpiacampos
    habilitatextos
    TxPliego.SetFocus
    CmdEliminar.Enabled = False
    Sw = True
    cmdNuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
    CmdEditar.Enabled = True
  Else                  'Grabar
    On Error GoTo err_trans
    Txcadena.Text = TxPliego + Txejecutora + Txsec_func + Txprograma + txProd_pry + Txac_ai_obra
    Txcadena = Txcadena + TxFuncion + TxDivision_func + Txgrupo_func + Txmeta + Txfinalidad
    If Trim(Txcadena.Text) = "" Then
       MsgBox ("Favor Verifique Informacion")
       Exit Sub
    End If
    Cnx.BeginTrans
    Dim sql As String
    If Sw Then
       sql = "INSERT INTO Cpcadena(ano,cadena,pliego,ejecutora,sec_func,programa,prod_pry,Act_ai_obra,funcion,division_func,grupo_func,meta,finalidad,descripcion,usuario,fecha) "
       sql = sql + " VALUES('" + Txano + "','" + Txcadena + "','" + TxPliego + "','" + Txejecutora + "','" + Txsec_func + "','" + Txprograma + "','" + txProd_pry + "','" + Txact_ai_obra + "','" & _
       TxFuncion + "','" + TxDivision_func + "','" + Txgrupo_func + "','" + Txmeta + "','" + Txdescripcion + "','" + Txfinalidad + "','" + wUsuario + "','" + Format(Date, "dd/mm/yyyy") + "')"
    Else
       sql = "UPDATE Cpcadena SET pliego='" & TxPliego + "',ejecutora='" + Txejecutora + "',set_func='" + Txsec_func + "',programa='" + Txprograma + "',prod_pry='" + txProd_pry + "', " & _
       "act_ai_obra='" + Txact_ai_obra + "', funcion='" & TxFuncion + "', division_func='" + TxDivision_func + "', grupo_func='" + Txgrupo_func + "', meta='" + Txmeta + "', " & _
       "Finalidad='" + Txfinalidad + "',descripcion='" + Txdescripcion.Text + "',usuario='" + wUsuario + "',fecha='" + Format(Date, "dd/mm/yyyy") + "'"
       sql = sql + " WHERE ano='" + Txano.Text + "' and cadena='" + Txcadena + "'"
    End If
    Cnx.Execute sql
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!ano = Txano.Text
       DD!cadena = Txcadena.Text
    End If
    DD!Pliego = TxPliego
    DD!Ejecutora = Txejecutora
    DD!Sec_func = Txsec_func
    DD!Programa = Txprograma
    DD!Prod_pry = txProd_pry
    DD!Act_ai_obra = Txact_ai_obra
    DD!Funcion = TxFuncion
    DD!Division_Func = TxDivision_func
    DD!Grupo_Func = Txgrupo_func
    DD!Meta = Txmeta
    DD!Finalidad = Txfinalidad
    DD!descripcion = Txdescripcion
    DD.Update
    desabilitatextos
    Limpiacampos
    CmdEliminar.Enabled = True
    cmdNuevo.Caption = "&Nuevo"
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
    Txano.Locked = True
    'TxCadena.Locked = True
    TxPliego.SetFocus
    CmdEliminar.Enabled = False
    Sw = False
    cmdNuevo.Caption = "&Grabar"
    CmdEditar.Caption = "&Cancelar"
  Else                   'Cancelar
    desabilitatextos
    Limpiacampos
    CmdEliminar.Enabled = True
    cmdNuevo.Caption = "&Nuevo"
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
    sql = "SELECT * FROM Cpdirectorio WHERE CadHep='" & Txcadena.Text & "'"
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If r_d.EOF Then
    Else
       MsgBox "Hay Servidores en dicha Cadena..", vbExclamation
       r_d.Close
       Exit Sub
    End If
    r_d.Close
    
    If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
       Cnx.Execute "DELETE FROM CpCadena WHERE ano='" + Txano.Text + "' and cadena='" & Txcadena.Text & "'"
       DD.Delete
       Limpiacampos
    End If
End Sub
Private Sub cmdsalir_Click()
  Unload Me
  Me.Hide
End Sub
Sub habilitatextos()
  Txano.Locked = False
  'Txcadena.Locked = False
   TxPliego.Locked = False
   Txejecutora.Locked = False
   Txsec_func.Locked = False
   Txprograma.Locked = False
   txProd_pry.Locked = False
   Txact_ai_obra.Locked = False
   TxFuncion.Locked = False
   TxDivision_func.Locked = False
   Txgrupo_func.Locked = False
   Txmeta.Locked = False
   Txfinalidad.Locked = False
   Txdescripcion.Locked = False
End Sub
Sub desabilitatextos()
  Txano.Locked = True
  'Txcadena.Locked = True
   TxPliego.Locked = True
   Txejecutora.Locked = True
   Txsec_func.Locked = True
   Txprograma.Locked = True
   txProd_pry.Locked = True
   Txact_ai_obra.Locked = True
   TxFuncion.Locked = True
   TxDivision_func.Locked = True
   Txgrupo_func.Locked = True
   Txmeta.Locked = True
   Txfinalidad.Locked = True
   Txdescripcion.Locked = True
End Sub
Sub Limpiacampos()
  Txano.Text = "": Txcadena.Text = ""
  TxPliego.Text = "": Txejecutora.Text = "": Txset_func.Text = "": Txprograma.Text = "": txProd_pry.Text = "": Txact_ai_obra.Text = ""
  TxFuncion.Text = "": TxDivision_func.Text = "": Txgrupo_func.Text = "": Txmeta.Text = "": Txfinalidad.Text = "": Txdescripcion.Text = ""
End Sub

Private Sub TxBuscar_Change()
    Dim cBuscar As String
    cpuntero = DataGrid1.Bookmark
    If Txbuscar.Text = "" Then Exit Sub
    cBuscar = "Descripcion like '%" & Txbuscar.Text & "%'"
    DD.MoveFirst
    DD.Find (cBuscar)
    If DD.EOF Then
       Txbuscar.BackColor = &HFF&
    Else
       Txbuscar.BackColor = &H80000005
       DataGrid1.Refresh
    End If
End Sub

Private Sub Txpliego_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txejecutora.SetFocus
     Exit Sub
   End If
End Sub
Private Sub Txdescripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txcadena.Text = TxPliego + " " + Txejecutora + " " + Txsec_func + " " + Txprograma + " " + txProd_pry + " " + Txac_ai_obra
     Txcadena = Txcadena + " " + TxFuncion + " " + TxDivision_func + " " + Txgrupo_func + " " + Txmeta + " " + Txfinalidad
     Txdescripcion.Text = UCase(Txdescripcion.Text)
     cmdNuevo.SetFocus
     Exit Sub
   End If
End Sub




