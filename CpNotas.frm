VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CpNotas 
   Caption         =   "Notas  del  Servidor"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "CpNotas.frx":0000
   ScaleHeight     =   4215
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtarjeta 
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      MaxLength       =   5
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   615
      Left            =   9720
      Picture         =   "CpNotas.frx":0172
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   615
      Left            =   5880
      Picture         =   "CpNotas.frx":02E4
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar"
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "Graba"
      Height          =   615
      Left            =   7560
      Picture         =   "CpNotas.frx":0456
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Grabar"
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   615
      Left            =   4200
      Picture         =   "CpNotas.frx":05C8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Agregar Registro "
      Top             =   3480
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   9615
      Begin VB.TextBox Txdetalle 
         Height          =   495
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   5
         Top             =   840
         Width           =   8175
      End
      Begin VB.TextBox Txfecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Lblresto 
         BackStyle       =   0  'Transparent
         Caption         =   "van.."
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
         Left            =   7680
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle (200posc)"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2778
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
      Caption         =   "Registro"
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
Attribute VB_Name = "CpNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean
Private Sub Form_Load()
    Txtarjeta = General.Txtarjeta
    Set DD = New Recordset
    With DD
      .Fields.Append "Fecha", adDate
      .Fields.Append "Detalle", adVarChar, 100
      .Open
    End With
    Set DataGrid1.DataSource = DD
    'Configura columnas
    For i = 0 To DataGrid1.Columns.Count - 1
      With DataGrid1.Columns(i)
         Select Case i
         Case 0
         '.Caption = "fecha"
          .Width = Me.TextWidth(String(10, "S"))
         Case 1
          '.Caption = "detalle"
          .Width = Me.TextWidth(String(100, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
      End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM CpNotas where tarjeta='" + Txtarjeta + "' order by fecha"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!fecha = rs!fecha
       DD!detalle = Left(rs!detalle, 100)
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
End Sub
Private Sub DataGrid1_Click()
   If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
   Txfecha = DataGrid1.Columns(0)
   fecha = Format(Txfecha, "yyyymmdd")
   Set r_d = New ADODB.Recordset
   sql = "SELECT * FROM CpNotas  WHERE tarjeta='" + Txtarjeta + "' and fecha='" + fecha + "'"
   r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not r_d.EOF Then
      Txdetalle = Trim(DataGrid1.Columns(1))
      cmdElimina.Enabled = True
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
  Call Limpiacampos
End Sub
Private Sub cmdNuevo_Click()
    If wrights = "CON" Then
       MsgBox "Usuario No Autorizado", vbExclamation
       Exit Sub
    End If
    Limpiacampos
    Txfecha = Date
    Txfecha.SetFocus
    cmdElimina.Enabled = False
    Sw = True
End Sub
Private Sub CmdElimina_Click()
   If wrights = "CON" Then
      MsgBox "Usuario No Autorizado", vbExclamation
      Exit Sub
   End If
   fecha = Format(Txfecha, "yyyymmdd")
   Set r_d = New ADODB.Recordset
   sql = "SELECT * FROM CpNotas  WHERE tarjeta='" + Txtarjeta + "' and fecha='" + fecha + "'"
   r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If MsgBox("Elimina registro ", vbOKCancel + vbQuestion, "Confirme") <> vbCancel Then
      Cnx.Execute "DELETE FROM CpNotas WHERE tarjeta='" + Txtarjeta + "' and fecha='" + fecha + "'"
      DD.Delete
      'DataGrid1.Refresh
      Limpiacampos
   End If
End Sub
Private Sub cmdGraba_Click()
    If wrights = "CON" Then
       MsgBox "Usuario No Autorizado", vbExclamation
       Exit Sub
    End If
    
    If Not IsDate(Txfecha) Then
       MsgBox "Fecha no valida...", vbExclamation
       Txfecha = Date
       Exit Sub
    End If
      
    On Error GoTo err_trans
    Cnx.BeginTrans
    Dim sql As String
    Set rs = New Recordset
    fecha = Format(Txfecha, "yyyymmdd")
    sql = "SELECT * FROM CpNotas WHERE tarjeta='" + Txtarjeta + "' and fecha='" + fecha + "'"
    rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF Then
       rs.AddNew
       rs!tarjeta = Txtarjeta
       rs!fecha = Txfecha
       rs!usuario = wUsuario
       rs!fecha = Date
    End If
    rs!detalle = Trim(Txdetalle)
    rs.Update
    rs.Close
    If Sw Then
       DD.AddNew
       DD!fecha = Txfecha
    End If
    DD!detalle = Left(Txdetalle, 100)
    DD.Update
    Cnx.CommitTrans
    Limpiacampos
    DataGrid1.Refresh
    cmdElimina.Enabled = False
    DataGrid1.SetFocus
    Exit Sub
err_trans:
    MsgBox Err.Description
 
End Sub
Private Sub cmdsalir_Click()
   Unload Me
   Me.Hide
End Sub
Sub Limpiacampos()
   Txfecha = "":  Txdetalle = ""
   Lblresto.Caption = ""
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

Private Sub Txfecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Txdetalle.SetFocus
      Exit Sub
   End If
End Sub

Private Sub Txdetalle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      cmdGraba.SetFocus
      Exit Sub
   End If
   nlen = Len(Trim(Txdetalle))
   Lblresto.Caption = "Van..  " + Trim(Str(nlen)) + " posc."
End Sub

