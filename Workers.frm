VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Workers 
   Caption         =   "Servidores en  Cadena"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Workers.frx":0000
   ScaleHeight     =   6030
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txocadhep 
      Enabled         =   0   'False
      Height          =   195
      Left            =   5400
      MaxLength       =   35
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdVerWkrs 
      Caption         =   "Integrantes"
      Height          =   735
      Left            =   360
      Picture         =   "Workers.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   8520
      Picture         =   "Workers.frx":80B30
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cerrar"
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox TxDnix 
      Enabled         =   0   'False
      Height          =   195
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   12
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox TxCadHep 
      Enabled         =   0   'False
      Height          =   195
      Left            =   7080
      MaxLength       =   52
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox TxNombres 
      Enabled         =   0   'False
      Height          =   195
      Left            =   4200
      MaxLength       =   20
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Txapelmat 
      Enabled         =   0   'False
      Height          =   195
      Left            =   3120
      MaxLength       =   20
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Txapelpat 
      Enabled         =   0   'False
      Height          =   195
      Left            =   1920
      MaxLength       =   20
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Txtarjeta 
      Enabled         =   0   'False
      Height          =   195
      Left            =   840
      MaxLength       =   5
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton CmdProcesa 
      Caption         =   "Procesa"
      Height          =   615
      Left            =   5520
      TabIndex        =   6
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton CmdOUT 
      Caption         =   "<"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Retirar"
      Top             =   5400
      Width           =   615
   End
   Begin VB.CommandButton CmdIN 
      Caption         =   ">"
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Agregar"
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox TxTrabajador 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      MaxLength       =   45
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
   Begin VB.TextBox TxDni 
      Height          =   285
      Left            =   840
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7646
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
      Caption         =   "Servidores a Incorporar en Cadena"
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
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Servidor"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Workers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim DD As ADODB.Recordset
Dim Sw As Boolean
Dim sql As String
Dim ndato As Long
Private Sub Form_Load()
  Workers.Caption = "Servidores en Cadena - " + tCadena.Txcadena.Text
  TxCadHep.Text = tCadena.Txcadena.Text
  Set DD = New Recordset
  With DD
     .Fields.Append "Apelpat", adVarChar, 20
     .Fields.Append "Apelmat", adVarChar, 20
     .Fields.Append "Nombres", adVarChar, 20
     .Fields.Append "Documento", adVarChar, 10
     .Fields.Append "tarjeta", adVarChar, 5
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
        Select Case i
         Case 0
           .Caption = "Apel.Pat"
           .Width = Me.TextWidth(String(15, "S"))
         Case 1
           .Caption = "Apel.Mat"
           .Width = Me.TextWidth(String(15, "S"))
         Case 2
           .Caption = "Nombres"
           .Width = Me.TextWidth(String(20, "S"))
         Case 3
           .Caption = "Dni"
           .Width = Me.TextWidth(String(10, "S"))
         Case 4
           .Caption = "Tarjeta"
           .Width = Me.TextWidth(String(5, "S"))
        End Select
        '.NumberFormat = "##,##0.0"
        .Locked = True
     End With
  Next
End Sub

Private Sub TxDni_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(TxDni) = "" Then
         Load Ayuda1
         Ayuda1.Tag = 0.1
         Ayuda1.Left = 495: Ayuda1.Top = 1245
         Ayuda1.Show 1
      Else
         Set rs = New Recordset
         sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and Documento = '" & TxDni.Text & "'"
         rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If rs.EOF Then
            MsgBox "Documento, No Ubicado", vbExclamation
            TxDni = ""
            TxTrabajador = ""
            Exit Sub
         End If
         TxTrabajador = Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres)
         Txapelpat = IIf(IsNull(rs!apelpat), "", Trim(rs!apelpat))
         Txapelmat = IIf(IsNull(rs!apelmat), "", Trim(rs!apelmat))
         TxNombres = IIf(IsNull(rs!nombres), "", Trim(rs!nombres))
         Txtarjeta = Trim(rs!tarjeta)
         Txocadhep = IIf(IsNull(rs!CadHep), "", Left(rs!CadHep, 35))
         rs.Close
      End If
      If Trim(TxTrabajador) <> "" Then
         '.backcolor=&H8000000F&      'Default
         Sw = False
         Call DNIenDD
         If Sw Then
            MsgBox "Ya registrado ...", vbExclamation
            Call Limpiacampos
         Else
            CmdIN.BackColor = &H80FF80
            CmdIN.SetFocus
         End If
      End If
   End If
End Sub
Private Sub DNIenDD()
     Dim cBusDni As String
     cpuntero = DataGrid1.Bookmark
     If Trim(TxDni) = "" Then Exit Sub
     cBusDni = "DOCUMENTO='" & Trim(TxDni) & "'"
     If Not DD.EOF Then
        DD.MoveFirst
        DD.Find (cBusDni)
     End If
     If DD.EOF Then
     Else
        Sw = True
        DataGrid1.Refresh
     End If
End Sub

Private Sub cmdIN_Click()
   If Trim(Txocadhep) <> "" And TxCadHep = Txocadhep Then
      MsgBox "Servidor pertenece a la Cadena..", vbExclamation
      Limpiacampos
      Exit Sub
   End If

   '.backcolor=&H8000000F&      'Default
   '.backcolor=&H0080FF80&
   DD.AddNew
   DD!apelpat = Txapelpat
   DD!apelmat = Txapelmat
   DD!nombres = TxNombres
   DD!documento = TxDni
   DD!tarjeta = Txtarjeta
   DD.Update
   '.BackColor = &H80FF80
   CmdIN.BackColor = &H8000000F     'Default
   Call Limpiacampos
End Sub

Private Sub DataGrid1_Click()
   If IsEmpty(DataGrid1.Columns(0)) Then Exit Sub
   TxDnix.Text = DataGrid1.Columns(3)
   '.backcolor=&H8000000F&      'Default
   CmdOUT.BackColor = &H8080FF
   CmdOUT.SetFocus
End Sub

Private Sub CmdOUT_Click()
    If wrights = "CON" Then
       MsgBox "No Autorizado...  ", vbExclamation
       Exit Sub
    End If
    If Trim(TxDnix.Text) = "" Then Exit Sub
    If MsgBox("Retira registro ?      ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       DD.Delete
    End If
    '.backcolor=&H008080FF&
    CmdOUT.BackColor = &H8000000F     'Default
    Call Limpiacampos
End Sub

Private Sub CmdProcesa_Click()
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If DD.RecordCount = 0 Then
      MsgBox "No incorpora Trabajador...  ", vbExclamation
      Exit Sub
   End If
''   On Error GoTo final
''   Cnx.BeginTrans
   DD.MoveFirst
   Dim sql As String
   Do While Not DD.EOF
      Cnx.BeginTrans
      If Trim(DD!documento) <> "" Then
         cDni = Trim(DD!documento)
         sql = "UPDATE CpGeneral SET oCadhep=CadHep, CadHep='" + Trim(TxCadHep.Text) + "' WHERE (baja is NULL) and  Documento='" + cDni & "' AND  CadHep<>'" + Trim(TxCadHep.Text) + "'"
      Else
         cDni = Trim(DD!tarjeta)
         sql = "UPDATE CpGeneral SET oCadhep=CadHep, CadHep='" + Trim(TxCadHep.Text) + "' WHERE (baja is NULL) and tarjeta='" + cDni & "' AND  CadHep<>'" + Trim(TxCadHep.Text) + "'"
      End If
      Cnx.Execute sql
      Cnx.CommitTrans
      DD.MoveNext
   Loop
   DD.MoveFirst
   Do While Not DD.EOF
      DD.Delete
      DD.MoveNext
   Loop
   DataGrid1.Refresh
   GoTo Final1
''final:
''   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
''   Cnx.RollbackTrans
Final1:
   Call Limpiacampos
 
End Sub
Private Sub Limpiacampos()
   TxDni.Text = "": Txtarjeta.Text = "": Txocadhep.Text = ""
   TxTrabajador.Text = "": Txapelpat.Text = "": Txapelmat.Text = "": TxNombres.Text = ""
   TxDnix.Text = ""
End Sub
Private Sub CmdVerWkrs_Click()
      Call ChrgWorkrs(1)
      If ndato > 0 Then
         nTipRpt = 0.1
         InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
         InfoCpSQL.Refresh
         InfoCpSQL.Show
      Else
         MsgBox "No hay servidores en dicha cadena..", vbExclamation
      End If
End Sub

Private Sub ChrgWorkrs(nopc)
'***  CARGA  registros  ***   0:Todos   1:cadena
    Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmGeneral"
    ndato = 0
    If nopc = 0 Then
       sql = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and estado='A' ORDER BY cadhep,apelpat,apelmat,nombres "
       wTituRpt = "(Todas las Cadenas)"
    Else
      sql = "SELECT * FROM CpGeneral where (baja is NULL) and  estado='A' and  cadhep='" + Trim(TxCadHep.Text) + "' ORDER BY apelpat,apelmat,nombres "
      wTituRpt = "(" + Trim(TxCadHep.Text) + ")"
    End If
    '''On Error Resume Next
    Set rs_d = New Recordset
    rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Set rt = New Recordset
    sq_l = "SELECT * FROM CpTmGeneral"
    rt.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    Do While Not rs_d.EOF
       ndato = ndato + 1
       rt.AddNew
       rt!tarjeta = rs_d!tarjeta
       rt!apelpat = IIf(IsEmpty(rs_d!apelpat), "", UCase(rs_d!apelpat))
       rt!apelmat = IIf(IsEmpty(rs_d!apelmat), "", UCase(rs_d!apelmat))
       rt!nombres = IIf(IsEmpty(rs_d!nombres), "", UCase(rs_d!nombres))
       rt!estado = IIf(IsEmpty(rs_d!estado), "", rs_d!estado)
       rt!documento = IIf(IsEmpty(rs_d!documento), "", rs_d!documento)
       rt!CadHep = IIf(IsEmpty(rs_d!CadHep), "", rs_d!CadHep)
       rt!CCosto = IIf(IsEmpty(rs_d!CCosto), "", rs_d!CCosto)
       rt!CodUpr = IIf(IsEmpty(rs_d!CodUpr), "", rs_d!CodUpr)
       cCodigo = IIf(IsNull(rs_d!CadHep), "", rs_d!CadHep)
       If Trim(cCodigo) <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM CpcadHep WHERE cadHep='" & Trim(cCodigo) & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dsccadena = "" + Left(rs!descripcion, 110)
          End If
          rs.Close
       End If
       cCodigo = IIf(IsNull(rs_d!CCosto), "", rs_d!CCosto)
       If IsEmpty(cCodigo) Then
       Else
          Set rs = New Recordset
          rs.Open "SELECT * FROM CpCcosto WHERE codigo='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dscccosto = "" + UCase(Left(rs!descripcion, 35))
          End If
          rs.Close
       End If
       cCodigo = IIf(IsNull(rs_d!CodUpr), "", rs_d!CodUpr)
       If IsEmpty(cCodigo) Then
       Else
          Set rs = New Recordset
          rs.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodigo & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dscUPR = "" + UCase(Left(rs!descripcion, 55))
          End If
          rs.Close
       End If
       rt.Update
       rs_d.MoveNext
    Loop
    Cnx.CommitTrans
    rt.Close
    rs_d.Close
End Sub

Private Sub cmdsalir_Click()
  Set rs = Nothing
  Me.Hide
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  Me.Hide
  Unload Me
End Sub

