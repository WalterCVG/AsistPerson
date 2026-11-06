VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Lactancia 
   Caption         =   "Inscripcion - Lactancia"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Lactancia.frx":0000
   ScaleHeight     =   6270
   ScaleWidth      =   11550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNuevo 
      Height          =   615
      Left            =   4680
      Picture         =   "Lactancia.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Agregar Hijo(a) "
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton CmdIN 
      Caption         =   ">>"
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      ToolTipText     =   "Cargar"
      Top             =   360
      Width           =   615
   End
   Begin VB.ListBox LstHijos 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   240
      TabIndex        =   20
      Top             =   2760
      Width           =   11175
      Begin VB.TextBox TxResolucionE 
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.CommandButton CmdResolucion 
         Caption         =   "Resoluc."
         Height          =   735
         Left            =   1200
         Picture         =   "Lactancia.frx":803D8
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2520
         Width           =   735
      End
      Begin VB.CommandButton cmdElimina 
         Caption         =   "Elimina"
         Height          =   735
         Left            =   8040
         Picture         =   "Lactancia.frx":8054A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Eliminar"
         Top             =   2520
         Width           =   735
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3625
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
         Caption         =   "Periodos de Lactancia"
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
   Begin VB.Frame Frame1 
      Caption         =   "Registro"
      Height          =   2295
      Left            =   5400
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.TextBox TxResolucion 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   3495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Salida"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   11
         Top             =   1920
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ingreso"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   10
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton cmdGraba 
         Caption         =   "Graba"
         Height          =   735
         Left            =   5040
         Picture         =   "Lactancia.frx":806BC
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Grabar"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox TxDniFam 
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TxSexo 
         Height          =   285
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   6
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox TxNombFam 
         Height          =   285
         Left            =   1080
         MaxLength       =   45
         TabIndex        =   4
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox TxFecNac 
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
         Left            =   3720
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TxFecFin 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Resolución:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "TOLERANCIA  al :"
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Dni"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Sexo"
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Beneficio Caduca:"
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Nacimiento"
         Height          =   255
         Left            =   2400
         TabIndex        =   18
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "LEGAJOS - Hijos Lactantes registrados"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Lactancia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DD As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim sql, cDni, cTarjeta, ccadena, cresolu, cAplica As String
Dim DFecha0 As Date
Private Sub Form_Load()
  cDni = Trim(General.TxDocumento)
  cTarjeta = Trim(General.Txtarjeta)
  Me.Caption = Trim(General.Txapelpat) + " " + Trim(General.Txapelmat) + " " + Trim(General.TxNombres) + " - Inscripción LACTANCIA"
  '-busca fecha en MBD
  dFbase = Format(DateAdd("YYYY", -1, Date), "mm/dd/yyyy")
  ccadena = ""
  nson = 0
  Set rs = New Recordset
  sql = "SELECT * FROM Familiar WHERE Dni='" & Trim(cDni) & "' AND Parentesco='H' AND Fecnac >=#" & dFbase & "#"
  rs.Open sql, Cnx1, adOpenStatic, adLockReadOnly, adCmdText
  Do While Not rs.EOF
     LstHijos.AddItem Left(rs!NombFam, 32) & " " + Format(rs!fecnac, "dd/mm/yyyy") + " " + rs!sexo
     '         1         2         3         4
     '123456789012345678901234567890123456789012345
     'xxxxxxxxx1xxxxxxxxx2xxxxaaaaa3a2 99/99/9999 X
     ccadena = ccadena + Format(rs!fecnac, "dd/mm/yyyy")   'c/10posc
     nson = nson + 1
     rs.MoveNext
  Loop
  rs.Close
  '-LACTANCIAS  ASIGNADAS
  Set DD = New Recordset
  With DD
     .Fields.Append "NombFam", adVarChar, 45
     .Fields.Append "fecnac", adDate
     .Fields.Append "sexo", adVarChar, 1
     .Fields.Append "DniFam", adVarChar, 8
     .Fields.Append "fecFin", adDate
     .Fields.Append "Aplica", adVarChar, 1
     .Fields.Append "Resolucion", adVarChar, 50
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
        Select Case i
         Case 0
          .Caption = "Nombre Hijo(a)"
          .Width = Me.TextWidth(String(35, "S"))
         Case 1
          '.Caption = "Fec.Nac."
          .Width = Me.TextWidth(String(10, "S"))
         Case 2
          .Caption = "Sexo"
          .Width = Me.TextWidth(String(5, "S"))
         Case 3
           .Caption = "Dni"
          .Width = Me.TextWidth(String(8, "S"))
         Case 4
          .Caption = "Caduca"
          .Width = Me.TextWidth(String(10, "S"))
         Case 5
          '.Caption = "Aplica"       'I/S
          .Width = Me.TextWidth(String(6, "S"))
         Case 6
          '.Caption = "Resolucion"       'I/S
          .Width = Me.TextWidth(String(30, "S"))
        End Select
        '.NumberFormat = "##,##0.0"
        .Locked = True
     End With
  Next
  Set rs = New Recordset
  sq2 = "SELECT * FROM CpLactancia WHERE tarjeta ='" & Trim(cTarjeta) & "'"
  rs.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
  Do While Not rs.EOF
     DD.AddNew
     DD!NombFam = rs!NombFam
     If IsDate(rs!fecnac) Then DD!fecnac = rs!fecnac
     DD!sexo = IIf(IsNull(rs!sexo), "", rs!sexo)
     DD!dnifam = IIf(IsNull(rs!dnifam), "", rs!dnifam)
     If IsDate(rs!FecFIN) Then DD!FecFIN = rs!FecFIN
     DD!Aplica = IIf(IsNull(rs!Aplica), "", rs!Aplica)
     DD!resolucion = IIf(IsNull(rs!resolucion), "", rs!resolucion)
     DD.Update
     rs.MoveNext
  Loop
  'rs!usuario = wUsuario      'Firma Usuario
  'rs!fecRegistro = Date
  rs.Close
  Call Limpiacampos
  If nson = 0 Then
     Beep 500, 350
     MsgBox ">> LEGAJOS <<..No registró, hijos en Lactancia..", vbExclamation
  End If
End Sub

Private Sub cmdIN_Click()
   Call Limpiacampos
   j = LstHijos.ListIndex
   '         1         2         3         4         5         6
   '123456789012345678901234567890123456789012345678901234567890
   'aaaaa6 xxxxxxxxx1xxxxxxxxx2 - 9900-9900 / 9900-9900  Disp.:99
   '                              12345678901234567890123
   If j <> -1 Then
      TxFecnac.Text = Mid(ccadena, j * 10 + 1, 10)
      fecha1 = Format(TxFecnac, "mm/dd/yyyy")
      Set rs = New Recordset
      sql = "SELECT * FROM Familiar WHERE Dni='" & Trim(cDni) & "' AND Parentesco='H' AND Fecnac =#" & fecha1 & "#"
      rs.Open sql, Cnx1, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then
         TxNombFam = rs!NombFam
         TxDniFam = rs!dnifam
         TxSexo = rs!sexo
      End If
      rs.Close
      Option1(0).Value = False: Option1(1).Value = False:
      TxFecFin = DateAdd("YYYY", 1, DateValue(TxFecnac.Text)) - 1
      Habilitacampos (False)
      Frame1.Enabled = True
      TxResolucion.Enabled = True
      Option1(0).Enabled = True: Option1(1).Enabled = True
   End If
End Sub
Private Sub cmdNuevo_Click()
   Call Limpiacampos
   Habilitacampos (True)
   TxNombFam.SetFocus
End Sub

Private Sub Limpiacampos()
   TxNombFam = "": TxDniFam = "": TxSexo = "": TxFecnac = "": TxFecFin = "": TxResolucion = ""
   TxResolucionE = "": TxResolucionE.Visible = False
   Option1(0).Value = False: Option1(1).Value = False
   Frame1.Enabled = False
   cmdElimina.Enabled = False
End Sub

Private Sub Habilitacampos(lopc As Boolean)
   TxNombFam.Enabled = lopc: TxDniFam.Enabled = lopc: TxSexo.Enabled = lopc: TxFecnac.Enabled = lopc: TxFecFin.Enabled = lopc: TxResolucion.Enabled = lopc
   Option1(0).Enabled = lopc: Option1(1).Enabled = lopc
   Frame1.Enabled = True
End Sub
Private Sub DataGrid1_Click()

'    Limpiacampos
'    If Not IsNull(DataGrid1.Columns(0)) Then
'       TxNombFam = DataGrid1.Columns(0)
'    Else
'       TxNombFam = ""
'    End If
'    If IsDate(DataGrid1.Columns(1)) Then TxFecNac = DataGrid1.Columns(1)
'    TxSexo = DataGrid1.Columns(2)
'    TxDniFam.Text = DataGrid1.Columns(3)
'    TxResolucion = DataGrid1.Columns(6)
'    Frame2.Enabled = True

   cmdElimina.Enabled = True
   cmdElimina.SetFocus
   
End Sub

Private Sub cmdGraba_Click()
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If Option1(0).Value = False And Option1(1).Value = False Then
      MsgBox "No Indica Tolerancia... ", vbExclamation
      Exit Sub
   End If
   If Trim(cTarjeta) = "" Then
      MsgBox "SERVIDOR no Indica Tarjeta... ", vbExclamation
      Exit Sub
   Else
      If Trim(cDni) = "" Then
         MsgBox "SERVIDOR No Indica DNI... ", vbExclamation
         Exit Sub
      End If
      If Trim(TxNombFam) = "" Then
         MsgBox "Indique nombres del lactante...  ", vbExclamation
         Exit Sub
      End If
     
      If Not IsDate(TxFecFin) Then
         MsgBox "Verifique fecha...  ", vbExclamation
         Exit Sub
      End If
      If Year(TxFecFin) < Val(genaño) Then
         MsgBox "Fecha.. Vencida No valida...  ", vbExclamation
         Exit Sub
      End If
      Set rs = New Recordset
      sql = "SELECT * FROM CpGeneral  WHERE (baja is NULL) and tarjeta='" & cTarjeta & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "Error, Tarjeta No existe..! ", vbExclamation
         rs.Close
         Exit Sub
      Else
         rs.Close
      End If
   End If
   fecha1 = Format(TxFecnac, "yyyymmdd")
   If Option1(0).Value = True Then
      cAplica = "I"
   ElseIf Option1(1).Value = True Then
      cAplica = "S"
   End If
      
      
   '-valida mas de 1 hijo en Fecnac y como aplica
   Set rt = New Recordset
   sq2 = "SELECT * FROM CpLactancia WHERE DNI ='" & Trim(cDni) & "' and fecNac='" & fecha1 & "'  order by Nombfam "
   rt.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   Do While Not rt.EOF
      If rt!NombFam <> Trim(TxNombFam) And rt!Aplica <> cAplica Then
         cAplica = rt!Aplica
         If cAplica = "I" Then
            Option1(0).Value = True
         ElseIf cAplica = "S" Then
                Option1(1).Value = True
         End If
         Exit Do
      End If
      rt.MoveNext
   Loop
   
   
   On Error GoTo final
   Cnx.BeginTrans
   Set rt = New Recordset
   sq2 = "SELECT * FROM CpLactancia WHERE DNI ='" & Trim(cDni) & "' and fecNac='" & fecha1 & "' and  NombFam ='" + Trim(TxNombFam) + "'"
   rt.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If rt.EOF Then
      rt.AddNew
      rt!Dni = UCase(cDni)
      rt!tarjeta = UCase(cTarjeta)
      If IsDate(TxFecnac) Then rt!fecnac = TxFecnac
      rt!NombFam = Trim(TxNombFam)
   End If
   rt!sexo = TxSexo
   rt!Aplica = cAplica
   rt!FecFIN = TxFecFin
   rt!resolucion = Trim(TxResolucion)
   rt!usuario = wUsuario      'Firma Usuario
   rt!fecRegistro = Date
   rt.Update
   rt.Close
   DD.AddNew
   DD!NombFam = UCase(Trim(TxNombFam))
   If IsDate(TxFecnac) Then DD!fecnac = TxFecnac
   DD!sexo = IIf(IsNull(TxSexo), "", UCase(TxSexo))
   DD!dnifam = IIf(IsNull(TxDniFam), "", TxDniFam)
   If IsDate(TxFecFin) Then DD!FecFIN = TxFecFin
   DD!Aplica = IIf(IsNull(cAplica), "", cAplica)
   DD!resolucion = UCase(Trim(TxResolucion))
   DD.Update
   Cnx.CommitTrans
   Call Limpiacampos
   GoTo Final1
final:
   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
   Cnx.RollbackTrans
Final1:
   Call Limpiacampos
   'Call Desabilitacampos
   'DataGrid1.Row = 0
   'DataGrid1.Col =
   DataGrid1.SetFocus
   Sw = False
End Sub

Private Sub CmdElimina_Click()
    If wrights = "CON" Then
       MsgBox "No Autorizado...  ", vbExclamation
       Exit Sub
    End If
    If Trim(DataGrid1.Columns(0)) = "" Then Exit Sub
    ccadena = Trim(DataGrid1.Columns(0))        'NombFam
    fecha1 = Format(DateValue(DataGrid1.Columns(1)), "yyyymmdd")
    Set rs = New Recordset
    sq2 = "SELECT * FROM CpLactancia  WHERE Dni ='" & Trim(cDni) & "' and fecnac='" & fecha1 & "' and NombFam='" + ccadena + "'"
    rs.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF Then
       MsgBox "No Ubico Registro...  ", vbExclamation
       rs.Close
       Exit Sub
    Else
       rs.Close
    End If
    If MsgBox("Eliminar  registro de '" + ccadena + "' ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       Cnx.Execute "DELETE FROM CpLactancia  WHERE Dni ='" & Trim(cDni) & "' and fecnac='" & fecha1 & "' and NombFam='" + ccadena + "'"
       DD.Delete
    End If
    Call Limpiacampos
    'Call Desabilitacampos
End Sub

Private Sub CmdResolucion_Click()
     ccadena = Trim(DataGrid1.Columns(0))        'NombFam
     TxResolucionE = Trim(DataGrid1.Columns(6))
     cresolu = Trim(DataGrid1.Columns(6))       'NombFam
     TxResolucionE.Visible = True
     TxResolucionE.SetFocus
End Sub
Private Sub TxResolucionE_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxResolucionE = UCase(TxResolucionE)
      fecha1 = Format(DateValue(DataGrid1.Columns(1)), "yyyymmdd")
      Cnx.CommandTimeout = 300
         
      sql = "Update CpLactancia Set Resolucion='" + Trim(TxResolucionE) + "' where  DNI ='" & Trim(cDni) & "' and fecNac='" & fecha1 & "' and NombFam='" + ccadena + "'"
      Set rs = Cnx.Execute(sql)
      
      DD!resolucion = TxResolucionE
      DD.Update
      DataGrid1.Refresh
      Call Limpiacampos
   End If
End Sub

Private Sub cmdsalir_Click()
   Set DD = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "General" Then
           Unload Forms(cont)
        End If
      End If
   Next
   Unload Me
End Sub
Private Sub Txfecnac_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxFecFin = DateAdd("YYYY", 1, DateValue(TxFecnac.Text)) - 1
   End If
End Sub
Private Sub TxFecNac_LostFocus()
  Call Txfecnac_KeyPress(13)
End Sub

Private Sub TxNombFam_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxNombFam = UCase(TxNombFam)
   End If
End Sub
Private Sub TxNombFam_LostFocus()
  Call TxNombFam_KeyPress(13)
End Sub
Private Sub Txsexo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxSexo = UCase(TxSexo)
   End If
   If Not ((KeyAscii = 70 Or KeyAscii = 102 Or KeyAscii = 77 Or KeyAscii = 109 Or KeyAscii = 77)) And Not KeyAscii = 8 Then
      KeyAscii = 0
   End If
   
End Sub
Private Sub TxSexo_LostFocus()
   Call Txsexo_KeyPress(13)
End Sub

Private Sub TxResolucion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxResolucion = UCase(TxResolucion)
   End If
End Sub
Private Sub TxResolucion_LostFocus()
  Call TxResolucion_KeyPress(13)
End Sub

