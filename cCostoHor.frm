VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form cCostoHor 
   Caption         =   "Horario Asistencial "
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8385
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "cCostoHor.frx":0000
   ScaleHeight     =   3345
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   0
      Picture         =   "cCostoHor.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   615
      Left            =   6720
      Picture         =   "cCostoHor.frx":80EA8
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar"
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "Graba"
      Height          =   615
      Left            =   5880
      Picture         =   "cCostoHor.frx":8101A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Grabar"
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   7680
      Picture         =   "cCostoHor.frx":8118C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cerrar"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Frame frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   5415
      Begin VB.CheckBox ChkExcepcional 
         Alignment       =   1  'Right Justify
         Caption         =   "Excepcional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Txdescripcion 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   9
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox Txhorario 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Txturno 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   720
         MaxLength       =   2
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Lblturno 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   "Horario"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Turno"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1455
      Left            =   1080
      TabIndex        =   6
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
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
      Caption         =   "Horarios Asignados"
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
Attribute VB_Name = "cCostoHor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs_d As ADODB.Recordset
Dim DD As ADODB.Recordset
Dim Sw As Boolean
Dim sql As String

Private Sub Command1_Click()
     c_grupo = "01": c_secuencia = "01B"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub

Private Sub Form_Load()
  cCostoHor.Caption = "Horario Asistencial - " + CCosto.Txdescripcion.Text
  Set DD = New Recordset
  With DD
     .Fields.Append "Turno", adVarChar, 2
     .Fields.Append "horario", adVarChar, 4
     .Fields.Append "Excepcional", adVarChar, 1
     .Fields.Append "descripcion", adVarChar, 30
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
          .Caption = "Turno"
          .Width = Me.TextWidth(String(5, "S"))
         Case 1
          .Caption = "Horario"
          .Width = Me.TextWidth(String(5, "S"))
         Case 2
          .Caption = "Excepcional"
          .Width = Me.TextWidth(String(5, "S"))
         Case 3
          .Caption = "Descripcion"
          .Width = Me.TextWidth(String(30, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
  Next
  
  Set rs = New Recordset
  sql = "SELECT * FROM Cpccostohor where cCosto='" & CCosto.Txcodigo & "' order by  turno,horario"
  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  Do While Not rs.EOF
     DD.AddNew
     DD!turno = rs!turno
     DD!Horario = rs!Horario
     DD!Excepcional = IIf(IsNull(rs!Excepcional), "", rs!Excepcional)
     Set rs_d = New Recordset
     sql = "SELECT * FROM cphorarioNew  WHERE codigo='" & rs!Horario & "'"
     rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
     If Not rs_d.EOF Then
        DD!descripcion = Left(rs_d!descripcion, 30)
     End If
     DD.Update
     rs_d.Close
     rs.MoveNext
  Loop
  rs.Close
  
End Sub
 Private Sub DataGrid1_Click()
   Call Limpiacampos
   TxTurno = Trim(DataGrid1.Columns(0))
   TxHorario = Trim(DataGrid1.Columns(1))
   ChkExcepcional.Value = IIf(DataGrid1.Columns(2) = "X", 1, 0)
   Txdescripcion = Trim(DataGrid1.Columns(3))
End Sub

Private Sub cmdGraba_Click()
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If IsNull(TxTurno) Then
      MsgBox "Ingrese turno...  ", vbExclamation
      Exit Sub
   Else
      Set rs = New Recordset
      '-Turno Asistenciales Programables valor='S'  ¿solo 12hrs?
      sql = "SELECT * FROM cptablas WHERE tabla='31' AND valor='S' AND codigo='" & TxTurno & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "Verifique Turno...  ", vbExclamation
         rs.Close
         Exit Sub
      End If
      rs.Close
   End If
   If IsNull(TxHorario) Then
      MsgBox "Ingrese Horario valido...  ", vbExclamation
      Exit Sub
   Else
      Set rs = New Recordset
      sql = "SELECT * FROM CpHorarioNew WHERE asistencial='S' AND codigo='" & TxHorario & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "Verifique Horario...  ", vbExclamation
         rs.Close
         Exit Sub
      End If
      rs.Close
   End If
   
   Sw = False
   On Error GoTo final
   Cnx.BeginTrans
   Set rs = New Recordset
   sql = "SELECT * FROM CpCcostoHor  WHERE  ccosto='" & CCosto.Txcodigo & "' AND Turno='" & TxTurno & "' AND horario='" & TxHorario & "'"
   rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If rs.EOF Then
      Sw = True
      rs.AddNew
      rs!CCosto = CCosto.Txcodigo
      rs!turno = TxTurno
      rs!Horario = TxHorario
   End If
   rs!Excepcional = IIf(ChkExcepcional.Value = 1, "X", "")
   rs!usuario = wUsuario       'Firma Usuario
   rs!fecha = Date
   rs.Update
   rs.Close
   If Sw Then
      DD.AddNew
      DD!turno = TxTurno
   End If
   DD!Horario = TxHorario
   DD!Excepcional = IIf(ChkExcepcional.Value = 1, "X", "")
   DD!descripcion = Left(Txdescripcion, 30)
   DD.Update
   Cnx.CommitTrans
   GoTo Final1
final:
   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
   Cnx.RollbackTrans
Final1:
   Call Limpiacampos
   DataGrid1.SetFocus
   Sw = False
End Sub
Private Sub CmdElimina_Click()
    If wrights = "CON" Then
       MsgBox "No Autorizado...  ", vbExclamation
       Exit Sub
    End If
    If IsNull(TxTurno.Text) Or IsNull(TxHorario) Then Exit Sub
    If MsgBox("Eliminar  registro ?      ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       Cnx.Execute "DELETE FROM CpCcostoHor WHERE ccosto ='" & CCosto.Txcodigo.Text & "' AND turno='" & TxTurno.Text & "' AND horario='" & TxHorario & "'"
       DD.Delete
    End If
    Call Limpiacampos
End Sub
Private Sub Limpiacampos()
   TxTurno = "": TxHorario = "": ChkExcepcional.Value = 0
   Txdescripcion = ""
   'DataGrid1.SetFocus
End Sub

Private Sub cmdsalir_Click()
  Set rs = Nothing
  Set rs_d = Nothing
  Me.Hide
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
  Set rs_d = Nothing
  Me.Hide
  Unload Me
End Sub
Private Sub TxTurno_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(TxTurno) = "" Then
        Load Ayuda2T
        Ayuda2T.Tag = 1
        Ayuda2T.Show 1
     Else
        TxTurno = UCase(TxTurno)
        Set rs = New Recordset
        sql = "SELECT * FROM Cptablas where tabla='31' and valor='S' and codigo='" & TxTurno & "'"
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If rs.EOF Then
           MsgBox "Turno no es Valido.."
        Else
           Lblturno.Caption = Left(rs!descripcion, 20)
           TxHorario.SetFocus
        End If
        rs.Close
       'DataGrid1.SetFocus
     End If
  End If
End Sub
Private Sub TxHorario_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(TxHorario) = "" Then
        Load Ayuda2H
        Ayuda2H.Tag = 2
        Ayuda2H.Show 1
     Else
        TxHorario = UCase(TxHorario)
        Set rs = New Recordset
        sql = "SELECT * FROM CpHorarioNew where (bloqueo='N' or Bloqueo is NULL) and asistencial='S' and codigo='" & TxHorario & "'"
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If rs.EOF Then
           MsgBox "Horario no es Valido.."
        Else
           Txdescripcion = Left(rs!descripcion, 30)
        End If
       'DataGrid1.SetFocus
     End If
     'valida horario no empleado  en otro turno
     Set rs = New Recordset
     sql = "SELECT * FROM CpCcostoHor  WHERE  ccosto='" & CCosto.Txcodigo & "' AND horario='" & TxHorario & "'"
     rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
     If Not rs.EOF Then
        If Trim(rs!turno) <> Trim(TxTurno) Then
           Beep 500, 350
           If MsgBox("Horario ya se emplea en turno [" + rs!turno + "], continua ?  ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
             TxHorario = ""
             Txdescripcion = ""
           End If
        End If
     End If
     rs.Close
  End If
End Sub

Private Sub ChkExcepcional_Click()
   If ChkExcepcional.Value = 1 Then
      Set rs = New Recordset
      sql = "SELECT COUNT(*)as Cant FROM CpCcostoHor WHERE  ccosto='" & CCosto.Txcodigo & "' AND Turno='" & TxTurno & "' AND Excepcional<>'X'"
      rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      If Not rs.EOF Then
         If rs!cant > 1 Then
            MsgBox "Verifique.. existen horarios de Programa Excepciones..", vbCritical, "[IMPORTANTE]"
         End If
      Else
         MsgBox "Verifique.. turno No tiene horario de Programación..", vbCritical, "[IMPORTANTE]"
      End If
      rs.Close
   End If
End Sub

