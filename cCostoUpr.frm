VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form cCostoUpr 
   Caption         =   "RESPONSABLES DE PROGRAMACIÓN"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "cCostoUpr.frx":0000
   ScaleHeight     =   5445
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   7695
      Begin VB.TextBox Txfn3 
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Txfn2 
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox TxFn1 
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TxEst3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxEst2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox TxEst1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Txdni3 
         Height          =   285
         Left            =   360
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Txdni2 
         Height          =   285
         Left            =   360
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox CboUpr 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   6015
      End
      Begin VB.TextBox TxCodUpr 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   2
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Txdni1 
         Height          =   285
         Left            =   360
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Lblusu3 
         Caption         =   "Lblusu3"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label Lblusu2 
         Caption         =   "Lblusu2"
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Label Lblusu1 
         Caption         =   "Lblusu1"
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Titulares (DNI)"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "Und.Prg.Rol"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   8400
      Picture         =   "cCostoUpr.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cerrar"
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "Graba"
      Height          =   615
      Left            =   8040
      Picture         =   "cCostoUpr.frx":807E8
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Grabar"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdElimina 
      Caption         =   "Elimina"
      Height          =   615
      Left            =   8040
      Picture         =   "cCostoUpr.frx":8095A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Eliminar"
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   0
      Picture         =   "cCostoUpr.frx":80ACC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3413
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
      Caption         =   "Grupos ""UPR"" Asignados"
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6960
      TabIndex        =   19
      Top             =   5040
      Width           =   2055
   End
End
Attribute VB_Name = "cCostoUpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs_d As ADODB.Recordset
Dim DD As ADODB.Recordset
Dim Sw As Boolean
Dim sql, cDni1, cDni2, cDni3 As String
Dim anoi, anof, mesi, mesf As String
Private Sub CmdExportaXLS_Click()
   Criterio = ""
   If TxfechaIni <> "" And TxFechaFin <> "" Then
      anoi = Left(Format(TxfechaIni, "YYYYMMDD"), 4)
      anof = Left(Format(TxFechaFin, "YYYYMMDD"), 4)
      mesi = Mid(Format(TxfechaIni, "YYYYMMDD"), 5, 2)
      mesf = Mid(Format(TxFechaFin, "YYYYMMDD"), 5, 2)
      Criterio = " ( ano>= '" & anoi & "' and  ano<= '" & anof & "')"
   End If
   If Criterio = "" Then
      MsgBox "Seleccione Criterio ", vbExclamation
      Exit Sub
   End If
 
   Sq0 = " CREATE Table XPROGUPR (Codupr varchar(2) NULL, ano  varchar(4) NULL, del_mes varchar(2) NULL, aL_mes varchar(2) NULL,dscUPR varchar(50) NULL, usuario varchar(10) NULL, Responsable  varchar(40) NULL, cantidad int NULL) "
   Cnx.Execute (Sq0)
   
   'Set rs_d = New Recordset
   'Sq0 = "SELECT  * FROM  XProgUPR ORDER BY Codupr,ano,mes"
   'rs_d.Open Sq0, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   
   Cnx.BeginTrans
   Set rs = New Recordset
   sql = "SELECT codupr, ANO, MES, usuario, COUNT(usuario) AS cant, FLAG From CPDIRECTORIO " & _
     "WHERE  (ANO >= '" + anoi + "' and ANO<='" + anof + "') " & _
     "GROUP BY codupr,ANO, MES,  usuario, FLAG "
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   
   wdato = 0
   Do While Not rs.EOF
      If (rs!ano = anoi And rs!mes < mesi) Or (rs!ano = anof And rs!mes > mesf) Or rs!Flag <> "0" Then
      Else
         cCodUpr = rs!CodUpr
         cano = rs!ano
         cmes = rs!mes
         CUSUARIO = rs!usuario
         Sq0 = "SELECT *  FROM XPROGuPR WHERE CODUPR='" + cCodUpr + "' AND ANO='" + cano + "' AND USUARIO='" + CUSUARIO + "'"
         rs_d.Open Sq0, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If rs_d.EOF() Then
            rs_d.AddNew
            rs_d!CodUpr = rs!CodUpr
            Set rt = New Recordset
            rt.Open "SELECT * FROM CpUndPrgRol WHERE codUpr='" & rs!CodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
            If Not rt.EOF Then
               rs_d!dscUPR = Left(rt!descripcion, 50)
            End If
            rt.Close
            rs_d!ano = rs!ano
            rs_d!del_MES = rs!mes
            rs_d!usuario = rs!usuario
            cdocum = rs!usuario
            If Val(cdocum) > 0 Then
               Set rt = New Recordset
               rt.Open "SElect * from CpGeneral where documento='" + cdocum + "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rt.EOF() Then
                  rs_d!responsable = Left(Trim(rt!apelpat) + " " + Trim(rt!apelmat) + " " + rt!nombres, 40)
               End If
               rt.Close
            Else
               Set rt = New Recordset
               rt.Open "Select * from CpUsuarios where usuario='" + Trim(cdocum) + "'", Cnx, adOpenStatic, adLockReadOnly
               If Not rt.EOF() Then
                  rs_d!responsable = Left(rt!nombre, 40)
               Else
                  rs_d!responsable = cdocum
               End If
               rt.Close
            End If
            rs_d!cantidad = rs!cant
         Else
            If rs_d!del_MES > rs!mes Then
               rs_d!Al_mes = rs_d!del_MES
               rs_d!del_MES = rs!mes
            Else
               If rs_d!Al_mes > rs!mes Then
               Else
                  rs_d!Al_mes = rs!mes
               End If
            End If
         End If
         rs_d.Update
         rs_d.Close
         wdato = wdato + 1
      End If
      rs.MoveNext
   Loop
   Cnx.CommitTrans
   rs.Close
   If wdato > 1 Then
      wTpoQuery = "0.5"
      wNameXLS = "C:\RPT\PROGUpr.xls"
      Call Exportar_ADO_Excel
   End If
   Beep 500, 350
   MsgBox "Se generó C:\RPT\ProgUpr.xls", vbInformation
   sql = "DROP Table XProgUpr "
   Cnx.Execute (sql)
   Exit Sub
End Sub
   
Private Sub Command1_Click()
     c_grupo = "01": c_secuencia = "01C"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub

Private Sub Form_Load()
  LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
  cCostoUpr.Caption = "RESPONSABLES DE PROGRAMACIÓN - " + CCosto.Txdescripcion.Text
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpUndPrgRol ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To CboUpr.ListCount - 1
         If CboUpr.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("Codupr") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        CboUpr.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("Codupr")
     End If
     rs.MoveNext
  Wend
  rs.Close
  Set DD = New Recordset
  With DD
     .Fields.Append "codupr", adVarChar, 2
     .Fields.Append "descripcion", adVarChar, 55
     .Fields.Append "dni1", adVarChar, 10
     .Fields.Append "dni2", adVarChar, 10
     .Fields.Append "dni3", adVarChar, 10
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         '.Caption =
          .Width = Me.TextWidth(String(5, "S"))
         Case 1
          .Caption = "Descripcion"
          .Width = Me.TextWidth(String(55, "S"))
         Case 2
          .Caption = "Dni 1"
          .Width = Me.TextWidth(String(8, "S"))
         Case 3
          .Caption = "Dni 2"
          .Width = Me.TextWidth(String(8, "S"))
         Case 4
          .Caption = "Dni 3"
          .Width = Me.TextWidth(String(8, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
  Next
  
  Set rs = New Recordset
  sql = "SELECT * FROM CpcCostoUpr where cCosto='" & CCosto.Txcodigo.Text & "'"
  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  Do While Not rs.EOF
     DD.AddNew
     DD!CodUpr = rs!CodUpr
     DD!Dni1 = IIf(IsNull(rs!Dni1), "", Trim(rs!Dni1))
     DD!Dni2 = IIf(IsNull(rs!Dni2), "", Trim(rs!Dni2))
     DD!dni3 = IIf(IsNull(rs!dni3), "", Trim(rs!dni3))
     Set rs_d = New Recordset
     rs_d.Open "SELECT * FROM CpUndPrgRol where codupr='" & rs!CodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
     If Not rs_d.EOF Then
        DD!descripcion = rs_d!descripcion
     End If
     rs_d.Close
     DD.Update
     rs.MoveNext
  Loop
  rs.Close
End Sub
 Private Sub DataGrid1_Click()
   Call Limpiacampos
   TxcodUPR = IIf(IsEmpty(DataGrid1.Columns(0)), "", DataGrid1.Columns(0))
   Txdni1 = IIf(IsEmpty(DataGrid1.Columns(2)), "", DataGrid1.Columns(2))
   cDni1 = Trim(Txdni1)
   Txdni2 = IIf(IsEmpty(DataGrid1.Columns(3)), "", DataGrid1.Columns(3))
   cDni2 = Trim(Txdni2)
   Txdni3 = IIf(IsEmpty(DataGrid1.Columns(4)), "", DataGrid1.Columns(4))
   cDni3 = Trim(Txdni3)
   If Trim(TxcodUPR) <> "" Then
       For i = 0 To CboUpr.ListCount - 1
         npos = InStr(CboUpr.List(i), ":")
         nlen = Len(CboUpr.List(i))
         dato = Trim(Right(CboUpr.List(i), nlen - npos))
         If dato = Trim(TxcodUPR) Then
            CboUpr.ListIndex = i
            Exit For
         End If
       Next i
   End If
   If Txdni1 <> "" Then
      Set rs = New Recordset
      sql = "SELECT * FROM cpgeneral WHERE (baja is NULL) and documento='" & Txdni1 & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then
         Lblusu1.Caption = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
         If IsDate(rs!fecnac) Then TxFn1 = rs!fecnac
         TxEst1 = rs!estado
      End If
      rs.Close
   End If
   If Txdni2 <> "" Then
      Set rs = New Recordset
      sql = "SELECT * FROM cpgeneral WHERE (baja is NULL) and documento='" & Txdni2 & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then
         Lblusu2.Caption = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
         If IsDate(rs!fecnac) Then Txfn2 = rs!fecnac
         TxEst2 = rs!estado
      End If
      rs.Close
   End If
   If Txdni3 <> "" Then
      Set rs = New Recordset
      sql = "SELECT * FROM cpgeneral WHERE  (baja is NULL) and documento='" & Txdni3 & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then
         Lblusu3.Caption = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
         If IsDate(rs!fecnac) Then Txfn3 = rs!fecnac
         TxEst3 = rs!estado
      End If
      rs.Close
   End If
End Sub

Private Sub cmdGraba_Click()
   If wrights = "CON" Then
      MsgBox "No Autorizado...  ", vbExclamation
      Exit Sub
   End If
   If IsNull(TxcodUPR) Then
      MsgBox "Ingrese Unid. Prog. Roles...  ", vbExclamation
      Exit Sub
   End If
   If IsNull(Txdni1) Then
      MsgBox "Ingrese titular de Roles...  ", vbExclamation
      Exit Sub
   Else
      Set rs = New Recordset
      sql = "SELECT * FROM cpgeneral WHERE (baja is NULL) and  documento='" & Txdni1 & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "Titular No es Ubicado...  ", vbExclamation
         rs.Close
         Exit Sub
      End If
      rs.Close
   End If
   If IsNull(Txdni2) Or Trim(Txdni2) = "" Then
   Else
      Set rs = New Recordset
      sql = "SELECT * FROM cpgeneral WHERE  (baja is NULL) and documento='" & Txdni2 & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "2º Titular  No es Ubicado...  ", vbExclamation
         rs.Close
         Exit Sub
      End If
      rs.Close
   End If
   If IsNull(Txdni3) Or Trim(Txdni3) = "" Then
   Else
      Set rs = New Recordset
      sql = "SELECT * FROM cpgeneral WHERE  (baja is NULL) and documento='" & Txdni3 & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox "3º Titular  No es Ubicado...  ", vbExclamation
         rs.Close
         Exit Sub
      End If
      rs.Close
   End If
   
   Sw = False
   On Error GoTo final
   Cnx.BeginTrans
   Set rs = New Recordset
   sql = "SELECT * FROM CpcCostoUPR  WHERE  ccosto='" & CCosto.Txcodigo & "' AND CodUpr='" & TxcodUPR & "'"
   rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   If rs.EOF Then
      Sw = True
      rs.AddNew
      rs!CCosto = CCosto.Txcodigo.Text
      rs!CodUpr = TxcodUPR
      cDni1 = "": cDni2 = "": cDni3 = ""
   End If
   rs!Dni1 = Trim(Txdni1)
   rs!editor1 = "S"
   rs!Dni2 = Trim(Txdni2)
   rs!editor2 = "S"
   rs!dni3 = Trim(Txdni3)
   rs!editor3 = "S"
   rs!usuario = wNick         'wusuario Firma Usuario
   rs!fecha = Date
   rs!hora = Format(Time, "hh:mm:ss")
   rs!estacion = w_clienName
   rs.Update
   rs.Close
   If Sw Then
      DD.AddNew
      DD!CodUpr = TxcodUPR
   End If
   DD!Dni1 = IIf(IsNull(Txdni1), "", Txdni1)
   DD!Dni2 = IIf(IsNull(Txdni2), "", Txdni2)
   DD!dni3 = IIf(IsNull(Txdni3), "", Txdni3)
   Set rs = New Recordset
   rs.Open "SELECT * FROM CpUndPrgRol where codupr='" & TxcodUPR & "'", Cnx, adOpenStatic, adLockReadOnly
   If Not rs.EOF Then
      DD!descripcion = Trim(rs!descripcion)
   End If
   DD.Update
   If Trim(cDni1) <> Trim(Txdni1) Then
      Encclave = ""
      WVencripta (Trim(Txdni1))
      sql = "UPDATE CPGeneral SET clave='" + Encclave + "', ActClave='SI' WHERE  (baja is NULL) and Documento='" + Trim(Txdni1) + "'"
      Cnx.Execute sql
   End If
   If Trim(cDni2) <> Trim(Txdni2) Then
      Encclave = ""
      WVencripta (Trim(Txdni2))
      sql = "UPDATE CPGeneral SET clave='" + Encclave + "', ActClave='SI' WHERE  (baja is NULL) and Documento='" + Trim(Txdni2) + "'"
      Cnx.Execute sql
   End If
   If Trim(cDni3) <> Trim(Txdni3) Then
      Encclave = ""
      WVencripta (Trim(Txdni3))
      sql = "UPDATE CPGeneral SET clave='" + Encclave + "', ActClave='SI' WHERE  (baja is NULL) and Documento='" + Trim(Txdni3) + "'"
      Cnx.Execute sql
   End If
   Cnx.CommitTrans
   rs.Close
      
      
   '--Update BD ZKPUSH.User_Login(username,pass,name,servicio,tipologin,isAdmin)
   Dim cnx2  As ADODB.Connection
   Dim cDSource2, cICatal2, CString2, cUsr2, cPwd2 As String
   Dim c_username, c_pass, c_name, cTipoRol As String
   Dim b_isadmin As Boolean
   Dim valoradmin As Integer
   '--OJO: SERVIDOR de COMEDOR
   cDSource2 = "192.168.1.198"    '"192.168.1.198"      "192.168.1.15"
   cICatal2 = "ZKPUSH"
   cUsr2 = "sa"                  'zk     'sa
   cPwd2 = "P4sscomedor"            '"zk"        '"P4ssw0rd"     '
    
   CString2 = "Provider=" + cProvider + ";Data Source='" + cDSource2 + "';" & _
       "initial catalog=" + cICatal2 + "; user id='" + cUsr2 + "'; password='" + cPwd2 + "';"
   Set cnx2 = New ADODB.Connection
   With cnx2
       .ConnectionString = CString2
       .Open
   End With
   cnx2.CommandTimeout = 300
   c_username = "": c_tipoRol = "PROGR":           b_isadmin = False
   For i = 1 To 3
   
      Select Case i
      Case 1
         If Len(Trim(Txdni1)) > 0 And Trim(TxEst1) = "A" Then    '  Trim(cDni1) <> Trim(Txdni1) and
            c_username = Trim(Txdni1)
            c_pass = Format(TxFn1, "dd/mm/yyyy")
            c_name = UCase(Trim(Lblusu1.Caption))
         End If
      Case 2
         If Len(Trim(Txdni2)) > 0 And Trim(TxEst2) = "A" Then      ' Trim(cDni2) <> Trim(Txdni2) And
            c_username = Trim(Txdni2)
            c_pass = Format(Txfn2, "dd/mm/yyyy")
            c_name = UCase(Trim(Lblusu2.Caption))
         End If
      Case 3
         If Len(Trim(Txdni3)) > 0 And Trim(TxEst3) = "A" Then      'Trim(cDni3) <> Trim(Txdni3) And
            c_username = Trim(Txdni3)
            c_pass = Format(Txfn3, "dd/mm/yyyy")
            c_name = UCase(Trim(Lblusu3.Caption))
         End If
      End Select
      ' Convertir Boolean a 1 o 0
      valoradmin = IIf(b_isadmin, 1, 0)
      If Trim(c_username) <> "" Then
         sql = "SELECT * FROM UserLogin  WHERE Username='" + c_username + "'"
         rs.Open sql, cnx2, adOpenStatic, adLockOptimistic
         If rs.EOF Then
            sql = "INSERT INTO UserLogin( Username,pass,name,cCosto,tipoRol,isadmin) " & _
            " values('" + c_username + "','" + c_pass + "','" & c_name & "','" + Trim(CCosto.Txcodigo) + "'," & _
            "'" & c_tipoRol & "'," & valoradmin & ")"
         Else
            ' Usuario existente -> UPDATE
            sql = "UPDATE UserLogin SET " & _
            "Pass='" & c_pass & "', " & _
            "Name='" & c_name & "', " & _
               "cCosto='" & CCosto.Txcodigo & "', " & _
            "TipoRol='" & c_tipoRol & "', " & _
            "IsAdmin=" & valoradmin & ", " & _
            "FechaActualizacion='" & Format(Now, "yyyy-mm-dd HH:mm:ss") & "' " & _
            "WHERE Username='" & c_username & "'"
         End If
         rs.Close
         cnx2.Execute sql
      End If
      c_username = "":   c_pass = "":  c_name = ""
   Next i
   
   Set cnx2 = Nothing
   GoTo Final1
final:
   MsgBox Err.Description & " Error: " & Err.Number & vbCr
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
    If IsNull(TxcodUPR) Then Exit Sub
    If MsgBox("Eliminar  registro ?      ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
       Cnx.Execute "DELETE FROM CpCcostoupr WHERE ccosto ='" & CCosto.Txcodigo & "' AND codupr='" & TxcodUPR & "'"
       DD.Delete
    End If
    Call Limpiacampos
End Sub
Private Sub Limpiacampos()
   TxcodUPR = "": CboUpr.Text = "": Txdni1 = "": TxEst1 = ""
   Txdni2 = "": TxEst2 = "":  Txdni3 = "": TxEst3 = ""
   TxFn1 = "": Txfn2 = "":  Txfn3 = ""
   Lblusu1.Caption = "": Lblusu2.Caption = "": Lblusu3.Caption = "":
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
Private Sub CboUPR_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboUpr.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboUPR_LostFocus()
  If IsEmpty(CboUpr.Text) Then
  Else
     j = CboUpr.ListIndex
     If j <> -1 Then
        nposicion = InStr(CboUpr.List(j), ":")
        nlen = Len(CboUpr.List(j))
        TxcodUPR = Right(CboUpr.List(j), nlen - nposicion)
     End If
     Exit Sub
  End If
End Sub
Private Sub Txdni1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(Txdni1) = "" Then
        Load AyudaDni
        AyudaDni.Tag = 1
        AyudaDni.Show 1
     Else
        Set rs = New Recordset
        sql = "SELECT * FROM Cpgeneral where (baja is NULL) and documento='" & Txdni1 & "'"
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If rs.EOF Then
           MsgBox "Documento no Ubicado.."
        Else
           Lblusu1.Caption = UCase(Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres) + " ", 45))
           TxEst1 = rs!estado
           If IsDate(rs!fecnac) Then TxFn1 = rs!fecnac
        End If
       'DataGrid1.SetFocus
     End If
     If TxEst1 <> "A" Then
        MsgBox "Operador No esta Activo...", vbExclamation
        Txdni1 = ""
        Lblusu1.Caption = ""
        TxEst1 = "": TxFn1 = ""
     End If
  End If
End Sub
Private Sub Txdni2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(Txdni2) = "" Then
        Load AyudaDni
        AyudaDni.Tag = 2
        AyudaDni.Show 1
     Else
        Set rs = New Recordset
        sql = "SELECT * FROM Cpgeneral where  (baja is NULL) and documento='" & Txdni2 & "'"
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If rs.EOF Then
           MsgBox "Documento no Ubicado.."
        Else
           Lblusu2.Caption = UCase(Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres) + " ", 45))
           TxEst2 = rs!estado
           If IsDate(rs!fecnac) Then Txfn2 = rs!fecnac
        End If
       'DataGrid1.SetFocus
     End If
     If TxEst2 <> "A" Then
        MsgBox "Operador No esta Activo...", vbExclamation
        Txdni2 = ""
        Lblusu2.Caption = ""
        TxEst2 = "": Txfn2 = ""
     End If
  End If
End Sub
Private Sub Txdni3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(Txdni3) = "" Then
        Load AyudaDni
        AyudaDni.Tag = 3
        AyudaDni.Show 1
     Else
        Set rs = New Recordset
        sql = "SELECT * FROM Cpgeneral where (baja is NULL) and  documento='" & Txdni3 & "'"
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If rs.EOF Then
           MsgBox "Documento no Ubicado.."
        Else
           Lblusu3.Caption = UCase(Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres) + " ", 45))
           TxEst3 = rs!estado
           If IsDate(rs!fecnac) Then Txfn3 = rs!fecnac
        End If
       'DataGrid1.SetFocus
     End If
     If TxEst3 <> "A" Then
        MsgBox "Operador No esta Activo...", vbExclamation
        Txdni3 = ""
        Lblusu3.Caption = ""
        TxEst3 = "": Txfn3 = ""
     End If
  End If
End Sub
