VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CCosto 
   Caption         =   "Unidades de la organización"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "CCosto.frx":0000
   ScaleHeight     =   8415
   ScaleMode       =   0  'User
   ScaleWidth      =   11057.53
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "[ Programadores  del periodo]"
      Height          =   975
      Left            =   4680
      TabIndex        =   17
      Top             =   7200
      Width           =   3375
      Begin VB.CommandButton CmdExportaXLS 
         Height          =   615
         Left            =   2160
         Picture         =   "CCosto.frx":80266
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Exporta a Xls"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxFechaIni 
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
         Left            =   840
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxFechaFin 
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
         Left            =   840
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Del:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Al:"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdUndPrgRol 
      Caption         =   "Unidad Programar Roles"
      Height          =   615
      Left            =   3000
      TabIndex        =   15
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   0
      Picture         =   "CCosto.frx":80EA8
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton CmdLnkHorario 
      Caption         =   "Horario Asistencial"
      Height          =   615
      Left            =   960
      TabIndex        =   13
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox txBuscar 
      Height          =   285
      Left            =   4800
      MaxLength       =   30
      TabIndex        =   11
      Top             =   4560
      Width           =   4455
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   9600
      Picture         =   "CCosto.frx":81AEA
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cerrar"
      Top             =   7440
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   1680
      TabIndex        =   5
      Top             =   5280
      Width           =   6735
      Begin VB.TextBox Txcodigo 
         Height          =   285
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Txdescripcion 
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   6
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   840
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
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Cmdnuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   840
      TabIndex        =   0
      Top             =   840
      Width           =   9615
      _ExtentX        =   16960
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
   Begin VB.Label LblDS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LblDS"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   9000
      TabIndex        =   16
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Denominacion a ubicar :"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1680
      TabIndex        =   12
      Top             =   4560
      Width           =   1725
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   8355
   End
End
Attribute VB_Name = "CCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim Sw As Boolean

Private Sub Command1_Click()
     c_grupo = "01": c_secuencia = "012"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub

Private Sub Form_Load()
    LblDS.Caption = cICatal + "-" + UCase(IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource)))
    Set DD = New Recordset
    With DD
     .Fields.Append "codigo", adVarChar, 2
     .Fields.Append "descripcion", adVarChar, 30
     .Open
    End With
    Set DataGrid1.DataSource = DD
    'Configura columnas
    For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
         Case 0
         .Caption = "Codigo"
          .Width = Me.TextWidth(String(6, "S"))
         Case 1
          .Caption = "Descripcion"
          .Width = Me.TextWidth(String(30, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    
    Set rs = New Recordset
    sql = "SELECT * FROM CpCcosto order by Descripcion"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!Codigo = rs!Codigo
       DD!descripcion = rs!descripcion
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
    DD.MoveFirst
End Sub
Private Sub DataGrid1_Click()
   Txcodigo.Text = DataGrid1.Columns(0)
   Txdescripcion.Text = DataGrid1.Columns(1)
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
     Txcodigo.SetFocus
     CmdEliminar.Enabled = False
     Sw = True
     Cmdnuevo.Caption = "&Grabar"
     CmdEditar.Caption = "&Cancelar"
     CmdEditar.Enabled = True
  Else                  'Grabar
     On Error GoTo err_trans
     If Trim(Txcodigo.Text) = "" Then
        MsgBox ("Favor Verifique Informacion")
        Exit Sub
     End If
     Cnx.BeginTrans
     Dim sql As String
     If Sw Then
        sql = "INSERT INTO CpCcosto(codigo,descripcion,usuario,fecha) "
        sql = sql + " VALUES('" + Txcodigo + "','" + Trim(Txdescripcion) + "','" + wUsuario + "','" + Format(Date, "YYYYMMDD") + "')"
     Else
        sql = "UPDATE CpCcosto SET descripcion='" + Trim(Txdescripcion) + "',usuario='" + wUsuario + "', fecha='" + Format(Date, "YYYYMMDD") + "'"
        sql = sql + " WHERE codigo='" + Txcodigo + "'"
     End If
     Cnx.Execute sql
    Cnx.CommitTrans
    If Sw Then
       DD.AddNew
       DD!Codigo = Txcodigo.Text
    End If
    DD!descripcion = Txdescripcion.Text
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
    Txcodigo.Locked = True
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
       Cnx.Execute "DELETE FROM CpCcosto WHERE codigo='" & Txcodigo.Text & "'"
       DD.Delete
       Limpiacampos
    End If
End Sub

Private Sub cmdsalir_Click()
  Set rs = Nothing
  Set rs_d = Nothing
  Set rt = Nothing
  For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "CCosto" Then
           Unload Forms(cont)
        End If
      End If
  Next

  Unload Me
  Me.Hide
End Sub
Sub habilitatextos()
  Txcodigo.Locked = False
  Txdescripcion.Locked = False
End Sub
Sub desabilitatextos()
  Txcodigo.Locked = True
  Txdescripcion.Locked = True
End Sub
Sub Limpiacampos()
  Txcodigo.Text = ""
  Txdescripcion.Text = ""
End Sub
Private Sub TxBuscar_Change()
    Dim cBuscar As String
    cpuntero = DataGrid1.Bookmark
    If txBuscar.Text = "" Then Exit Sub
    cBuscar = "Descripcion like '%" & txBuscar.Text & "%'"
    DD.MoveFirst
    DD.Find (cBuscar)
    If DD.EOF Then
       txBuscar.BackColor = &HFF&
    Else
       txBuscar.BackColor = &H80000005
       DataGrid1.Refresh
    End If
End Sub
Private Sub Txcodigo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     'Txcodigo.Text = Right("00" + Trim(Str(Val(Txcodigo.Text), 2)), 2)
     Txcodigo.Text = Right("00" + Trim(Txcodigo.Text), 2)
     Txdescripcion.SetFocus
     Exit Sub
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
       KeyAscii = 0
   End If
End Sub
Private Sub Txdescripcion_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Txdescripcion.Text = UCase(Txdescripcion.Text)
     Cmdnuevo.SetFocus
     Exit Sub
   End If
End Sub
Private Sub CmdLnkHorario_Click()
 If Trim(Txcodigo.Text) = "" Then Exit Sub
 Load cCostoHor
 cCostoHor.Left = 2760: cCostoHor.Top = 5070
 cCostoHor.Refresh
 cCostoHor.Show
End Sub

Private Sub CmdUndPrgRol_Click()
  If Trim(Txcodigo.Text) = "" Then Exit Sub
  Load cCostoUpr
  cCostoUpr.Left = 1860: cCostoUpr.Top = 3420
  cCostoUpr.Refresh
  cCostoUpr.Show
End Sub

Private Sub CmdExportaXLS_Click()
   Dim anoi, anof, mesi, mesf As String
   Criterio = ""
   If TxFechaIni <> "" And TxFechaFin <> "" Then
      anoi = Left(Format(TxFechaIni, "YYYYMMDD"), 4)
      anof = Left(Format(TxFechaFin, "YYYYMMDD"), 4)
      mesi = Mid(Format(TxFechaIni, "YYYYMMDD"), 5, 2)
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
      If Trim(rs!CodUpr) = "" Or (rs!ANO = anoi And rs!MES < mesi) Or (rs!ANO = anof And rs!MES > mesf) Or rs!Flag <> "0" Then
      Else
         cCodUpr = rs!CodUpr
         cano = rs!ANO
         cmes = rs!MES
         CUSUARIO = rs!USUARIO
         Sq0 = "SELECT *  FROM XPROGuPR WHERE CODUPR='" + cCodUpr + "' AND ANO='" + cano + "' AND USUARIO='" + CUSUARIO + "'"
         rs_d.Open Sq0, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If rs_d.EOF() Then
            rs_d.AddNew
            rs_d!CodUpr = rs!CodUpr
            Set rt = New Recordset
            rt.Open "SELECT * FROM CpUndPrgRol WHERE codUpr='" & rs!CodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
            If Not rt.EOF Then
               rs_d!DscUpr = Left(rt!descripcion, 50)
            End If
            rt.Close
            rs_d!ANO = rs!ANO
            rs_d!del_MES = rs!MES
            rs_d!USUARIO = rs!USUARIO
            cdocum = rs!USUARIO
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
            If rs_d!del_MES > rs!MES Then
               rs_d!Al_mes = rs_d!del_MES
               rs_d!del_MES = rs!MES
            Else
               If rs_d!Al_mes > rs!MES Then
               Else
                  rs_d!Al_mes = rs!MES
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
   

