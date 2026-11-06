VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ChkTarjeta 
   Caption         =   "Actualiza Tarjeta del Servidor"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "ChkTarjeta.frx":0000
   ScaleHeight     =   6390
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   8880
      Picture         =   "ChkTarjeta.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Limpia plantilla"
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   0
      Picture         =   "ChkTarjeta.frx":80B30
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Txbuscar 
      Height          =   285
      Left            =   6000
      MaxLength       =   5
      TabIndex        =   14
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Txbuscar1 
      Height          =   285
      Left            =   8400
      MaxLength       =   15
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   9960
      Picture         =   "ChkTarjeta.frx":81772
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Cerrar"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton CmdProcesa 
      Caption         =   "Procesar Cambios"
      Height          =   495
      Left            =   9720
      TabIndex        =   9
      Top             =   4800
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   600
      TabIndex        =   1
      Top             =   3360
      Width           =   9015
      Begin VB.CommandButton CmdGraba 
         Caption         =   "GRABA"
         Height          =   495
         Left            =   6480
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox TxObserva 
         Height          =   285
         Left            =   1680
         MaxLength       =   40
         TabIndex        =   7
         Top             =   960
         Width           =   6495
      End
      Begin VB.TextBox TxTarjeta 
         Height          =   285
         Left            =   600
         MaxLength       =   5
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox TxTarjetaACT 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         MaxLength       =   5
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Txservidor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         MaxLength       =   40
         TabIndex        =   2
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label8 
         Caption         =   "Descripcion"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Tarjeta ACTUAL"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Observacion"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Tarjeta NUEVA"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DtgValida 
      Height          =   2535
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4471
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
      Caption         =   "Maestra de Servidores"
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicar       Tarjeta"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Apel. Paterno"
      Height          =   255
      Left            =   7320
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "ChkTarjeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim cTarjeta, cTarjetaOLD, cobserva As String
Dim nposicion As Integer

Private Sub Command1_Click()
    c_grupo = "05": c_secuencia = "056"
    Load AydAcerca
    AydAcerca.Left = 0: AydAcerca.Top = 0
    AydAcerca.Show
End Sub

Private Sub Form_Load()
    Set DD = New Recordset
    With DD
      .Fields.Append "TarjetaACT", adVarChar, 5
      .Fields.Append "Tarjeta", adVarChar, 5
      .Fields.Append "Servidor", adVarChar, 40
      .Fields.Append "Observa", adVarChar, 40
      .Open
    End With
    Set DtgValida.DataSource = DD
    For i = 0 To DtgValida.Columns.Count - 1
       With DtgValida.Columns(i)
       Select Case i
         Case 0
          '.Caption = "Codigo"
          .Width = Me.TextWidth(String(9, "S"))
         Case 1
          '.Caption = "Codigo"
          .Width = Me.TextWidth(String(9, "S"))
         Case 2
          '.Caption = "Codigo"
          .Width = Me.TextWidth(String(40, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
    Next
    Set rs = New Recordset
    sql = " SELECT * FROM CPGeneral WHERE  (baja is NULL) and  estado='A' ORDER BY apelpat, apelmat"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    nposicion = 0
    Do While Not rs.EOF
       'verifica pendientes de Actualizar
       cTarjetaOLD = rs!tarjeta
       cTarjeta = ""
       cobserva = ""
       Set rt = New Recordset
       sql = "SELECT * FROM CpChgTarje WHERE TarjetaOLD='" & cTarjetaOLD & "' and ( flg is null or FLG='X')"
       rt.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       If Not rt.EOF Then
          cTarjeta = IIf(IsNull(rt!tarjeta), "", rt!tarjeta)
          cobserva = IIf(IsNull(rt!observa), "", rt!observa)
       End If
       rt.Close
       DD.AddNew
       DD!TarjetaACT = rs!tarjeta
       DD!tarjeta = cTarjeta
       DD!servidor = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 40)
       DD!observa = cobserva
       DD.Update
       nposicion = nposicion + 1
       rs.MoveNext
    Loop
    End If
    rs.Close
    DD.MoveFirst
    DtgValida.Refresh
End Sub
Private Sub TxBUscar_KeyPress(KeyAscii As Integer)
  If nposicion = 0 Then Exit Sub
  If KeyAscii = 13 Then
     Dim cBuscar As String
     cpuntero = DtgValida.Bookmark
     If Txbuscar.Text = "" Then Exit Sub
     cBuscar = "tarjetaACT like '" & Trim(Txbuscar.Text) & "*'"
     DD.MoveFirst
     DD.Find (cBuscar)
     If DD.EOF Then
        Txbuscar.BackColor = &HFF&
     Else
        Txbuscar.BackColor = &H80000005
        DtgValida.Refresh
        Txbuscar.Text = ""
        DtgValida_Click
     End If
  End If
End Sub
Private Sub Txbuscar1_KeyPress(KeyAscii As Integer)
   If nposicion = 0 Then Exit Sub
   If KeyAscii = 13 Then
     Dim cBuscar As String
     cpuntero = DtgValida.Bookmark
     If Txbuscar1.Text = "" Then Exit Sub
     cBuscar = "Servidor like '" & Trim(Txbuscar1.Text) & "*'"
     DD.MoveFirst
     DD.Find (cBuscar)
     If DD.EOF Then
        Txbuscar1.BackColor = &HFF&
     Else
        Txbuscar1.BackColor = &H80000005
        DtgValida.Refresh
        Txbuscar1.Text = ""
        DtgValida_Click
     End If
  End If
End Sub

Private Sub DtgValida_Click()
   Call LimpiarCampos
   TxTarjetaACT.Text = DtgValida.Columns(0)
   Txtarjeta.Text = DtgValida.Columns(1)
   Txservidor.Text = DtgValida.Columns(2)
   TxObserva.Text = DtgValida.Columns(3)
End Sub
Private Sub DtGvalida_HeadClick(ByVal ColIndex As Integer)
  With DD
    If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
       .Sort = .Fields(ColIndex).[Name] & " Desc"
    Else
       .Sort = .Fields(ColIndex).[Name] & " Asc"
    End If
  End With
  Call LimpiarCampos
End Sub
Private Sub LimpiarCampos()
    Txtarjeta.Text = "": TxTarjetaACT.Text = "": Txservidor.Text = "":     TxObserva.Text = ""
    Txbuscar.Text = "": Txbuscar1.Text = ""
End Sub
Private Sub Txtarjeta_KeyPress(KeyAscii As Integer)
   If Txtarjeta.Text = "" Or Txtarjeta.Text = TxTarjetaACT.Text Then Exit Sub
   If KeyAscii = 13 Then
      Txtarjeta.Text = UCase(Txtarjeta.Text)
      Set rs = New Recordset
      sql = "SELECT * FROM CpGeneral WHERE  (baja is NULL) And Tarjeta='" & Txtarjeta.Text & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then
         'Tarjeta NUEVA existe
         cNombre = Left(Trim(rs!apelpat) + " " + Trim(rs!apelpat) + " " + Trim(rs!nombres), 35)
         MsgBox "Tarjeta Pertenece a.. " + cNombre, vbExclamation
         Txtarjeta.Text = ""
         rs.Close
         Exit Sub
      End If
      rs.Close
      Set rs = New Recordset
      sql = "SELECT * FROM CpChgTarje WHERE Tarjeta='" & Txtarjeta.Text & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs.EOF Then
         If rs!FLG = "X'" Then
            MsgBox ("Cambio Ya PROCESADO..!")
         Else
            MsgBox ("Cambio Ya REGISTRADO..!")
         End If
         Txtarjeta.Text = ""
         rs.Close
         Exit Sub
      End If
      rs.Close
      TxObserva.SetFocus
   End If
End Sub
Private Sub TxObserva_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxObserva.Text = UCase(TxObserva.Text)
      cmdGraba.SetFocus
   End If
End Sub

Private Sub cmdGraba_Click()
      If Txtarjeta.Text = TxTarjetaACT.Text Then
         MsgBox "<E> Tarjeta .. ", vbExclamation
         Call LimpiarCampos
         DtgValida.SetFocus
         Exit Sub
      End If
      If Txtarjeta.Text = "" Then
         'Anula cambio de existir Record
      Else
         Set rs = New Recordset
         sql = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and Tarjeta='" & Txtarjeta.Text & "'"
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then
            cNombre = Left(Trim(rs!apelpat) + " " + Trim(rs!apelpat) + " " + Trim(rs!nombres), 35)
            MsgBox "<E> Tarjeta Pertenece a... " + cNombre, vbExclamation
            Txtarjeta.Text = ""
            rs.Close
            Exit Sub
         End If
         rs.Close
         Set rs = New Recordset
         sql = "SELECT * FROM CpChgTarje WHERE Tarjeta='" & Txtarjeta.Text & "'"
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then
            If rs!FLG <> "X" Then
               MsgBox "<E> Ya Registro Cambio..!", vbExclamation
            End If
            Txservidor.Text = rs!TarjetaOld + " " + Left(rs!servidor, 34)
            rs.Close
            Exit Sub
         End If
         rs.Close
      End If
      'Set rs = New Recordset
      'sql = "SELECT * FROM CpChgTarje WHERE TarjetaOLD='" & TxTarjetaACT.Text & "' and FLG='X'"
      'rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      'If Not rs.EOF Then
      '   MsgBox "<E> Tarjeta Origen, PROCESADA..!", vbExclamation
      '   rs.Close
      '   Exit Sub
      'End If
      'rs.Close
      If Txtarjeta.Text = "" Then
         Cnx.Execute "DELETE FROM CpChgTarje WHERE tarjetaOld ='" & TxTarjetaACT.Text & "'"
         DD.Delete
      Else
         Set rs = New Recordset
         sql = "SELECT * FROM CpChgTarje WHERE TarjetaOLD='" & TxTarjetaACT.Text & "' and (FLG is null or FLG<>'X')"
         rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
         If rs.EOF Then
            rs.AddNew
            rs!tarjeta = UCase(Txtarjeta.Text)
            rs!servidor = IIf(IsNull(Txservidor.Text), "", Txservidor.Text)
            rs!TarjetaOld = UCase(TxTarjetaACT.Text)
         End If
         rs!observa = TxObserva.Text
         rs!fecha = Date
         If Len(wTerminal) = 0 Then    'Nick Usuario+ pto acceso cliente
           rs!usuario = Trim(Left(wNick, 10))
         ElseIf Len(wTerminal) = 3 Then    'Nick Usuario+ pto acceso cliente
                rs!usuario = Trim(Left(wNick, 7)) + wTerminal
         ElseIf Len(wTerminal) = 2 Then
                rs!usuario = Trim(Left(wNick, 8)) + wTerminal
         ElseIf Len(wTerminal) = 1 Then
                rs!usuario = Trim(Left(wNick, 9)) + wTerminal
         End If
         rs.Update
         DD!tarjeta = UCase(Txtarjeta.Text)
         DD!observa = IIf(IsNull(TxObserva.Text), "", TxObserva.Text)
         DD.Update
      End If
      Call LimpiarCampos
End Sub

Private Sub CmdClear_Click()
   Call LimpiarCampos
   DtgValida.SetFocus
End Sub

Private Sub CmdProcesa_Click()
   Call LimpiarCampos
   Set rs = New Recordset
   sql = "SELECT * FROM CpChgTarje WHERE (FLG is null or FLG<>'X') "
   rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
   rs.MoveFirst
   Do While Not rs.EOF
      If Trim(rs!tarjeta) = "" Or Trim(rs!TarjetaOld) = "" Then
      Else
         If Trim(rs!tarjeta) <> Trim(rs!TarjetaOld) Then
            cTarjeta = rs!tarjeta
            cTarjetaOLD = rs!TarjetaOld
            'Generales
            sql = "UPDATE CPGeneral SET Tarjeta='" & cTarjeta & "' WHERE  (baja is NULL) and tarjeta='" & Trim(cTarjetaOLD) & "'"
            Cnx.Execute sql
            sql = "UPDATE CPDirectorio SET Tarjeta='" & cTarjeta & "' WHERE  tarjeta='" & Trim(cTarjetaOLD) & "'"
            Cnx.Execute sql
            sql = "UPDATE CPGrupos SET Tarjeta='" & cTarjeta & "' WHERE tarjeta='" & Trim(cTarjetaOLD) & "'"
            Cnx.Execute sql
            sql = "UPDATE CPHor" & genaño & " SET Tarjeta='" & cTarjeta & "' WHERE tarjeta='" & Trim(cTarjetaOLD) & "'"
            Cnx.Execute sql
            'Marcacion
            sql = "UPDATE CPMarca" & genaño & " SET Tarjeta='" & cTarjeta & "' WHERE tarjeta='" & Trim(cTarjetaOLD) & "'"
            Cnx.Execute sql
            '-hasta 2013 sql = "UPDATE CPManual" & genaño & " SET Tarjeta='" & cTarjeta & "' WHERE tarjeta='" & Trim(cTarjetaOLD) & "'"
            'Cnx.Execute sql
            'Autorizaciones/Motivos
            sql = "UPDATE CPAutori" & genaño & " SET Tarjeta='" & cTarjeta & " ' WHERE tarjeta='" & Trim(cTarjetaOLD) & "'"
            Cnx.Execute sql
            sql = "UPDATE CPChgTur" & genaño & " SET Tarjeta='" & cTarjeta & " ' WHERE tarjeta='" & Trim(cTarjetaOLD) & "'"
            Cnx.Execute sql
            sql = "UPDATE CPChgTurM" & genaño & " SET Tarjeta='" & cTarjeta & " ' WHERE tarjeta='" & Trim(cTarjetaOLD) & "'"
            Cnx.Execute sql
            rs!FLG = "X"
            rs.Update
         End If
      End If
      rs.MoveNext
   Loop
   rs.Close
   Call cmdsalir_Click
End Sub
Private Sub cmdsalir_Click()
   Set rs = Nothing
   Set DD = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "ChkTarjeta" Then
           Unload Forms(cont)
        End If
      End If
   Next
   
   Unload Me
   Me.Hide
End Sub

