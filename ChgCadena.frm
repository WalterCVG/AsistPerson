VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ChgCadena 
   Caption         =   "Cambiar Cadena"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "ChgCadena.frx":0000
   ScaleHeight     =   5940
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxUbiCadena 
      Height          =   285
      Left            =   2160
      MaxLength       =   35
      TabIndex        =   1
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   10680
      Begin VB.TextBox TxDescripcionO 
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         MaxLength       =   110
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.ComboBox CboCadenaO 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   8655
      End
      Begin VB.TextBox TxDescripcionD 
         Enabled         =   0   'False
         Height          =   195
         Left            =   3360
         MaxLength       =   150
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.ComboBox CboCadenaD 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   840
         Width           =   8655
      End
      Begin VB.TextBox TxCadenaD 
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   52
         TabIndex        =   7
         Top             =   1560
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox TxCadenaO 
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         MaxLength       =   35
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton CmdGraba 
         Caption         =   "GRABA"
         Height          =   495
         Left            =   9600
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Cadena Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdCambio 
      Caption         =   "Ejecuta CAMBIOS"
      Height          =   495
      Left            =   9120
      TabIndex        =   13
      Top             =   5160
      Width           =   1680
   End
   Begin MSDataGridLib.DataGrid DtgValida 
      Height          =   2535
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   9855
      _ExtentX        =   17383
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
      Caption         =   "Registro de  Cambios"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cadena Actual"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "ChgCadena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim ccadena, cCadenaO, cCadenaD As String
Dim nposicion As Integer
Private Sub Form_Load()
    '-ACTUAL
    Set rs = New Recordset
    rs.Open "SELECT * FROM cpCadHep  ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
    rs.MoveFirst
    While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboCadenaO.ListCount - 1
           If CboCadenaO.List(i) = rs.Fields("Descripcion") + Space(80) + ":" + rs.Fields("cadhep") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboCadenaO.AddItem rs.Fields("Descripcion") + Space(80) + ":" + rs.Fields("cadhep")
       End If
       rs.MoveNext
    Wend
    rs.Close
    '-NUEVO
    Set rs = New Recordset
    rs.Open "SELECT * FROM cpCadena  ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
    rs.MoveFirst
    While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboCadenaD.ListCount - 1
           If CboCadenaD.List(i) = rs.Fields("Descripcion") + Space(80) + ":" + rs.Fields("cadena") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboCadenaD.AddItem rs.Fields("Descripcion") + Space(80) + ":" + rs.Fields("cadena")
       End If
       rs.MoveNext
    Wend
    rs.Close
  
    Set DD = New Recordset
    With DD
      .Fields.Append "CadenaO", adVarChar, 35
      .Fields.Append "CadenaD", adVarChar, 52
      .Fields.Append "descripcion", adVarChar, 50
      .Fields.Append "Flg", adVarChar, 1
      .Fields.Append "Usuario", adVarChar, 10
      .Fields.Append "fecha", adDate
      .Open
    End With
    Set DtgValida.DataSource = DD
    For i = 0 To DtgValida.Columns.Count - 1
       With DtgValida.Columns(i)
         Select Case i
           Case 0
             .Caption = "Actual"
             .Width = Me.TextWidth(String(35, "S"))
           Case 1
             .Caption = "Destino"
             .Width = Me.TextWidth(String(52, "S"))
           Case 2
             '.Caption = "descripcion"
             .Width = Me.TextWidth(String(50, "S"))
           Case 3
             '.Caption = "Flg"
             .Width = Me.TextWidth(String(3, "S"))
           Case 4
             '.Caption = "usuario"
             .Width = Me.TextWidth(String(10, "S"))
           Case 5
             '.Caption = "fecha"
             .Width = Me.TextWidth(String(10, "S"))
         End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
       End With
    Next
    Set rs = New Recordset
    sql = "SELECT * FROM CpChgCad WHERE ano='" & genaño & "' AND ( flg is null )"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
    rs.MoveFirst
    Do While Not rs.EOF
       DD.AddNew
       DD!cadenaO = IIf(IsNull(rs!cadenaO), "", rs!cadenaO)
       DD!cadenaD = IIf(IsNull(rs!cadenaD), "", rs!cadenaD)
       DD!descripcion = IIf(IsNull(rs!descripcionD), "", rs!descripcionD)
       DD!FLG = IIf(IsNull(rs!FLG), "", rs!FLG)
       DD!usuario = IIf(IsNull(rs!usuario), "", rs!usuario)
       If IsDate(rs!fecha) Then DD!fecha = rs!fecha
       DD.Update
       rs.MoveNext
    Loop
    End If
    rs.Close
End Sub

Private Sub TxUbiCadena_Change()
    Dim cBuscar As String
    cpuntero = DtgValida.Bookmark
    If TxUbiCadena.Text = "" Then Exit Sub
    cBuscar = "CadenaO like  '" & TxUbiCadena.Text & "*'"
    DD.MoveFirst
    DD.Find (cBuscar)
    If DD.EOF Then
       TxUbiCadena.BackColor = &HFF&
    Else
       TxUbiCadena.BackColor = &H80000005
       DtgValida.Refresh
    End If
    Call Limpiacampos
End Sub
Private Sub CboCadenaO_GotFocus()
    Const CB_SHOWDROPDOWN = &H14F
    Dim Tmp
    Tmp = SendMessage(CboCadenaO.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboCadenaO_LostFocus()
  If IsNull(CboCadenaO.Text) Then
  Else
     j = CboCadenaO.ListIndex
     nposicion = InStr(CboCadenaO.List(j), ":")
     nlen = Len(CboCadenaO.List(j))
     TxCadenaO.Text = Right(CboCadenaO.List(j), nlen - nposicion)
     Set rs = New Recordset
     rs.Open "SELECT * FROM cpCadHep  WHERE  cadhep='" & TxCadenaO & "'", Cnx, adOpenStatic, adLockReadOnly
     If rs.EOF Then
        MsgBox ("Error cad. Actual..."), vbExclamation
        CboCadenaO.Text = "": TxCadenaO.Text = "": TxDescripcionO.Text = ""
        rs.Close
        Exit Sub
     End If
     TxDescripcionO = rs!descripcion
     rs.Close
  End If
  CboCadenaD.SetFocus
End Sub

Private Sub CboCadenaD_GotFocus()
    Const CB_SHOWDROPDOWN = &H14F
    Dim Tmp
    Tmp = SendMessage(CboCadenaD.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboCadenaD_LostFocus()
  If IsNull(CboCadenaD.Text) Then
  Else
     j = CboCadenaD.ListIndex
     nposicion = InStr(CboCadenaD.List(j), ":")
     nlen = Len(CboCadenaD.List(j))
     TxCadenaD.Text = Right(CboCadenaD.List(j), nlen - nposicion)
     Set rs = New Recordset
     rs.Open "SELECT * FROM cpCadena  WHERE  cadena='" & TxCadenaD & "'", Cnx, adOpenStatic, adLockReadOnly
     If rs.EOF Then
        MsgBox ("Error cad. NUEVA..."), vbExclamation
        CboCadenaD.Text = "": TxCadenaD.Text = "": TxDescripcionD.Text = ""
        rs.Close
        Exit Sub
     End If
     TxDescripcionD = rs!descripcion
     rs.Close
  End If
  cmdGraba.SetFocus
End Sub

Private Sub DtgValida_Click()
   Call Limpiacampos
   TxCadenaO.Text = DtgValida.Columns(0)
   TxCadenaD.Text = DtgValida.Columns(1)
   Set rs = New Recordset
   sql = "SELECT * FROM CpCadHep WHERE  cadHep='" & Trim(TxCadenaO.Text) & "'"
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      TxDescripcionO.Text = rs!descripcion
   End If
   rs.Close
   If Trim(TxCadenaO.Text) <> "" Then
      For i = 0 To CboCadenaO.ListCount - 1
         nposicion = InStr(CboCadenaO.List(i), ":")
         nlen = Len(CboCadenaO.List(i))
         dato = Trim(Right(CboCadenaO.List(i), nlen - nposicion))
         If dato = Trim(TxCadenaO.Text) Then
            CboCadenaO.ListIndex = i
            Exit For
         End If
      Next i
   End If
   
   Set rs = New Recordset
   sql = "SELECT * FROM CpCadena WHERE  cadena='" & Trim(TxCadenaD.Text) & "'"
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      TxDescripcionD.Text = rs!descripcion
   End If
   rs.Close
   If Trim(TxCadenaD.Text) <> "" Then
      For i = 0 To CboCadenaD.ListCount - 1
         nposicion = InStr(CboCadenaD.List(i), ":")
         nlen = Len(CboCadenaD.List(i))
         dato = Right(CboCadenaD.List(i), nlen - nposicion)
         If dato = Trim(TxCadenaD.Text) Then
            CboCadenaD.ListIndex = i
            Exit For
         End If
      Next i
   End If
End Sub
Private Sub DtGvalida_HeadClick(ByVal ColIndex As Integer)
   With DD
      If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
         .Sort = .Fields(ColIndex).[Name] & " Desc"
      Else
         .Sort = .Fields(ColIndex).[Name] & " Asc"
      End If
   End With
   Call Limpiacampos
End Sub
Private Sub Limpiacampos()
   TxUbiCadena.Text = ""
   TxCadenaO.Text = "": CboCadenaO.Text = "": TxDescripcionO.Text = ""
   TxCadenaD.Text = "": CboCadenaD.Text = "":  TxDescripcionD.Text = ""
End Sub
Private Sub cmdGraba_Click()
     If TxCadenaO.Text = TxCadenaD.Text Or TxCadenaD.Text = "" Then
        MsgBox ("Error campo DESTINO..."), vbExclamation
        Exit Sub
     End If
     Set rs = New Recordset
     sql = "SELECT * FROM CpCadhep  WHERE CadHep='" & Trim(TxCadenaO.Text) & "'"
     rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
     If rs.EOF Then
        MsgBox ("Error Cadena Origen..."), vbExclamation
        rs.Close
        Exit Sub
     Else
        TxDescripcionO.Text = IIf(IsNull(rs!descripcion), "", RTrim(rs!descripcion))
     End If
     rs.Close
     Set rs = New Recordset
     sql = "SELECT * FROM CpCadena  WHERE Cadena='" & Trim(TxCadenaD.Text) & "'"
     rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
     If rs.EOF Then
        MsgBox ("Error Cadena Destino..."), vbExclamation
        rs.Close
        Exit Sub
     End If
     TxDescripcionD.Text = IIf(IsNull(rs!descripcion), "", RTrim(rs!descripcion))
     rs.Close
     Sw = False
     'On Error GoTo final
     'Cnx.BeginTrans
     Set rs = New Recordset
     sql = "SELECT * FROM CpChgCad  WHERE ano='" & genaño & "' And CadenaO='" & Trim(TxCadenaO.Text) & "' and  (FLG is Null)"
     rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
     If rs.EOF Then
        Sw = True
        rs.AddNew
        rs!ano = genaño
        rs!cadenaO = TxCadenaO
     Else
        If rs!FLG = "X" Then
           MsgBox ("Error, Cambio Ya realizado..."), vbExclamation
           rs.Close
           Exit Sub
        End If
     End If
     rs!descripcionO = TxDescripcionO
     rs!cadenaD = TxCadenaD
     rs!descripcionD = TxDescripcionD
     rs!usuario = wUsuario      'Firma Usuario
     rs!fecha = Date
     rs.Update
     If Sw Then
        DD.AddNew
        DD!cadenaO = TxCadenaO.Text
        DD!cadenaD = TxCadenaD.Text
        DD!usuario = wUsuario
        DD!fecha = Date
     End If
     DD!descripcion = IIf(IsNull(TxDescripcionD.Text), "", Trim(TxDescripcionD.Text))
     DD.Update
     'Cnx.CommitTrans
    'GoTo Final1
'final:
'   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
'   Cnx.RollbackTrans
'Final1:
   Call Limpiacampos
  DtgValida.SetFocus
End Sub

Private Sub CmdCambio_Click()
   Call Limpiacampos
   DD.MoveFirst
   nposicion = 0
   Do While Not DD.EOF
      If Trim(DD!cadenaD) = "" Or Trim(DD!cadenaO) = "" Then
      Else
         If Trim(DD!cadenaO) <> Trim(DD!cadenaD) Then
            cCadenaO = Trim(DD!cadenaO)
            cCadenaD = Trim(DD!cadenaD)
            sql = "UPDATE cPGeneral SET  (baja is NULL) and OCadHep='" & cCadenaO & "' AND CadHep='" & cCadenaD & "' WHERE cadHep='" & cCadenaO & "'"
            Cnx.Execute sql
            sql = "UPDATE CpChgCad SET flg='X' WHERE Ano='" & genaño & "' AND cadenaO='" & cCadenaO & "' AND cadenaD='" & cCadenaD & "'"
            Cnx.Execute sql
            nposicion = nposicion + 1
         End If
         DD!FLG = "X"
         DD.Update
      End If
      DD.MoveNext
   Loop
   If nposicion > 0 Then
      MsgBox ("Efectuo .." + Trim(Str(nposicion)) + " cambios.")
   End If
   Call cmdsalir_Click
End Sub
Private Sub cmdsalir_Click()
   Set rs = Nothing
   Set DD = Nothing
   Unload Me
   Me.Hide
End Sub
