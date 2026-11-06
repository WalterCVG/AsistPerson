VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form QMultiple 
   Caption         =   "Consulta Multiple"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "QMultiple.frx":0000
   ScaleHeight     =   8430
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   615
      Left            =   11280
      Picture         =   "QMultiple.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Cerrar"
      Top             =   7800
      Width           =   735
   End
   Begin VB.CheckBox ChkXLS 
      Alignment       =   1  'Right Justify
      Caption         =   "XLS instalado"
      Height          =   195
      Left            =   10920
      TabIndex        =   36
      Top             =   4680
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Frame rAnalisMult 
      Caption         =   "Parametros"
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
      Height          =   4335
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   10575
      Begin VB.TextBox TxDscCargo 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   35
         Top             =   2760
         Width           =   4815
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo Activos"
         Height          =   375
         Left            =   5880
         TabIndex        =   1
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   6480
         TabIndex        =   28
         Top             =   2520
         Width           =   3975
         Begin VB.ListBox Lb1 
            Height          =   1230
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1575
         End
         Begin VB.ListBox Lb2 
            Height          =   645
            Left            =   2400
            TabIndex        =   31
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   ">"
            Height          =   375
            Left            =   1800
            TabIndex        =   30
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "<"
            Height          =   375
            Left            =   1800
            TabIndex        =   29
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Lblcampos 
            Height          =   255
            Left            =   2280
            TabIndex        =   33
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.TextBox Txcampo1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Txcampo2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   26
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Txcategoria 
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         MaxLength       =   1
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox CboRegPlh 
         Height          =   315
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox CboTipoPla 
         Height          =   315
         Left            =   8160
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox TxPlh 
         Enabled         =   0   'False
         Height          =   225
         Left            =   3360
         MaxLength       =   3
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Txtipopla 
         Enabled         =   0   'False
         Height          =   225
         Left            =   7560
         MaxLength       =   3
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox Txcargo 
         Height          =   285
         Left            =   840
         MaxLength       =   4
         TabIndex        =   19
         Top             =   2760
         Width           =   495
      End
      Begin VB.ComboBox CboUpr 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   2160
         Width           =   6015
      End
      Begin VB.TextBox Txcodupr 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox CboCCosto 
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   1560
         Width           =   6015
      End
      Begin VB.TextBox Txccosto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   15
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox Cbocategoria 
         Height          =   315
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox Txfiltro 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   3360
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.Label Label23 
         Caption         =   "Regimen Laboral"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label22 
         Caption         =   "Tipo Planilla"
         Height          =   255
         Left            =   7080
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cargo"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "U.P.R."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Und. Org."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   11040
      TabIndex        =   9
      Top             =   5760
      Width           =   1095
      Begin VB.CommandButton mnuFicImpSelec 
         Height          =   495
         Left            =   240
         Picture         =   "QMultiple.frx":803B0
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Selec.Impresora"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton CmdPrint 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         Picture         =   "QMultiple.frx":80B8E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Imprimir"
         Top             =   960
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdcarga 
      Height          =   615
      Left            =   11160
      Picture         =   "QMultiple.frx":81458
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ubicar Registros"
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton CmdClear 
      Height          =   735
      Left            =   11160
      Picture         =   "QMultiple.frx":81633
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Otra Consulta"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton CmdExportaXLS 
      Height          =   615
      Left            =   11280
      Picture         =   "QMultiple.frx":81EFD
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Exporta a Xls"
      Top             =   4920
      Width           =   735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   12
      Top             =   4680
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
      Caption         =   "Registros Encontrados"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "QMultiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim rt As New ADODB.Recordset
Dim Criterio, f_iltro, ccamp1, ccamp2, cDscCat As String
Dim nposicion, wdato As Integer
Dim sql2 As String

Private Sub Form_Load()
   Lb1.Clear
   Lb1.AddItem "Estado        :ESTADO"
   Lb1.AddItem "Fec. Nac.     :FECNAC"
   Lb1.AddItem "Est.Civil     :ESTCIVIL"
   Lb1.AddItem "Grpo Horario  :GRPGEN"
   Lb1.AddItem "Tarif.Guardia :NIVGUARDIA"
   Lb1.AddItem "Cod. Nivel    :CODNIV"
   Lb1.AddItem "Rem. Total    :REMTOT"
  Lb2.Clear
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpTablasAux WHERE grupo='REG' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To CboRegPlh.ListCount - 1
           If CboRegPlh.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          CboRegPlh.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpTablasAux WHERE grupo='CAT' ORDER BY descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
       ExisteYa = False
       For i = 0 To Cbocategoria.ListCount - 1
           If Cbocategoria.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("codigo") Then
              ExisteYa = True
              Exit For
           End If
       Next i
       If Not ExisteYa Then
          Cbocategoria.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("codigo")
       End If
       rs.MoveNext
  Wend
  rs.Close
  
  ChkXLS.Value = 1
  'LLena el Combo TipoPLA
    CboTipoPla.Clear
    CboTipoPla.AddItem "ASISTENCIAL"
    CboTipoPla.AddItem "ADMINISTRATIVO"
    CboTipoPla.ListIndex = 0
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpCcosto ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To CboCCosto.ListCount - 1
         If CboCCosto.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("Codigo") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        CboCCosto.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("Codigo")
     End If
     rs.MoveNext
  Wend
  rs.Close
 
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpUndPrgRol ORDER BY Descripcion", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
     ExisteYa = False
     For i = 0 To CboUpr.ListCount - 1
         If CboUpr.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR") Then
            ExisteYa = True
            Exit For
         End If
     Next i
     If Not ExisteYa Then
        CboUpr.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR")
     End If
     rs.MoveNext
  Wend
  rs.Close
    
  Txcategoria = "": Cbocategoria.Text = ""
  TxPlH = "":  CboRegPlh.Text = ""
  TxTipoPla = "":  CboTipoPla.Text = ""
  Txccosto = "":  CboCCosto.Text = ""
  TxcodUPR = "": CboUpr.Text = ""
  TxCargo = "": TxDscCargo = ""
  Txfiltro = ""
  Set DD = New Recordset
  With DD
     .Fields.Append "tarjeta", adChar, 5
     .Fields.Append "Documento", adChar, 10
     .Fields.Append "estado", adChar, 1
     .Fields.Append "apelpat", adChar, 20
     .Fields.Append "apelmat", adChar, 20
     .Fields.Append "nombres", adChar, 20
     .Fields.Append "campo1", adChar, 15
     .Fields.Append "campo2", adChar, 15
     .Open
  End With
  Set DataGrid1.DataSource = DD
  For i = 0 To DataGrid1.Columns.Count - 1
     With DataGrid1.Columns(i)
         Select Case i
          Case 0
          '.Caption = "tarjeta"
          .Width = Me.TextWidth(String(6, "S"))
          Case 1
          '.Caption = "D.N.I."
          .Width = Me.TextWidth(String(10, "S"))
          Case 2
          .Width = Me.TextWidth(String(6, "S"))
          Case 3
          .Caption = "Paterno"
          .Width = Me.TextWidth(String(15, "S"))
           Case 4
          .Caption = "Materno"
          .Width = Me.TextWidth(String(15, "S"))
          Case 5
          .Width = Me.TextWidth(String(20, "S"))
           Case 6
          .Width = Me.TextWidth(String(15, "S"))
          Case 7
          .Width = Me.TextWidth(String(15, "S"))
          
          End Select
         '.NumberFormat = "##,##0.0"
         .Locked = True
     End With
  Next
  'Call CreaTabla
  Me.Refresh
End Sub
Private Sub CboRegPlh_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboRegPlh.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub CboRegPlh_LostFocus()
  If IsEmpty(CboRegPlh.Text) Then
  Else
     j = CboRegPlh.ListIndex
     nposicion = InStr(CboRegPlh.List(j), ":")
     nlen = Len(CboRegPlh.List(j))
     TxPlH.Text = Right(CboRegPlh.List(j), nlen - nposicion)
  End If
End Sub
Private Sub Cbotipopla_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(CboTipoPla.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbotipopla_LostFocus()
  If IsEmpty(CboTipoPla.Text) Then
  Else
     j = CboTipoPla.ListIndex
     TxTipoPla = Left(CboTipoPla.List(j), 3)
  End If
End Sub
Private Sub Cbocategoria_GotFocus()
  Const CB_SHOWDROPDOWN = &H14F
  Dim Tmp
  Tmp = SendMessage(Cbocategoria.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cbocategoria_LostFocus()
  If IsEmpty(Cbocategoria.Text) Then
  Else
     j = Cbocategoria.ListIndex
     If j <> -1 Then
        nposicion = InStr(Cbocategoria.List(j), ":")
        nlen = Len(Cbocategoria.List(j))
        Txcategoria = Trim(Right(Cbocategoria.List(j), nlen - nposicion))
     End If
     Exit Sub
  End If
End Sub

Private Sub Cboccosto_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(CboCCosto.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub
Private Sub Cboccosto_LostFocus()
  If IsEmpty(CboCCosto.Text) Then
  Else
     j = CboCCosto.ListIndex
     nposicion = InStr(CboCCosto.List(j), ":")
     nlen = Len(CboCCosto.List(j))
     Txccosto = Right(CboCCosto.List(j), nlen - nposicion)
     
         sql = "SELECT     tab1.CCOSTO, tab1.CODUPR, tab2.Descripcion " & _
         "FROM         CPCCostoUPR AS tab1 INNER JOIN " & _
         "                 CpUndPrgRol AS tab2 ON tab1.CODUPR = tab2.CodUPR " & _
         "WHERE CCOSTO='" & Txccosto + "'  ORDER BY tab2.descripcion "
     
         CboUpr.Clear
         Set rs = New Recordset
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly
         While Not rs.EOF
            ExisteYa = False
            For i = 0 To CboUpr.ListCount - 1
                If CboUpr.List(i) = rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR") Then
                    ExisteYa = True
                    Exit For
                End If
            Next i
            If Not ExisteYa Then
                CboUpr.AddItem rs.Fields("Descripcion") & "  :" & rs.Fields("CodUPR")
            End If
            rs.MoveNext
         Wend
         rs.Close
         
  End If
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
     nposicion = InStr(CboUpr.List(j), ":")
     nlen = Len(CboUpr.List(j))
     TxcodUPR = Right(CboUpr.List(j), nlen - nposicion)
  End If
End Sub

Private Sub TxCargo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(TxCargo) = "" Then
         Load Ayuda3
         Ayuda3.Tag = 3
         Ayuda3.Show 1
      End If
      Set rs = New Recordset
      sql = "SELECT * FROM CpCargos WHERE codigo='" & Trim(TxCargo) & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      TxDscCargo = ""
      If Not rs.EOF Then
         TxDscCargo = Trim(rs!descripcion)
      Else
         MsgBox "Cargo no existe...", vbExclamation, "[ ERROR ]"
         TxCargo = ""
      End If
      rs.Close
      Exit Sub
   End If
   If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
      KeyAscii = 0
   End If
End Sub
Private Sub Command1_Click()
   If Lb2.ListCount = 2 Then Exit Sub     'maximo permitido
   j = Lb1.ListIndex
   'check NO duplicados
   nposicion = InStr(Lb1.List(j), ":")
   nlen = Len(Lb1.List(j))
   ccadena = Right(Lb1.List(j), nlen - nposicion)
   For X = 0 To Lb2.ListCount - 1
      If Lb2.List(X) = ccadena Then Exit Sub
   Next X
   Lb2.AddItem ccadena
   If Txcampo1 = "" Then
      Txcampo1 = ccadena
   ElseIf Txcampo2 = "" Then
          Txcampo2 = ccadena
   End If
   Lblcampos.Caption = "Hay " & Lb2.ListCount & " Item(s)"
   Lb2.ListIndex = Lb2.ListCount - 1
   Lb1.SetFocus
End Sub
Private Sub Command2_Click()
   j = Lb2.ListIndex
   If Lb2.List(j) = Txcampo1 Then
         Txcampo1 = ""
   ElseIf Lb2.List(j) = Txcampo2 Then
          Txcampo2 = ""
   End If
   Lb2.RemoveItem Lb2.ListIndex
   Lblcampos.Caption = "Hay " & Lb2.ListCount & " Item(s)"
End Sub

Private Sub cmdcarga_Click()
   Criterio = " (cpgeneral.baja is NULL) "
   f_iltro = ""
   If Trim(TxPlH) <> "" Then
      Criterio = Criterio + " and  ( cpgeneral.plh='" & TxPlH & "')"
      f_iltro = TxPlH
   End If
   If TxTipoPla <> "" Then
      If Len(Criterio) > 0 Then
        Criterio = Criterio + " and  ( cpgeneral.Tipopla='" & TxTipoPla & "')"
        f_iltro = f_iltro & ", " + TxTipoPla
      Else
        Criterio = "(cpgeneral.Tipopla='" & TxTipoPla & "')"
        f_iltro = TxTipoPla
      End If
   End If
   If Txcategoria <> "" Then
      If Len(Criterio) > 0 Then
         Criterio = Criterio + " and ( categoria='" & Trim(Txcategoria) & "')"
         f_iltro = f_iltro + " Cat: " & Trim(Txcategoria)
      Else
         Criterio = "( categoria='" & Trim(Txcategoria) & "')"
         f_iltro = "Cat: " & Trim(Txcategoria)
      End If
   End If
   If Txccosto <> "" Then
      If Len(Criterio) > 0 Then
        Criterio = Criterio + " and  (cpgeneral.ccosto='" & Txccosto & "')"
        f_iltro = f_iltro & ", C.Costo "
      Else
        Criterio = " (cpgeneral.ccosto='" & Txccosto & "')"
        f_iltro = "C.Costo "
      End If
   End If
   If TxcodUPR <> "" Then
      If Len(Criterio) > 0 Then
         Criterio = Criterio + " and  (cpgeneral.codupr='" & TxcodUPR & "')"
         f_iltro = f_iltro & ", UPR "
      Else
         Criterio = " (cpgeneral.codupr='" & TxcodUPR & "')"
         f_iltro = "UPR "
      End If
   End If
   If Check1.Value = 1 Then        'activos
      If Len(Criterio) > 0 Then
          Criterio = Criterio + " and ( cpgeneral.estado='A')"
      Else
          Criterio = " ( cpgeneral.estado='A')"
      End If
   End If
   
   If Criterio = "" Then
      MsgBox "Seleccione Criterio ", vbExclamation
      Exit Sub
   End If
   Txfiltro = f_iltro
   DataGrid1.Visible = False
   Do While Not DD.EOF
      DD.Delete
      DD.MoveNext
   Loop
   Cnx.BeginTrans
   Set rs = New Recordset
   Sq0 = "SELECT  * FROM  cpgeneral WHERE " & Criterio & " ORDER BY ccosto,codUpr,apelpat,apelmat"
   rs.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   wdato = 0
   Do While Not rs.EOF
      DD.AddNew
      DD!tarjeta = rs!tarjeta
      'If IsDate(rs!fecrec) Then DD!fecrec = rs!fecrec
      DD!documento = IIf(IsNull(rs!documento), "", rs!documento)
      DD!estado = IIf(IsNull(rs!estado), "", rs!estado)
      DD!apelpat = IIf(IsNull(rs!apelpat), "", rs!apelpat)
      DD!apelmat = IIf(IsNull(rs!apelmat), "", rs!apelmat)
      DD!nombres = IIf(IsNull(rs!nombres), "", rs!nombres)
      DD.Update
      wdato = wdato + 1
      rs.MoveNext
   Loop
   Cnx.CommitTrans
   rs.Close
   If wdato > 1 Then
      CargaData (0)
   End If
   DataGrid1.Refresh
   DataGrid1.Caption = "REGISTROS SELECCIONADOS " & wdato
   DataGrid1.Visible = True
   DataGrid1.SetFocus
End Sub

Private Sub CmdClear_Click()
   DataGrid1.Visible = False
   Do While Not DD.EOF
      DD.Delete
      DD.MoveFirst
   Loop
   
   TxPlH = "":  CboRegPlh.Text = ""
   TxTipoPla = "":  CboTipoPla.Text = ""
   Txcategoria = "":  Cbocategoria.Text = ""
   Txccosto = "":  CboCCosto.Text = ""
   TxcodUPR = "": CboUpr.Clear
   TxCargo = "": TxDscCargo = ""
   Lb2.Clear
   Txcampo1 = "": Txcampo2 = ""
   Lblcampos.Caption = ""
   Txfiltro = "": ChkXLS.Value = 1
   wdato = 0
   DataGrid1.Caption = "REGISTROS SELECCIONADOS "
   DataGrid1.Visible = True
End Sub
Private Sub cmdsalir_Click()
   wTpoQuery = ""
   Set DD = Nothing
   For cont = Forms.Count - 1 To 0 Step -1
      If Forms(cont).Visible = False Then
      Else
        ' Form Abierto
        If Forms(cont).Name <> "CpMenu" And Forms(cont).Name <> "QMultiple" Then
           Unload Forms(cont)
        End If
      End If
   Next
   
   Unload Me
End Sub

Private Sub DataGrid1_Click()
If DD.RecordCount = 0 Then Exit Sub      'bloquea seleccion en dbgrid vacio
On Error GoTo Errfinal
Errfinal:
   MsgBox Err.Description & " ErRor: " & Err.Number & vbCr
End Sub
Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
  With DD
    If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
       .Sort = .Fields(ColIndex).[Name] & " Desc"
    Else
       .Sort = .Fields(ColIndex).[Name] & " Asc"
    End If
  End With
  
End Sub
Private Sub CmdPrint_Click()
   If wdato > 0 Then
     'CreaTabla
      nTipRpt = 1.1
      wTituRpt = f_iltro
      If Check1.Value = 1 Then
         wTituRpt = wTituRpt + " [ACTIVOS]"
      End If
      InfoCpSQL.Left = 0: InfoCpSQL.Top = 0
      InfoCpSQL.Refresh
      InfoCpSQL.Show
   Else
      MsgBox "No Hay datos Validos..", vbExclamation
   End If
   
End Sub

Private Sub mnuFicImpSelec_Click()
    ' Seleccionar la impresora a usar                               (23/Ene/00)
    ' La detección de errores es por si no hay impresora instalada
    On Error Resume Next
    With CommonDialog1
        .DialogTitle = "Seleccionar impresora"
        .Flags = cdlPDPrintSetup
        .ShowPrinter
    End With
    Err = 0
 End Sub
Private Sub CmdExportaXLS_Click()
    wTpoQuery = "1.1"
    If ChkXLS.Value = 0 Then
       Call ExportTableToXLS_click
    Else
       wNameXLS = "C:\RPT\QMultiple.xls"
       Call Exportar_ADO_Excel
    End If
       
    Beep 500, 350
    MsgBox "Genera C:\RPT\QMultiple.xls", vbInformation
    Exit Sub
End Sub
Private Sub CargaData(nopc)
'***  CARGA  registros  ***   0:Todos   1:Orden
    Cnx.BeginTrans
    Cnx.Execute "DELETE  FROM CpTmGeneral"
    Set r_d = New ADODB.Recordset
    sql = "SELECT * FROM CpGeneral WHERE " & Criterio
    r_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    Set rt = New Recordset
    sq_l = "SELECT * FROM CpTmGeneral"
    rt.Open sq_l, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    Do While Not r_d.EOF
       ccamp1 = "": ccamp2 = ""
       rt.AddNew
       rt!tarjeta = r_d!tarjeta
       rt!documento = IIf(IsEmpty(r_d!documento), "", r_d!documento)
       rt!estado = IIf(IsEmpty(r_d!estado), "", r_d!estado)
       rt!apelpat = IIf(IsEmpty(r_d!apelpat), "", r_d!apelpat)
       rt!apelmat = IIf(IsEmpty(r_d!apelmat), "", r_d!apelmat)
       rt!nombres = IIf(IsEmpty(r_d!nombres), "", r_d!nombres)
       rt!PLH = IIf(IsEmpty(r_d!PLH), "", r_d!PLH)
       rt!tipoPla = IIf(IsEmpty(r_d!tipoPla), "", r_d!tipoPla)
       rt!Categoria = IIf(IsEmpty(r_d!Categoria), "", r_d!Categoria)
       rt!CCosto = IIf(IsEmpty(r_d!CCosto), "", r_d!CCosto)
       rt!CodUpr = IIf(IsEmpty(r_d!CodUpr), "", r_d!CodUpr)
       rt!Cargo = IIf(IsEmpty(r_d!Cargo), "", r_d!Cargo)
       If IsDate(r_d!fecnac) Then rt!fecnac = r_d!fecnac
       rt!grpgen = IIf(IsEmpty(r_d!grpgen), "", r_d!grpgen)
       rt!CodNiv = IIf(IsEmpty(r_d!CodNiv), "", r_d!CodNiv)
       rt!NivGuardia = IIf(IsEmpty(r_d!NivGuardia), "", r_d!NivGuardia)
       rt!RemTot = IIf(IsEmpty(r_d!RemTot), "", r_d!RemTot)
       cCodigo = r_d!Categoria
       If Trim(cCodigo) <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM CpTablasAux WHERE grupo='CAT' and codigo='" + cCodigo + "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dsccategoria = "" + Trim(rs!descripcion)
          End If
          rs.Close
       End If
       cCodigo = r_d!CCosto
       If Trim(cCodigo) <> "" Then
          Set rs = New Recordset
          rs.Open "SELECT * FROM Cpccosto WHERE ( codigo='" & Trim(cCodigo) & "')", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!dscccosto = "" + rs!descripcion
          End If
          rs.Close
       End If
       cCodigo = r_d!Cargo
       If IsEmpty(cCodigo) Then
       Else
          Set rs = New Recordset
          rs.Open "SELECT * FROM CpCargos WHERE ( codigo='" & cCodigo & "')", Cnx, adOpenStatic, adLockReadOnly
          If Not rs.EOF Then
             rt!DscCargo = "" + Left(rs!descripcion, 35)
          End If
          rs.Close
       End If
       cCodigo = r_d!CodUpr
       If IsEmpty(cCodigo) Then
       Else
          Set rs = New Recordset
          rs.Open "SELECT * FROM Cpundprgrol WHERE ( codupr='" & cCodigo & "')", Cnx, adOpenStatic, adLockOptimistic
          If Not rs.EOF Then
             rt!dscUPR = "" + rs!descripcion
          End If
          rs.Close
       End If
       
       If Trim(Txcampo1) = "FECNAC" Or Trim(Txcampo2) = "FECNAC" Then
          If ccamp1 = "" And Txcampo1 = "FECNAC" Then
             ccamp1 = IIf(IsDate(r_d!fecnac), Format(r_d!fecnac, "dd/mm/yyyy"), "")
          ElseIf ccamp2 = "" And Txcampo2 = "FECNAC" Then
                 ccamp2 = IIf(IsDate(r_d!fecnac), Format(r_d!fecnac, "dd/mm/yyyy"), "")
          End If
       End If
       If Trim(Txcampo1) = "ESTADO" Or Trim(Txcampo2) = "ESTADO" Then
          cdsc = ""
          cCodigo = IIf(IsNull(r_d!estado), "", r_d!estado)
          If IsEmpty(cCodigo) Then
          Else
            Set rs = New Recordset
            rs.Open "SELECT * FROM Cptablas WHERE tabla='11' and codigo='" + cCodigo + "'", Cnx, adOpenStatic, adLockReadOnly
            If Not rs.EOF Then
               cdsc = "" + Left(rs!descripcion, 15)
            End If
            rs.Close
            If ccamp1 = "" And Txcampo1 = "ESTADO" Then
               ccamp1 = cdsc
            ElseIf ccamp2 = "" And Txcampo2 = "ESTADO" Then
                   ccamp2 = cdsc
            End If
          End If
       End If
       If Txcampo1 = "GRPGEN" Or Txcampo2 = "GRPGEN" Then
          If ccamp1 = "" And Txcampo1 = "GRPGEN" Then
             ccamp1 = IIf(IsNull(r_d!grpgen), "", r_d!grpgen)
          ElseIf ccamp2 = "" And Txcampo2 = "GRPGEN" Then
                 ccamp2 = IIf(IsNull(r_d!grpgen), "", r_d!grpgen)
          End If
       End If
       If Txcampo1 = "ESTCIVIL" Or Txcampo2 = "ESTCIVIL" Then
          cdsc = ""
          cCodigo = IIf(IsNull(r_d!EstCivil), "", r_d!EstCivil)
          If IsEmpty(cCodigo) Then
          Else
             Set rs = New Recordset    '-MBD Legajos
             rs.Open "SELECT * FROM TablasAux WHERE grupo='SCI' AND Codigo='" + cCodigo + "'", Cnx1, adOpenStatic, adLockOptimistic
             If Not rs.EOF Then
                cdsc = "" + Left(rs!descripcion, 15)
             End If
             rs.Close
             If ccamp1 = "" And Txcampo1 = "ESTCIVIL" Then
                ccamp1 = cdsc
             ElseIf ccamp2 = "" And Txcampo2 = "ESTCIVIL" Then
                    ccamp2 = cdsc
             End If
          End If
       End If
        If Txcampo1 = "NIVGUARDIA" Or Txcampo2 = "NIVGUARDIA" Then
          If ccamp1 = "" And Txcampo1 = "NIVGUARDIA" Then
             ccamp1 = IIf(IsNull(r_d!NivGuardia), "", r_d!NivGuardia)
          ElseIf ccamp2 = "" And Txcampo2 = "NIVGUARDIA" Then
                 ccamp2 = IIf(IsNull(r_d!NivGuardia), "", r_d!NivGuardia)
          End If
       End If
      
       If Txcampo1 = "CODNIV" Or Txcampo2 = "CODNIV" Then
          If ccamp1 = "" And Txcampo1 = "CODNIV" Then
             ccamp1 = IIf(IsNull(r_d!CodNiv), "", r_d!CodNiv)
          ElseIf ccamp2 = "" And Txcampo2 = "CODNIV" Then
                 ccamp2 = IIf(IsNull(r_d!CodNiv), "", r_d!CodNiv)
          End If
       End If
       If Txcampo1 = "REMTOT" Or Txcampo2 = "REMTOT" Then
          If ccamp1 = "" And Txcampo1 = "REMTOT" Then
             ccamp1 = IIf(IsNull(r_d!RemTot), "", Format(r_d!RemTot, "######0.00"))
          ElseIf ccamp2 = "" And Txcampo2 = "REMTOT" Then
                 ccamp2 = IIf(IsNull(r_d!RemTot), "", Format(r_d!RemTot, "######0.00"))
          End If
       End If
       rt!Campo1 = Left(ccamp1, 15)
       rt!campo2 = Left(ccamp2, 15)
       rt.Update
       r_d.MoveNext
    Loop
    Cnx.CommitTrans
    rt.Close
    r_d.Close
End Sub

