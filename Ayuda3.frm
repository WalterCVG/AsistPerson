VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Ayuda3 
   Caption         =   " Ubicar CARGO"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Ayuda3.frx":0000
   ScaleHeight     =   5685
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   3975
      Left            =   240
      ScaleHeight     =   3915
      ScaleWidth      =   8775
      TabIndex        =   3
      Top             =   240
      Width           =   8835
      Begin MSDataGridLib.DataGrid Grd 
         Height          =   3450
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   6085
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         BorderStyle     =   0
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Registros ubicados"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "codigo"
            Caption         =   "Código"
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
            DataField       =   "descripcion"
            Caption         =   "Descripción"
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
            MarqueeStyle    =   1
            BeginProperty Column00 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   6270.236
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox SSPanel2 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   1080
      ScaleHeight     =   795
      ScaleWidth      =   6960
      TabIndex        =   0
      Top             =   4320
      Width           =   7020
      Begin VB.TextBox Txbuscar 
         Height          =   315
         Left            =   240
         MaxLength       =   50
         TabIndex        =   1
         Top             =   360
         Width           =   6435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción a ubicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Top             =   75
         Width           =   2145
      End
   End
End
Attribute VB_Name = "Ayuda3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dt As New ADODB.Recordset
Dim cpuntero As Integer
Private Sub Form_Load()
  Set Dt = New Recordset
  With Dt
     .Fields.Append "codigo", adChar, 4
     .Fields.Append "descripcion", adChar, 50
     .Open
  End With
  Set Grd.DataSource = Dt
  
  Set rs = New Recordset
  rs.Open "SELECT * FROM cpcargos ORDER BY Descripcion ", Cnx, adOpenStatic, adLockReadOnly
  rs.MoveFirst
  While Not rs.EOF
     Dt.AddNew
     Dt!codigo = rs!codigo
     Dt!descripcion = rs!descripcion
     rs.MoveNext
  Wend
  rs.Close
  Set Grd.DataSource = Dt
End Sub
Private Sub Grd_DblClick()
    Call Grd_KeyDown(13, 0)
End Sub
Private Sub Grd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo final
    If KeyCode = 13 Then
       If Me.Tag = 1 Then
          General.Txcargo.Text = Grd.Columns(0).Text        'codigo
          General.Txdsccargo.Text = Grd.Columns(1).Text   'descripción
       ElseIf Me.Tag = 2 Then
          TardaInasis.Txcargo.Text = Grd.Columns(0).Text        'codigo
          TardaInasis.Txdsccargo.Text = Grd.Columns(1).Text   'descripción
       ElseIf Me.Tag = 3 Then
          QMultiple.Txcargo.Text = Grd.Columns(0).Text        'codigo
          QMultiple.Txdsccargo.Text = Grd.Columns(1).Text   'descripción
       End If
       Unload Me
    End If
    Exit Sub
final:
End Sub
Private Sub Grd_HeadClick(ByVal ColIndex As Integer)
  With Dt
    If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
       .Sort = .Fields(ColIndex).[Name] & " Desc"
    Else
       .Sort = .Fields(ColIndex).[Name] & " Asc"
    End If
  End With
End Sub
Private Sub TxBuscar_Change()
    Dim cBuscar As String
    ''cpuntero = Grd.Bookmark
    If Txbuscar = "" Then Exit Sub
    cBuscar = "descripcion like '" & Txbuscar & "*'"
    Dt.MoveFirst
    Dt.Find (cBuscar)
    If Dt.EOF Then
       Txbuscar.BackColor = &HFF&
    Else
       Txbuscar.BackColor = &H80000005
       Grd.Refresh
    End If
End Sub
