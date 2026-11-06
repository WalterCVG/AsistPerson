VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Ayuda2H 
   Caption         =   "Ayuda"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "Ayuda2H.frx":0000
   ScaleHeight     =   4935
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   6855
      Begin VB.TextBox TxBuscar 
         Height          =   315
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   3885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR DENOMINACION"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   285
         Width           =   1965
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   8535
      Begin MSDataGridLib.DataGrid Grd 
         Bindings        =   "Ayuda2H.frx":80266
         Height          =   2970
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   5239
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         BorderStyle     =   0
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
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
         Caption         =   "registro seleccionados"
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
            MarqueeStyle    =   1
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado 
      Height          =   330
      Left            =   120
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Ado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Ayuda2H"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Grd_DblClick()
    Call Grd_KeyDown(13, 0)
End Sub
Private Sub Grd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo final
    If KeyCode = 13 Then
       If Me.Tag = 1 Then
       ElseIf Me.Tag = 2 Then
              cCostoHor.TxHorario.Text = Grd.Columns(0).Text          'codigo
              cCostoHor.Txdescripcion.Text = Grd.Columns(1).Text
       ElseIf Me.Tag = 3 Then
       End If
       Unload Me
    End If
    Exit Sub
final:
End Sub
Private Sub TxBUscar_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
       Ado.ConnectionString = CString
       If Me.Tag = 1 Then
       ElseIf Me.Tag = 2 Then
              Ado.RecordSource = "SELECT * FROM CpHorarioNew WHERE asistencial='S' AND descripcion like '%" & Trim(Txbuscar.Text) & "%' ORDER BY descripcion "
       End If
       Ado.Refresh
       Grd.SetFocus
       Grd.Row = 0
       Grd.Col = 0
    End If
End Sub


