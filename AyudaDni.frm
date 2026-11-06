VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form AyudaDni 
   Caption         =   "Maestra de Servidores"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11310
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "AyudaDni.frx":0000
   ScaleHeight     =   7275
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   10215
      TabIndex        =   8
      Top             =   1680
      Width           =   10275
      Begin MSDataGridLib.DataGrid Grd 
         Bindings        =   "AyudaDni.frx":80266
         Height          =   4695
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   9435
         _ExtentX        =   16642
         _ExtentY        =   8281
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
         Caption         =   "Catalogo de Pacientes"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "tarjeta"
            Caption         =   "Tarjeta"
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
            DataField       =   "apelpat"
            Caption         =   "A.PATERNO"
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
         BeginProperty Column02 
            DataField       =   "apelmat"
            Caption         =   "A.MATERNO"
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
         BeginProperty Column03 
            DataField       =   "Nombres"
            Caption         =   "Nombres"
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
         BeginProperty Column04 
            DataField       =   "Documento"
            Caption         =   "DNI"
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
         BeginProperty Column05 
            DataField       =   "estado"
            Caption         =   "Estado"
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
         BeginProperty Column06 
            DataField       =   "fecnac"
            Caption         =   "FEC.NAC."
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
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ubicar"
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   9615
      Begin VB.TextBox TxUbiAP 
         Height          =   315
         Left            =   240
         MaxLength       =   15
         TabIndex        =   0
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox TxUbiAM 
         Height          =   315
         Left            =   2760
         MaxLength       =   15
         TabIndex        =   4
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox TxUbiNm 
         Height          =   315
         Left            =   5880
         MaxLength       =   20
         TabIndex        =   3
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Apel.  Paterno"
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Apel. Materno"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Nombres"
         Height          =   255
         Left            =   6960
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "AyudaDni.frx":80278
      Height          =   615
      Left            =   10320
      Picture         =   "AyudaDni.frx":806BA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Ado 
      Height          =   330
      Left            =   2205
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
Attribute VB_Name = "AyudaDni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sq2 As String
Private Sub TxubiAP_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(TxUbiAP) = "" Then
        Exit Sub
     Else
        Call Command1_Click
     End If
  End If
End Sub
Private Sub TxubiAM_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(TxUbiAP) = "" Or Trim(TxUbiAM) = "" Then
        Exit Sub
     Else
        Call Command1_Click
     End If
  End If
End Sub
Private Sub Command1_Click()
      Dim rt As Recordset
      Set rt = New Recordset
      If Me.Tag = 1 Or Me.Tag = 2 Or Me.Tag = 3 Or Me.Tag = 4 Then
         If TxUbiAP <> "" And TxUbiAM = "" And TxUbiNm = "" Then
            sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and ApelPat LIKE '" & Trim(TxUbiAP) & "%'  ORDER BY Apelpat , apelmat , nombres"
         ElseIf TxUbiAP <> "" And TxUbiAM <> "" And TxUbiNm = "" Then
             sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and ApelPat='" & Trim(TxUbiAP) & "' and ApelMat='" & Trim(TxUbiAM) & "'  ORDER BY Apelpat , apelmat , nombres"
         ElseIf TxUbiAP <> "" And TxUbiAM <> "" And TxUbiNm <> "" Then
             sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and ApelPat='" & Trim(TxUbiAP) & "' and ApelMat='" & Trim(TxUbiAM) & "' and  Nombres='" & Trim(TxUbiNm) & "' ORDER BY Apelpat, Apelmat , Nombres"
         End If
      Else
        If TxUbiAP.Text <> "" And TxUbiAM = "" And TxUbiNm.Text = "" Then
            sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and ApelPat='" & Trim(TxUbiAP.Text) & "' ORDER BY Apelpat , apelmat , nombres"
         ElseIf TxUbiAP.Text <> "" And TxUbiAM <> "" And TxUbiNm = "" Then
             sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and ApelPat='" & Trim(TxUbiAP) & "' and ApelMat='" & Trim(TxUbiAM) & "'  ORDER BY Apelpat , apelmat , nombres"
         ElseIf TxUbiAP <> "" And TxUbiAM <> "" And TxUbiNm <> "" Then
             sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and ApelPat='" & Trim(TxUbiAP) & "' and ApelMat='" & Trim(TxUbiAM) & "' and  Nombres='" & Trim(TxUbiNm) & "'  ORDER BY Apelpat, Apelmat , Nombres"
         End If
      End If
      rt.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
      If rt.EOF Then
         MsgBox "No Ubicado", vbExclamation
         rt.Close
         TxUbiAP = "": TxUbiAM = "": TxUbiNm = ""
         TxUbiAP.SetFocus
         Exit Sub
      Else
        Ado.ConnectionString = CString
        Ado.RecordSource = sq2
        Ado.Refresh
        Grd.SetFocus
        Grd.Row = 0
        Grd.Col = 0
      End If
      rt.Close
End Sub
Private Sub Grd_DblClick()
    Call Grd_KeyDown(13, 0)
End Sub
Private Sub Grd_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo final
    If KeyCode = 13 Then
       If Me.Tag = 0 Then
       ElseIf Me.Tag = 1 Then
             cCostoUpr.Txdni1 = Trim(Grd.Columns(4).Text)    'documento->Dni
             cCostoUpr.Lblusu1.Caption = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text) + " ", 35)
             cCostoUpr.TxEst1 = Trim(Grd.Columns(5).Text)
             cCostoUpr.TxFn1 = Trim(Grd.Columns(6).Text)
       ElseIf Me.Tag = 2 Then
              cCostoUpr.Txdni2 = Trim(Grd.Columns(4).Text)   'documento->Dni
              cCostoUpr.Lblusu2.Caption = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text) + " ", 35)
              cCostoUpr.TxEst2 = Trim(Grd.Columns(5).Text)
              cCostoUpr.Txfn2 = Trim(Grd.Columns(6).Text)
       ElseIf Me.Tag = 3 Then
              cCostoUpr.Txdni3 = Trim(Grd.Columns(4).Text)    'documento->Dni
              cCostoUpr.Lblusu3.Caption = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text) + " ", 35)
              cCostoUpr.TxEst3 = Trim(Grd.Columns(5).Text)
              cCostoUpr.Txfn3 = Trim(Grd.Columns(6).Text)
       ElseIf Me.Tag = 4 Then
              CargaMarcas.Tx_dni = Trim(Grd.Columns(4).Text)    'documento->Dni
              CargaMarcas.Txdscdni = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text) + " ", 35)
       End If
       Unload Me
    End If
    Exit Sub
final:
End Sub


