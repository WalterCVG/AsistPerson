VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Ayuda1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maestra de Servidores"
   ClientHeight    =   7200
   ClientLeft      =   840
   ClientTop       =   1425
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Ayuda1.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      DownPicture     =   "Ayuda1.frx":80266
      Height          =   615
      Left            =   9600
      Picture         =   "Ayuda1.frx":806A8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ubicar"
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   9135
      Begin VB.TextBox TxUbiNm 
         Height          =   315
         Left            =   5400
         MaxLength       =   20
         TabIndex        =   3
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox TxUbiAM 
         Height          =   315
         Left            =   2760
         MaxLength       =   15
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox TxUbiAP 
         Height          =   315
         Left            =   240
         MaxLength       =   15
         TabIndex        =   0
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label5 
         Caption         =   "Nombres"
         Height          =   255
         Left            =   6960
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Apel. Materno"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Apel.  Paterno"
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   5295
      Left            =   240
      ScaleHeight     =   5235
      ScaleWidth      =   10095
      TabIndex        =   4
      Top             =   1680
      Width           =   10155
      Begin MSDataGridLib.DataGrid Grd 
         Bindings        =   "Ayuda1.frx":80883
         Height          =   4695
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   9555
         _ExtentX        =   16854
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
         Caption         =   "Catalogo Maestro"
         ColumnCount     =   15
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
            DataField       =   "codupr"
            Caption         =   "Codupr"
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
            DataField       =   "grpgen"
            Caption         =   "grpgen"
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
         BeginProperty Column06 
            DataField       =   "actividad"
            Caption         =   "actividad"
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
         BeginProperty Column07 
            DataField       =   "ccosto"
            Caption         =   "ccosto"
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
         BeginProperty Column08 
            DataField       =   "carest"
            Caption         =   "carest"
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
         BeginProperty Column09 
            DataField       =   "Plh"
            Caption         =   "Plh"
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
         BeginProperty Column10 
            DataField       =   "cargo"
            Caption         =   "cargo"
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
         BeginProperty Column11 
            DataField       =   "FecNac"
            Caption         =   "FecNac"
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
         BeginProperty Column12 
            DataField       =   "tipoPla"
            Caption         =   "TipoPla"
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
         BeginProperty Column13 
            DataField       =   "Documento"
            Caption         =   "Dni"
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
         BeginProperty Column14 
            DataField       =   "CadHep"
            Caption         =   "Cadena"
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
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
            EndProperty
            BeginProperty Column14 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado 
      Height          =   330
      Left            =   2565
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione haciendo Doble Click "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1440
      Width           =   2895
   End
End
Attribute VB_Name = "Ayuda1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sq2 As String

Private Sub TxubiAP_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(TxUbiAP.Text) = "" Then
        Exit Sub
     Else
        Call Command1_Click
     End If
  End If
End Sub
Private Sub TxubiAM_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Trim(TxUbiAP.Text) = "" Or Trim(TxUbiAM.Text) = "" Then
        Exit Sub
     Else
        Call Command1_Click
     End If
  End If
End Sub
Private Sub Command1_Click()
      Dim rt As Recordset
      Set rt = New Recordset
      If Me.Tag = 6.1 Then
         If TxUbiAP.Text <> "" And TxUbiAM.Text = "" And TxUbiNm.Text = "" Then
            sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and ApelPat LIKE '" & Trim(TxUbiAP.Text) & "%' And ccosto='" & ProgAdministrativo.Tx_ccosto.Text & "' ORDER BY Apelpat , apelmat , nombres"
         ElseIf TxUbiAP.Text <> "" And TxUbiAM.Text <> "" And TxUbiNm = "" Then
             sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and ApelPat='" & Trim(TxUbiAP.Text) & "' and ApelMat='" & Trim(TxUbiAM.Text) & "' And ccosto='" & ProgAdministrativo.Tx_ccosto.Text & "'  ORDER BY Apelpat , apelmat , nombres"
         ElseIf TxUbiAP <> "" And TxUbiAM <> "" And TxUbiNm <> "" Then
             sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and ApelPat='" & Trim(TxUbiAP.Text) & "' and ApelMat='" & Trim(TxUbiAM.Text) & "' and  Nombres='" & Trim(TxUbiNm.Text) & "'  And ccosto='" & ProgAdministrativo.Tx_ccosto.Text & "' ORDER BY Apelpat, Apelmat , Nombres"
         End If
      ElseIf Me.Tag = 6.3 Then
         If TxUbiAP.Text <> "" And TxUbiAM.Text = "" And TxUbiNm.Text = "" Then
            sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and ApelPat LIKE '" & Trim(TxUbiAP.Text) & "%' And ccosto='" & ProgEspecial.Tx_ccosto.Text & "' ORDER BY Apelpat , apelmat , nombres"
         ElseIf TxUbiAP.Text <> "" And TxUbiAM.Text <> "" And TxUbiNm = "" Then
             sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and ApelPat='" & Trim(TxUbiAP.Text) & "' and ApelMat='" & Trim(TxUbiAM.Text) & "' And ccosto='" & ProgEspecial.Tx_ccosto.Text & "'  ORDER BY Apelpat , apelmat , nombres"
         ElseIf TxUbiAP <> "" And TxUbiAM <> "" And TxUbiNm <> "" Then
             sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and ApelPat='" & Trim(TxUbiAP.Text) & "' and ApelMat='" & Trim(TxUbiAM.Text) & "' and  Nombres='" & Trim(TxUbiNm.Text) & "'  And ccosto='" & ProgEspecial.Tx_ccosto.Text & "' ORDER BY Apelpat, Apelmat , Nombres"
         End If
      Else
        If TxUbiAP.Text <> "" And TxUbiAM.Text = "" And TxUbiNm.Text = "" Then
           sq2 = "SELECT * FROM CpGeneral WHERE  (baja is NULL) and ApelPat LIKE '" & Trim(TxUbiAP.Text) & "%' ORDER BY Apelpat , apelmat , nombres"
         ElseIf TxUbiAP.Text <> "" And TxUbiAM.Text <> "" And TxUbiNm = "" Then
             sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and ApelPat='" & Trim(TxUbiAP.Text) & "' and ApelMat='" & Trim(TxUbiAM.Text) & "'  ORDER BY Apelpat , apelmat , nombres"
         ElseIf TxUbiAP <> "" And TxUbiAM <> "" And TxUbiNm <> "" Then
             sq2 = "SELECT * FROM CpGeneral WHERE (baja is NULL) and ApelPat='" & Trim(TxUbiAP.Text) & "' and ApelMat='" & Trim(TxUbiAM.Text) & "' and  Nombres='" & Trim(TxUbiNm.Text) & "'  ORDER BY Apelpat, Apelmat , Nombres"
         End If
      End If
      rt.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rt.EOF Then
         MsgBox "No Ubicado", vbExclamation
         rt.Close
         TxUbiAP.Text = "": TxUbiAM.Text = "": TxUbiNm.Text = ""
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
          tGrupos.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
          tGrupos.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
       ElseIf Me.Tag = 0.1 Then
              Workers.Txtarjeta.Text = Grd.Columns(0).Text   'tarjeta
              Workers.Txapelpat.Text = Trim(Grd.Columns(1).Text)
              Workers.Txapelmat.Text = Trim(Grd.Columns(2).Text)
              Workers.TxNombres.Text = Trim(Grd.Columns(3).Text)
              Workers.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 45)
              Workers.TxDni.Text = Grd.Columns(13).Text   'Dni
              Workers.Txocadhep.Text = Grd.Columns(14).Text   'CadHep
       ElseIf Me.Tag = 1 Then
              EditaMarcas.Tx_tarjeta.Text = Grd.Columns(0).Text    'tarjeta
              EditaMarcas.Tx_trabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
       ElseIf Me.Tag = 2 Then
              EditaMarcas.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
              EditaMarcas.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
              EditaMarcas.TxGrpGen.Text = Trim(Grd.Columns(5).Text)
       ElseIf Me.Tag = 3.1 Then
              RecMarcasRS.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
              RecMarcasRS.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
              RecMarcasRS.TxGrpGen.Text = Trim(Grd.Columns(5).Text)
              RecMarcasRS.Txactividad.Text = Grd.Columns(6).Text
              RecMarcasRS.Txccosto.Text = Left(Trim(Grd.Columns(7).Text) + " " + Trim(Grd.Columns(7).Text) + " " + Trim(Grd.Columns(3).Text), 35)
              RecMarcasRS.Txcodupr.Text = Trim(Grd.Columns(4).Text)
              RecMarcasRS.TxCarEst.Text = Trim(Grd.Columns(8).Text)
              RecMarcasRS.TxCargo.Text = Grd.Columns(10).Text
              RecMarcasRS.Txfecnac.Text = Trim(Grd.Columns(11).Text)
              RecMarcasRS.TxDni.Text = Trim(Grd.Columns(13).Text)
       ElseIf Me.Tag = 4 Then
              Autorizaciones.Tx_tarjeta.Text = Grd.Columns(0).Text    'tarjeta
              Autorizaciones.Tx_trabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
       ElseIf Me.Tag = 4.1 Then
              Autorizaciones.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
              Autorizaciones.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
              Autorizaciones.Txcodupr.Text = Trim(Grd.Columns(4).Text)   'codupr
              Autorizaciones.Txccosto.Text = Trim(Grd.Columns(7).Text)
              Autorizaciones.TxCarEst.Text = Trim(Grd.Columns(8).Text)
              Autorizaciones.TxPlH.Text = Trim(Grd.Columns(9).Text)
              Autorizaciones.TxCargo.Text = Trim(Grd.Columns(10).Text)
              Autorizaciones.Txtipopla.Text = Trim(Grd.Columns(12).Text)
       ElseIf Me.Tag = 4.2 Then
              AutorizaQ.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
              AutorizaQ.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
       ElseIf Me.Tag = 4.3 Then
              AutorizaMIN.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
              AutorizaMIN.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
       ElseIf Me.Tag = 5 Then
              CambioTurno.Tx_tarjeta.Text = Grd.Columns(0).Text    'tarjeta
              CambioTurno.Tx_trabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
              CambioTurno.Txcodupr.Text = Trim(Grd.Columns(4).Text)
              CambioTurno.Txccosto.Text = Trim(Grd.Columns(7).Text)
              CambioTurno.Tx_Plh.Text = Trim(Grd.Columns(9).Text)
              CambioTurno.Tx_Cargo.Text = Grd.Columns(10).Text
       ElseIf Me.Tag = 5.1 Then
              CambioTurno.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
              CambioTurno.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
              CambioTurno.Txcodupr.Text = Trim(Grd.Columns(4).Text)   'codupr
              CambioTurno.Txccosto.Text = Trim(Grd.Columns(7).Text)
              CambioTurno.TxCarEst.Text = Trim(Grd.Columns(8).Text)
              CambioTurno.TxPlH.Text = Trim(Grd.Columns(9).Text)
              CambioTurno.TxCargo.Text = Trim(Grd.Columns(10).Text)
       ElseIf Me.Tag = 6.1 Then
              ProgAdministrativo.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
              ProgAdministrativo.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
              ProgAdministrativo.TxGrpGen.Text = Trim(Grd.Columns(5).Text)   'grpgen
              ProgAdministrativo.Txccosto.Text = Left(Trim(Grd.Columns(7).Text) + " " + Trim(Grd.Columns(7).Text) + " " + Trim(Grd.Columns(3).Text), 35)
              ProgAdministrativo.TxCarEst.Text = Trim(Grd.Columns(8).Text)
              ProgAdministrativo.TxPlH.Text = Trim(Grd.Columns(9).Text)
              ProgAdministrativo.TxCargo.Text = Trim(Grd.Columns(10).Text)
              ProgAdministrativo.Txfecnac.Text = Trim(Grd.Columns(11).Text)
       ElseIf Me.Tag = 6.3 Then
              ProgEspecial.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
              ProgEspecial.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
              ProgEspecial.Txcodupr.Text = Trim(Grd.Columns(4).Text)   'codupr
              ProgEspecial.Tx_grupo.Text = Trim(Grd.Columns(5).Text)   'grpgen
              ProgEspecial.Txccosto.Text = Left(Trim(Grd.Columns(7).Text) + " " + Trim(Grd.Columns(7).Text) + " " + Trim(Grd.Columns(3).Text), 35)
              ProgEspecial.TxCarEst.Text = Trim(Grd.Columns(8).Text)
              ProgEspecial.TxPlH.Text = Trim(Grd.Columns(9).Text)
              ProgEspecial.TxCargo.Text = Trim(Grd.Columns(10).Text)
              ProgEspecial.Txfecnac.Text = Trim(Grd.Columns(11).Text)
              ProgEspecial.Txtipopla.Text = Trim(Grd.Columns(12).Text)   'Asi/Adm
        ElseIf Me.Tag = 7 Then
              RptDirectorio.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
              RptDirectorio.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
       'ElseIf Me.Tag = 9 Then
        '     ResPgoGuar.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
        '     ResPgoGuar.TxDni.Text = Grd.Columns(13).Text    'Dni
   End If
       Unload Me
    End If
    Exit Sub
final:
End Sub


