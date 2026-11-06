VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Ayuda0 
   Caption         =   "PENDIENTES de Programaciòn"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "Ayuda0.frx":0000
   ScaleHeight     =   7350
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txcodupr 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Txmes 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "Ayuda0.frx":80266
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      Picture         =   "Ayuda0.frx":806A8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox SSPanel1 
      BackColor       =   &H00C0C0C0&
      Height          =   6615
      Left            =   120
      ScaleHeight     =   6555
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   600
      Width           =   8115
      Begin MSDataGridLib.DataGrid Grd 
         Height          =   6135
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   10821
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
         Caption         =   "Miembros de la UPR"
         ColumnCount     =   12
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
            DataField       =   "PLH"
            Caption         =   "PlH"
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
            DataField       =   "ccosto"
            Caption         =   "Ung. ORG"
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
            DataField       =   "Cargo"
            Caption         =   "Cargo."
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
            DataField       =   "GrpGen"
            Caption         =   "GrpGen"
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
            DataField       =   "fecnac"
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
         BeginProperty Column11 
            DataField       =   "Cadhep"
            Caption         =   "cadena"
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
         EndProperty
      End
   End
End
Attribute VB_Name = "Ayuda0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim r_s As New ADODB.Recordset
Dim r_d As New ADODB.Recordset
Dim DD As New ADODB.Recordset
Dim sql, sq2 As String

Private Sub Form_Load()
  If ProgASISTENCIAL.Tx_codupr = "" Then
     MsgBox "<E> No indico UPR...", vbExclamation
     Exit Sub
  End If
  Txmes = Trim(ProgASISTENCIAL.Txmes)
  Txcodupr = Trim(ProgASISTENCIAL.Tx_codupr)
  
 '--MUESTRA  registros de CPGENERAL (codupr='11' and estado='A'), que no figuran en CPDIRECTORIO (ANO-MES-TARJETA)
 '-SELECT     t1.Tarjeta, t1.CodUPR, t1.Estado, t2.tarjeta AS cod
 '-FROM         CpGeneral AS t1 LEFT OUTER JOIN
 '-                     CPdire AS t2 ON t2.tarjeta = t1.Tarjeta
 '-WHERE     (t2.tarjeta IS NULL) AND (t1.estado = 'A') AND (t1.CodUPR = '11')

   '-USO TEMPORAL
   Cnx.BeginTrans
'On Error GoTo CREA
'   Cnx.Execute "DELETE  FROM #tabla3 "
'   GoTo Sigue
'CREA:
   Cnx.Execute ("SELECT * INTO #tabla3 FROM Cpdirectorio WHERE ano ='" & genaño & "' AND mes='" & Txmes & "'  and codupr = '" + Txcodupr + "'")
'Sigue:
    Cnx.CommitTrans
   
  'sq2 = "SELECT   t1.*, t2.tarjeta AS cod " & _
  '"FROM CpGeneral AS t1 LEFT OUTER JOIN " & _
  '" #tabla3 AS t2 ON t2.tarjeta = t1.Tarjeta " & _
  '"WHERE   (t2.tarjeta IS NULL) AND (t1.estado = 'A') AND (t1.CodUPR = '" + ProgASISTENCIAL.Tx_codupr & "') and (tipopla='ASI') " & _
  '"ORDER BY Apelpat , apelmat , nombres"
  
  sq2 = "SELECT  t1.*, #tabla3.tarjeta AS cod " & _
  "FROM CpGeneral AS t1 LEFT OUTER JOIN " & _
  " #tabla3 ON #tabla3.tarjeta = t1.Tarjeta " & _
  "WHERE   (t1.baja is NULL) and  (#tabla3.tarjeta IS NULL) AND (t1.estado = 'A') AND (t1.CodUPR = '" + Txcodupr & "') and (tipopla='ASI') " & _
  "ORDER BY Apelpat , apelmat , nombres"
  
   Set DD = New Recordset
   With DD
     .Fields.Append "tarjeta", adChar, 5
     .Fields.Append "apelpat", adChar, 25
     .Fields.Append "apelmat", adChar, 25
     .Fields.Append "nombres", adChar, 25
     .Fields.Append "codupr", adChar, 2
     .Fields.Append "Plh", adChar, 3
     .Fields.Append "ccosto", adChar, 2
     .Fields.Append "cargo", adChar, 4
     .Fields.Append "carEst", adChar, 2
     .Fields.Append "grpgen", adChar, 4
     .Fields.Append "fecnac", adDate
     .Fields.Append "cadhep", adChar, 52
     .Fields.Append "Director", adChar, 1
     .Open
  End With
  Set Grd.DataSource = DD
    
  Call Command1_Click
      
End Sub
Private Sub Command1_Click()
   '---MsgBox sq2
   Dim rt As Recordset
   Set rt = New Recordset
   rt.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If rt.EOF Then
      MsgBox "No Ubicados en UPR " + Txcodupr, vbExclamation
      rt.Close
      Exit Sub
   Else
      wdato = 0
      Do While Not rt.EOF
           DD.AddNew
           DD!tarjeta = rt!tarjeta
           DD!apelpat = IIf(IsNull(rt!apelpat), "", rt!apelpat)
           DD!apelmat = IIf(IsNull(rt!apelmat), "", rt!apelmat)
           DD!nombres = IIf(IsNull(rt!nombres), "", rt!nombres)
           DD!CodUpr = IIf(IsNull(rt!CodUpr), "", rt!CodUpr)
           DD!PLH = IIf(IsNull(rt!PLH), "", rt!PLH)
           DD!CCosto = IIf(IsNull(rt!CCosto), "", rt!CCosto)
           DD!Cargo = IIf(IsNull(rt!Cargo), "", rt!Cargo)
           DD!carEst = IIf(IsNull(rt!carEst), "", rt!carEst)
           DD!grpgen = IIf(IsNull(rt!grpgen), "", rt!grpgen)
           If IsDate(rt!fecnac) Then DD!fecnac = rt!fecnac
           DD!CadHep = IIf(IsNull(rt!CadHep), "", rt!CadHep)
           DD.Update
           wdato = wdato + 1
           rt.MoveNext
      Loop
      Grd.Row = 0
      Grd.Col = 0
      'Grd.SetFocus
   End If
   rt.Close
End Sub

Private Sub Grd_DblClick()
    Call Grd_KeyDown(13, 0)
End Sub
Private Sub Grd_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo final
   If KeyCode = 13 Then
      If Me.Tag = 6.2 Then
             ProgASISTENCIAL.Txtarjeta.Text = Grd.Columns(0).Text    'tarjeta
             ProgASISTENCIAL.Txtrabajador.Text = Left(Trim(Grd.Columns(1).Text) + " " + Trim(Grd.Columns(2).Text) + " " + Trim(Grd.Columns(3).Text), 35)
             ProgASISTENCIAL.Txcodupr.Text = Trim(Grd.Columns(4).Text)   'codupr
             ProgASISTENCIAL.TxPlH.Text = Trim(Grd.Columns(5).Text)
             ProgASISTENCIAL.Txccosto.Text = Trim(Grd.Columns(6).Text)
             ProgASISTENCIAL.TxCargo.Text = Trim(Grd.Columns(7).Text)
             ProgASISTENCIAL.TxCarEst.Text = Trim(Grd.Columns(8).Text)
             ProgASISTENCIAL.Tx_grupo.Text = Trim(Grd.Columns(9).Text)   'grpgen
             ProgASISTENCIAL.Txfecnac.Text = Trim(Grd.Columns(10).Text)
             ProgASISTENCIAL.TxCadHep.Text = Trim(Grd.Columns(11).Text)
             ProgASISTENCIAL.TxDirector.Text = Trim(Grd.Columns(12).Text)
                If IsDate(Grd.Columns(13).Text) Then ProgASISTENCIAL.TxFecIngr = Trim(Grd.Columns(13).Text)
       End If
      Unload Me
   End If
   Exit Sub
final:
End Sub
Private Sub Grd_HeadClick(ByVal ColIndex As Integer)
   With DD      ' Adodc1.Recordset
      If (.Sort = .Fields(ColIndex).[Name] & " Asc") Then
         .Sort = .Fields(ColIndex).[Name] & " Desc"
      Else
         .Sort = .Fields(ColIndex).[Name] & " Asc"
      End If
   End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Cnx.Execute "Drop table #tabla3"
End Sub

