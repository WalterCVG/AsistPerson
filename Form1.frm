VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   5400
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cerrar"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton CmdProcesa 
      Height          =   615
      Left            =   4080
      Picture         =   "Form1.frx":0172
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ejecutar Carga"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.CommandButton Cmdbuscar 
      Height          =   495
      Left            =   2280
      Picture         =   "Form1.frx":047C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Explorador"
      Top             =   240
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar Progreso 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog D1 
      Left            =   4560
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo Origen"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ubica File"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim cubi, cdistrito, cprov, cdpto As String
Dim nson As Integer
Dim clinea As String
Private Sub Form_Load()
   Me.Caption = "Carga Ubigeo "
    'Dim fso, f
    'Set fso = CreateObject("Scripting.FileSystemObject")
    'Set f = fso.GetFile("c:\otraruta\mifichero.txt")
    'f.Copy "c:\mirutadestino"
   CmdProcesa.Enabled = True
End Sub
Private Sub Cmdbuscar_Click()
 D1.ShowOpen
 If D1.FileName <> "" Then
   Text1.Text = UCase(D1.FileName)
   Text3.Text = Dir(D1.FileName)
   If Right(UCase(D1.FileName), 4) = ".TXT" Then
      Text1.Text = D1.FileName
      Text1.Visible = True
      CmdProcesa.Enabled = True
      CmdProcesa.SetFocus
   Else
     MsgBox "Verifique archivo Txt...  " + Right(D1.FileName, 13), vbExclamation
   End If
 Else
   MsgBox "No Selecciono File...  ", vbExclamation
 End If
End Sub


Private Sub CmdProcesa_Click()
  If Text1 = "" Then Exit Sub
  On Error GoTo errSub
  Dim n_File, nlen As Integer
  Dim linea, texto As String
  Dim lectura As String
  Dim path As String
  n_File = FreeFile 'Asignamos el numero de archivo
  path = Text1.Text
  nreg = 0
  
  Open path For Input As n_File   'Abrimos el archivo
  While Not EOF(n_File)           'Leemos todo el archivo linea a linea:
     Line Input #n_File, Garbage  'Leemos un linea.
     nreg = nreg + 1          'Incrementamos el contador de lineas.
  Wend
  Close n_File 'Cerramos el archivo.
  
  Progreso.Value = 0
  Progreso.Max = nreg
  nson = 0
  Open path For Input As n_File    'Abrimos el archivo
  '-Montamos un bucle para cada línea del archivo hasta que este se acabe.
  Do While Not EOF(n_File)
     With Progreso
        If .Value < .Max Then
           .Value = .Value + 1
        End If
     End With
     
     Line Input #n_File, lectura
     nlen = Len(lectura)
     '         1         2         3         4         5         6         7         8
     '12354678901234567890123456789012345678901234567890123456789012345678901234567890123456789
     '999999 xxxxxxxxx1xx3 xxxxxxxxx1xxxxxxxxx2xx3 xxxxxxxxx1xxxxxxxxx2xxxxxxxxx3xxxxx6
     If nlen > 0 Then
        texto = UCase(Trim(lectura))
        cubi = Left(lectura, 6)
        cdpto = Trim(Mid(lectura, 8, 13))
        cprov = Trim(Mid(lectura, 22, 23))
        cdistrito = Trim(Mid(lectura, 46, 36))
     End If
             
           Cnx0.BeginTrans
           Set rs = New Recordset
           sql = "SELECT * FROM  ubigeonew "
           rs.Open sql, Cnx0, adOpenStatic, adLockOptimistic, adCmdText
              rs.AddNew
              CmdProcesa.Caption = "OK [" + cubi + "]"
              
           rs!ubi_geogr = cubi
           rs!Distrito = cdistrito
           rs!prov = cprov
           rs!depto = cdpto
           rs.Update
           rs.Close
           Cnx0.CommitTrans
           nson = nson + 1
  Loop
  'Cierro archivo de Lectura
  Close n_File
  CmdProcesa.Caption = ""
  If nson > 0 Then
     MsgBox "Se Cargo ...  " & nson & " Registros.", , "Barra de Progreso"
  End If
  Limpiacampos
  Exit Sub
errSub:
    MsgBox "Número de error: " & Err.Number & vbNewLine & _
    "Descripción del error: " & Err.Description, vbCritical
 End Sub
Private Sub Limpiacampos()
  cdistrito = "": cprov = "": cdpto = ""
  nson = 0
  Text1.Text = ""
  
End Sub
Private Sub cmdsalir_Click()
  Set rs = Nothing
  Set FSO = Nothing
  Unload Me
  Me.Hide
End Sub

