VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CargaMarcas 
   Caption         =   "Marcas RR_ddmmaa.txt"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "CargaMarcas.frx":0000
   ScaleHeight     =   3405
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   6480
      TabIndex        =   18
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   130678785
      CurrentDate     =   45149
   End
   Begin VB.TextBox Txdscdni 
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox Tx_dni 
      Height          =   285
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   15
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Solo marcas del servidor          D.O.I. : "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   3015
   End
   Begin MSComctlLib.ProgressBar Progreso 
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "Reloj2"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Reloj1"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   130678785
         CurrentDate     =   41275
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Carga"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Acerca de"
      Height          =   735
      Left            =   120
      Picture         =   "CargaMarcas.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   7320
      Picture         =   "CargaMarcas.frx":80EA8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cerrar"
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton CmdProcesa 
      Height          =   735
      Left            =   6480
      Picture         =   "CargaMarcas.frx":8101A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ejecutar Carga"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton Cmdbuscar 
      Height          =   495
      Left            =   1440
      Picture         =   "CargaMarcas.frx":81324
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Explorador"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog D1 
      Left            =   2640
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "INICIO DE CARGA"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6360
      TabIndex        =   17
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ubica File"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo Origen"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "CargaMarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs_d As New ADODB.Recordset
Dim cPatron, cParametro, ccadena As String
Dim cyear, cmont, cday, creloj, cTarjeta, cfecha, cHora, cFecMin As String
Dim cGrpPer, cGrpGen, cHorario, Ctipo, cIng, cSal, cIng1, cSal1 As String
Dim nposcampm, nDifLDni, nson, nrepro, nyear, nmont, nday, nRechazo As Integer
Dim unafecha As Date
Dim Frechazo, cRuta, clinea, cservidor As String

Private Sub Command1_Click()
     c_grupo = "02": c_secuencia = "026"
     Load AydAcerca
     AydAcerca.Left = 0: AydAcerca.Top = 0
     AydAcerca.Show
End Sub


Private Sub Form_Load()
   CargaMarcas.Caption = "Carga Marcas del Reloj - " & genaño
   DTPicker1.Value = Date - 1
   DTPicker2.Value = Date - 5
   Progreso.Visible = False
   cservidor = "NOVA"
   Set rs = New Recordset
   sql = "SELECT * FROM CPNumeros WHERE tipdoc= 70"
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   cParametro = (rs!parametro)    'Unidad:
   rs.Close
   Set rs = New Recordset
   sql = "SELECT * FROM CPNumeros WHERE tipdoc= 75"
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   cPatron = Trim(rs!parametro) + genaño + "\"    'Carpeta  anual  .BCK
   rs.Close
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
   '01/02_ddmmaa.TXT
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
Private Sub Option1_Click(Index As Integer)
   Select Case Index
      Case 0
         creloj = "01"
         DTPicker1.Value = Date - 1
      Case 1
         creloj = "02"
         DTPicker1.Value = Date - 1
   End Select
End Sub

Private Sub DTPicker1_LostFocus()
    If DTPicker1.Value < Date - 10 And Year(DTPicker1.Value) <> Val(genaño) Then Exit Sub
    If Option1(0).Value = False And Option1(1).Value = False Then
       MsgBox "No indico reloj de marcas...", vbExclamation
       Exit Sub
    End If
    nday = Day(DTPicker1.Value)
    nmont = Month(DTPicker1.Value)
    nyear = Year(DTPicker1.Value)
    cday = Right("00" + Trim(Str(nday)), 2)
    cmont = Right("00" + Trim(Str(nmont)), 2)
    cyear = Right(Trim(Str(nyear)), 2)
    Text1.Visible = True
    '              "C:"     + "\c"+genaño+"\" +  "01_"      + ddmmyy + ".BCK"
    Text1.Text = cParametro + cPatron + creloj + "_" + cday + cmont + cyear + ".TXT"
    If Dir(Text1.Text) > "" Then
      'ubica  dbf de carga
    Else
      MsgBox "File indicado no Existe...  ", vbExclamation
      Text1.Text = ""
      Exit Sub
    End If
    
    Text3.Text = UCase(creloj + "_" + cday + cmont + cyear + ".txt")
    CmdProcesa.Enabled = True
    CmdProcesa.SetFocus
End Sub

Private Sub Check1_Click()
   Tx_dni = "": Txdscdni = ""
   cservidor = "NOVA"
   If Check1.Value = 1 Then
      Tx_dni.Enabled = True
   Else
      Tx_dni.Enabled = False
   End If
End Sub

Private Sub Tx_dni_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Tx_dni = "" Then
        Load AyudaDni
        AyudaDni.Tag = 4
        AyudaDni.Show
     End If
     Set rs = New Recordset
     sql = "SELECT * FROM Cpgeneral where (baja is NULL) and documento='" & Trim(Tx_dni) & "'"
     rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
     If rs.EOF Then
        MsgBox "Documento no ubicado..", vbExclamation
     Else
        Txdscdni = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres) + " ", 35)
        cservidor = Trim(Tx_dni)
     End If
  End If
  If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) And Not KeyAscii = 8 Then
     KeyAscii = 0
  End If
End Sub
Private Sub Tx_dni_LostFocus()
   Call Tx_dni_KeyPress(13)
End Sub
Private Sub CmdProcesa_Click()
  If Text1 = "" Then Exit Sub
  '-Extrae Nuevos reg RR_ddmmaa.Bck a tabla CpMarcaAAAA
  On Error GoTo errSub
  Dim n_File, nlen As Integer
  Dim linea, texto As String
  Dim lectura As String
  Dim path As String
  n_File = FreeFile 'Asignamos el numero de archivo
  path = Text1
  Progreso.Visible = True
  cFecMin = Format(DTPicker2.Value, "yyyymmdd")
  nreg = 0
  
  '-Define Archivo de Escritura
  cRuta = "c:\Rpt\"
  Frechazo = "ErrCarga.txt"
  Dim FSO As New FileSystemObject
  Dim fsoStream As TextStream
  '-Create a text file, and return a reference to a TextStream
  Set fsoStream = FSO.CreateTextFile(cRuta + Frechazo, True)
  
  Open path For Input As n_File   'Abrimos el archivo
  While Not EOF(n_File)           'Leemos todo el archivo linea a linea:
     Line Input #n_File, Garbage  'Leemos un linea.
     nreg = nreg + 1          'Incrementamos el contador de lineas.
  Wend
  Close n_File 'Cerramos el archivo.
  
  Progreso.Value = 0
  Progreso.Max = nreg
  nson = 0: nrepro = 0: nRechazo = 0
  Open path For Input As n_File    'Abrimos el archivo
  '-Montamos un bucle para cada línea del archivo hasta que este se acabe.
  Dim OpenForms
  i = 0
  nposcampm = 0: nDifLDni = 0
  Do While Not EOF(n_File)
  
    '-Dar control a Windows y retomar hilo
    i = i + 1
    If i Mod 20 = 0 Then        'if loop has repeated
       OpenForms = DoEvents     'Yield to operating system
    End If
      
     With Progreso
        If .Value < .Max Then
           .Value = .Value + 1
        End If
     End With
     
     Line Input #n_File, lectura
     nlen = Len(lectura)
     '         1         2         3         4         5         6         7         8         9         10
     '12354678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234
     ' HOSPITAL DE EMERGENCIAS PEDIAT  RODRIGUEZ NORIEGA JULIO   00245393 01/08/2023 07:34:50 p.m.           1                    FP
     'xHOSPITAL DE EMERGENCIAS PEDIAT VASQUEZ MIRANDA DANIEL A 01308011 30/08/2015 07:24:28 a.m.           1            Contraseña
     If nlen > 0 Then
        texto = UCase(Trim(lectura))
        cnomb = Mid(lectura, 33, 24)    '-Caso No existe servidor graba file TXT DNI(8), Nomb(24)
        
        If nposcampm = 0 Then
           'ubica a.m./p.m.  -> posc fija: 87
           For j = 33 To 93
              If Mid(lectura, j, 4) = "a.m." Or Mid(lectura, j, 4) = "p.m." Then
                 nposcampm = j
                 nDifLDni = j - 87
                 
                 'If nson + nrepo < 5 Then
                 '     MsgBox Str(nposcampm) + ";" + Str(nDifLDni)
                 'End If
                 
                 Exit For
              End If
           Next j
        End If
        
        
        cDni = Trim(Mid(lectura, 58 + nDifLDni, 8))
        cTarjeta = ""
        cfecha = Mid(lectura, 67 + nDifLDni, 10)
        cHh = Mid(lectura, 78 + nDifLDni, 2)     'hh:mm: 78-79-80-81-82
        cMm = Mid(lectura, 81 + nDifLDni, 2)
        campm = Mid(lectura, 87 + nDifLDni, 4)
        If cHh <> "12" Then                        '12:45pm
           If campm = "p.m." Then
              cHh = Trim(Str(Val(cHh) + 12))       '13/08=>07:15am,12:15pm,03:00pm 14/08=>12:15am,07:15am
           End If                                  '13/08: 07:15,  12:15,  15:00   14/08: 00:15,  07:15
        Else
           If campm = "a.m." Then
              cHh = "00"
           End If
        End If
        cHora = cHh + cMm
     End If
     
     
     'If cDni = "43390782" And nson + nrepo < 5 Then
     '   MsgBox "[" + cservidor + "-" + cFecMin + "]Dni:" + cDni + "-" + cfecha + ":" + cHora + campm
     'End If
     
     fecha1 = Format(cfecha, "yyyymmdd")
     If fecha1 < fecMin Or (Check1.Value = 1 And cservidor <> cDni) Then
        '-fuera de fecha ó no esel servidor requerido
     Else
        If Val(cHora) > 0 Then
           '-Califica Tipo Marca  I/S
           Set rs_d = New Recordset
           sql = "SELECT * FROM  CpGeneral WHERE  (baja is NULL) and Documento ='" & cDni & "'"
           rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
           If Not rs_d.EOF Then
              cTarjeta = rs_d!tarjeta
              cGrpPer = IIf(IsNull(rs_d!grpPer), "", rs_d!grpPer)
              cGrpGen = IIf(IsNull(rs_d!grpgen), "", rs_d!grpgen)
           Else
              If InStr(ccadena, cDni) = 0 Then
                 nRechazo = nRechazo + 1
                 clinea = "Dni: " + cDni + ";   " + cnomb + "."
                 '-Write to the file texto
                 fsoStream.WriteLine clinea
                 ccadena = ccadena + cDni + "_"
              End If
           End If
           rs_d.Close
           Set rs_d = New Recordset
           sql = "SELECT * FROM  CpHor" & genaño & " WHERE grupo ='" & cGrpPer & "' AND fecha='" & fecha1 & "'"
           rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
           cHorario = ""
           If Not rs_d.EOF Then
              cHorario = IIf(IsNull(rs_d!Horario), "", rs_d!Horario)
           End If
           rs_d.Close
           If cHorario = "" Then
              Set rs_d = New Recordset
              sql = "SELECT * FROM  CpHor" & genaño & " WHERE grupo ='" & cGrpGen & "' AND fecha='" & fecha1 & "'"
              rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
              If Not rs_d.EOF Then
                 cHorario = IIf(IsNull(rs_d!Horario), "", rs_d!Horario)
              End If
              rs_d.Close
           End If

           cIng = "": cSal = "": cIng1 = "": cSal1 = "": Ctipo = ""
           Set rs_d = New Recordset
           sql = "SELECT * FROM  CpHorarioNew WHERE Codigo ='" & Trim(cHorario) & "'"
           rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
           If rs_d.EOF Then
              rs_d.Close
              Set rs_d = New Recordset
              sql = "SELECT * FROM  CpHorarios WHERE Codigo ='" & Trim(cHorario) & "'"
              rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
           End If
         
           If Not rs_d.EOF Then
              cIng = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
              cSal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
              cIng1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
              cSal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
           End If
           rs_d.Close
           If cHorario = "" Then
              Ctipo = "S"
           ElseIf Abs(Val(cHora) - Val(cIng)) < Abs(Val(cHora) - Val(cSal)) Or Abs(Val(cHora) - Val(cIng1)) < 60 Then
                  Ctipo = "I"
           ElseIf Abs(Val(cHora) - Val(cSal)) < Abs(Val(cHora) - Val(cIng)) Then
                  Ctipo = "S"
           End If
        
           Cnx.BeginTrans
           Set rs = New Recordset
           sql = "SELECT * FROM  CpMarca" & genaño & " WHERE  fecha='" & fecha1 & "' AND hora='" & cHora & "' AND tarjeta ='" & cTarjeta & "'"
           rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
           If rs.EOF Then
              rs.AddNew
              CmdProcesa.Caption = "OK " + cDni + "[" + Left(cfecha, 5) + "-" + cHora + "]"
              nson = nson + 1
           Else
              CmdProcesa.Caption = "X/ " + cDni + "[" + Left(cfecha, 5) + "-" + cHora + "]"
              nrepro = nrepro + 1
           End If
           rs!Dni = "" & cDni
           rs!tarjeta = "" & cTarjeta
           rs!fecha = Format(cfecha, "dd/mm/yyyy")
           rs!hora = cHora
           rs!tipo = Ctipo
           rs!usuario = wNick       'wUsuario
           rs!fecRegistro = Date
           rs!HorRegistro = Format(Time, "HH:mm:ss")
           rs!estacion = w_clienName
           rs.Update
           rs.Close
           Cnx.CommitTrans
        End If
     End If
  Loop
  'Cierro archivo de Escritura
  fsoStream.Close
  'Cierro archivo de Lectura
  Close n_File
  CmdProcesa.Caption = ""
  If nson > 0 Then
     Call UltFechaC
     MsgBox "Se Cargo ...  " & nson & " Registros.", , "Barra de Progreso"
  Else
     MsgBox "No hubo carga, reproceso: " & nrepro, , "Barra de Progreso"
  End If
  If nRechazo > 0 Then
     MsgBox "Marca de " + Trim(Str(nRechazo)) + " Servidores no fue CARGADA.. Verificar c:\rpt\ErrCarga.txt", vbCritical
  End If
  Limpiacampos
  Exit Sub
errSub:
    MsgBox "Número de error: " & Err.Number & vbNewLine & _
    "Descripción del error: " & Err.Description, vbCritical
 End Sub
Private Sub Limpiacampos()
  cyear = "": cmont = "": cday = "": cPerCod = "": cfecha = "": cHora = ""
  nson = 0: nrepro = 0: nyear = 0: nmont = 0: nday = 0:
  Text1 = "": Text3 = ""
  DTPicker1.Value = Date - 1
  DTPicker2.Value = Date - 5
  Check1.Value = 0
  Progreso.Visible = False
  Tx_dni.Enabled = False
  Tx_dni = "": Txdscdni = "": cservidor = "NOVA"
End Sub
Private Sub UltFechaC()
     Set rs = New Recordset
     sql = "SELECT * FROM CPNumeros WHERE tipdoc= 80"
     rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
     rs!parametro = cfecha
     rs.Update
     rs.Close
End Sub
Private Sub cmdsalir_Click()
  Set rs = Nothing
  Set rs_d = Nothing
  Set fsoStream = Nothing
  Set FSO = Nothing
  Unload Me
  Me.Hide
End Sub


'++++++++++++++++++++++++++++++++
'' Dim Base As Database
'' Set Base = OpenDatabase(RutaDBF, False, False, "Dbase IV;")
'' Base.Execute "Select * Into " & TablaAccess & " IN '" & RutaBaseAccess & "' from " & ArchivoDBF

'' recuerda que
'' RutaDBF = Es solo el nombre del Directorio donde esta el Archivo DBf
'' TablaAccess = nombre de la Tabla donde se guardaran los Datos en access
'' RutaBaseAccess = ruta de la base de datos con Extension ( .MDB) incluido
'' ArchivoDBF = Nombre de Archivo Dbase con extension (.DBF) incluido

'' La tabla dentro de la base access NO debe existir ,
'' Esta sentencia la genera sola.

'++++++++++++++++++++++++++++++++

'OTRA FORMA
 'rutaDbf = "c:\sistemas\adm\datos\mov"
 'archivoDBF = "HIST" + Trim(Str(genaño)) + ".DBF"
 'tablaaccess = "historia"
 'rutabaseAccess = App.Path
 
 'Dim Base As DatabaseAccess
 'Set Base = OpenDatabase(rutaDbf, False, False, "Dbase IV;")
 'Base.Execute "Select * Into " & tablaaccess & " IN '" & rutabaseAccess & "' from " & archivoDBF


''+++++++++++++++++++++++++++++++
'PRPIEDADES Y METODOS DEL SISTEMA DE ARCHIVOS
'Para utilizar una propiedad o método del sistema de archivo has de recurrir a una variable de objeto como se muestra en el siguiente ejemplo:

'Dim Fso As New FileSystemObject
 
'BuildPath Método
' Construye una ruta completa a partir de un directorio y un nombre de fichero. El símbolo '\' se agregará automáticamente si es preciso
 
'CopyFile Método
' Copia un fichero de una ubicación a otra.
'Sintaxis: Fso.CopyFile {Origen}, {Destino}, {Sobreescribir}
 
'CopyFolder Método
' Copia una carpeta de una ubicación a otra.
'Sintaxis: Fso.CopyFile {Origen}, {Destino}, {Sobreescribir}
 
'CreateFolder Método
' Crea una carpeta.
'Sintaxis: Fso.CreateFolder {Destino}

'Si la carpeta existe devolverá un error.
 
'CreateTextFile Método
' Crea un archivo de texto

'DeleteFile Método
' Elimina un archivo.
'Sintaxis: Fso.DeleteFile {Destino}, {Ignorar}
 
'DeleteFolder Método
' Elimina una carpeta.
'Sintaxis: Fso.DeleteFolder {Destino}, {Ignorar}
 
'DriveExists Propiedad SL
' Devuelve un booleano que expresa si la unidad de almacenamiento pasada como argumento existe en el equipo.
'Sintaxis:
'Bool = Fso.DriveExists("c:\")
 
'FileExists Propiedad SL
' Devuelve un booleano que expresa mediante Verdadero o Falso si el archivo pasado como argumento existe en el equipo.
'Sintaxis:
'Bool = Fso.FileExists({Destino})
 
'FolderExists Propiedad SL
' Devuelve un booleano que expresa mediante Verdadero o Falso si la carpeta pasada como argumento existe en el equipo.
'Sintaxis:
'Bool = Fso.FolderExists({Destino})
 


