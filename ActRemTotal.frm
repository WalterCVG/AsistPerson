VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCx"
Begin VB.Form ActRemTotal 
   Caption         =   "Actualiza RemuneraciónTotal desde PLH"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "ActRemTotal.frx":0000
   ScaleHeight     =   3135
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "CAMPOS DE TRABAJO"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   7455
      Begin VB.TextBox TxNivel 
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxRemunera 
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxDNI 
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Nivel:"
         Height          =   255
         Left            =   5520
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Remuneraciòn :"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ID. Servidor :"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComctlLib.ProgressBar Progreso 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   7320
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   7920
      Picture         =   "ActRemTotal.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cerrar"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton CmdProcesa 
      Caption         =   "Procesa"
      Height          =   615
      Left            =   7800
      Picture         =   "ActRemTotal.frx":803D8
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ejecuta Act."
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Cmdbuscar 
      Height          =   495
      Left            =   4800
      Picture         =   "ActRemTotal.frx":806E2
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Buscar File"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   5055
   End
   Begin MSComDlg.CommonDialog D1 
      Left            =   6480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicar tabla de datos (.dbf)"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo Origen"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "ActRemTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rst As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim sql, sql2, cDni, cremunera, cNivel, clibele, cid, cou, cdscou, cservidor As String
Dim nson, npos, npos1, npos2, npos3, nposO, nposDO, nposS As Integer
Dim nremunera As Double
Private Sub Form_Load()
   ActRemTotal.Caption = "Actualizar Nivel y Remuneración Total desde PLH "
   Set rs = New Recordset
   CmdProcesa.Enabled = False
End Sub
Private Sub Cmdbuscar_Click()
   D1.ShowOpen
   If D1.FileName <> "" Then
      Text1 = UCase(D1.FileName)
      Text3 = UCase(Dir(D1.FileName))
      If Right(UCase(D1.FileName), 4) <> ".DBF" Then
         MsgBox "No es un archivo DBF .." + D1.FileName, vbCritical
         Text1 = "":         Text3 = ""
         Text1.Visible = True
         TxDNI = "":            TxRemunera = ""
         Exit Sub
      End If
         Text1 = UCase(D1.FileName)
         Text1.Visible = True
         CmdProcesa.Enabled = True
         CmdProcesa.SetFocus
   Else
      MsgBox "No Selecciono File...  ", vbExclamation
   End If
End Sub

Private Sub CmdProcesa_Click()
  If Text1 = "" Then Exit Sub
  
  If Trim(TxDNI) = "" Then
     MsgBox "No Campo ID. para selecciòn de datos..", vbCritical
     TxDNI = "": TxRemunera = "": TxNivel = ""
     Exit Sub
  End If
  If Trim(TxRemunera) = "" And Trim(TxNivel) = "" Then
     MsgBox "No indica campos de datos, para la actualizaciòn..", vbCritical
     TxDNI = "": TxRemunera = "": TxNivel = ""
     Exit Sub
  End If
  
  cDni = UCase(TxDNI)
  cremunera = UCase(TxRemunera)
  cNivel = UCase(TxNivel)
    
  npos = InStr(Trim(Text1), Text3)
  path = Left(Text1, npos - 1)
  'Actualiza RemTot desde DataNomB.dbf a tabla CpGreneral
  Dim nreg, i, X, j As Integer
  Set Cndos = New ADODB.Connection
  With Cndos
    .Provider = "Microsoft.Jet.OLEDB.4.0"
    .ConnectionString = "Data Source = " & path
    .Properties("Extended Properties") = "dBASE IV;"
    .Open
  End With
  sql = "select * from " + Text3
  Dim Rst As ADODB.Recordset
  Set Rst = New ADODB.Recordset
  With Rst
    .CursorLocation = adUseClient
    .CursorType = adOpenStatic
    .LockType = adLockReadOnly
    .Open sql, Cndos, , , adCmdText
  End With

'-El control ProgressBar nos permite medir el progreso de una tarea o proceso en nuestro programa
'-Este control viene con el conjunto de controles que se incorporan en el ocx MsComctl32.ocx con el
'-nombre de Microsoft Windows Common Control 6
  
  
  'Actualiza RemTot    Dbf  => SQL
  nreg = Rst.RecordCount
  Rst.MoveFirst
  '-check existe campos clave y dato
  npos1 = 88: npos2 = 88: npos3 = 88: nposO = 88: nposS = 88: nposDS = 88
  cid = ""
  For i = 0 To Rst.Fields.Count - 1
         'On Error Resume Next
         ccampo = UCase(Rst.Fields(i).Name)
         If ccampo = cDni Then
            npos1 = i
            cid = cDni
         ElseIf cremunera <> "" And ccampo = cremunera Then
                npos2 = i
         ElseIf cNivel <> "" And ccampo = cNivel Then
                npos3 = i
         'ElseIf ccampo = "APELLIDOS_" Then
         '      nposS = i
         'ElseIf ccampo = "O_OU" Then
         '     nposO = i
         'ElseIf ccampo = "DSCOU" Then
         '     nposDO = i
         End If
  Next i
  If cid = "" Or (cid <> "" And (npos2 = 88 And npos3 = 88)) Then
     MsgBox "Verifique, hay campos indicados que no existen..", vbCritical
     Rst.Close
     Set Cndos = Nothing
     Exit Sub
  End If
  
  
  Progreso.Value = 0
  Progreso.Max = nreg
  nson = 0: nvan = 0
  Sw = True
  Do While Not Rst.EOF
     With Progreso
        If .Value < .Max Then
           .Value = .Value + 1
        End If
     End With
     nvan = nvan + 1
     clibele = ""
     nremunera = 0
    '' On Error Resume Next
     clibele = Trim(Rst.Fields(npos1))
     If npos2 <> 88 And npos3 <> 88 Then
        nremunera = IIf(IsNull(Rst.Fields(npos2)), 0, Rst.Fields(npos2))
        cNivel = Trim(Rst.Fields(npos3))
        sq2 = "UPDATE CpGeneral  SET actRemu='" + wUsuario + "', fecremu='" + Format(Date, "yyyymmdd") + "', horremu='" + Format(Time, "HH:mm:dd") + "', codniv='" & cNivel + "', RemTot=" & nremunera & " WHERE Documento='" & clibele & "'"
     ElseIf npos2 <> 88 And npos3 = 88 Then
             nremunera = IIf(IsNull(Rst.Fields(npos2)), 0, Rst.Fields(npos2))
             sq2 = "UPDATE CpGeneral  SET actRemu='" + wUsuario + "', fecremu='" + Format(Date, "yyyymmdd") + "', horremu='" + Format(Time, "HH:mm:dd") + "', RemTot=" & nremunera & " WHERE Documento='" & clibele & "'"
     ElseIf npos2 = 88 And npos3 <> 88 Then
            cNivel = Trim(Rst.Fields(npos3))
            sq2 = "UPDATE CpGeneral  SET actRemu='" + wUsuario + "', fecremu='" + Format(Date, "yyyymmdd") + "', horremu='" + Format(Time, "HH:mm:dd") + "', codniv='" & cNivel & "' WHERE Documento='" & clibele & "'"
     End If
     
     ' cou = IIf(IsNull(Rst.Fields(nposO)), "", Rst.Fields(nposO))
     ' cdscou = IIf(IsNull(Rst.Fields(nposDO)), "", Rst.Fields(nposDO))
     ' cservidor = IIf(IsNull(Rst.Fields(nposS)), "", Rst.Fields(nposS))
     ' sq2 = "UPDATE CpGeneral  SET OU_plh='" & cou & "',s_ervidor ='" & cservidor + "', actRemu='" + wUsuario + "', fecremu='" + Format(Date, "yyyymmdd") + "', horremu='" + Format(Time, "HH:mm:dd") + "' WHERE Documento='" & clibele & "'"
     If clibele <> "" Then
        Cnx.Execute sq2
        nson = nson + 1
        CmdProcesa.Caption = "Van : " + CStr(nson)
        
        If npos3 <> 88 Then
           sq2 = "Select * from cpNivel WHERE codNivel='" & cNivel + "'"
           rs.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
           If rs.EOF Then
              rs.AddNew
              rs!codnivel = cNivel
              rs.Update
           End If
           rs.Close
        End If
        
        'sq2 = "Select * from OUPLH WHERE codigo='" & cou + "'"
        'rs.Open sq2, Cnx, adOpenStatic, adLockOptimistic, adCmdText
        '   If rs.EOF Then
        '     rs.AddNew
        '     rs!codigo = cou
        '     rs!descripcion = cdscou
        '     rs.Update
        '  End If
        '  rs.Close
        
     End If
     Rst.MoveNext
  Loop
  Rst.Close
  CmdProcesa.Caption = ""
  If nson > 0 Then
     MsgBox "Se Actualizo ...  " & nson & " Registros.", , "Barra de Progreso"
  Else
     MsgBox "No hubo reproceso ", , "Barra de Progreso"
  End If
  Set Cndos = Nothing
  Limpiacampos
End Sub
Private Sub Limpiacampos()
   nson = 0
   Text1 = "": Text3 = ""
   TxDNI = "": TxRemunera = "": TxNivel = ""
   cDni = "": cremunera = "": cNivel = ""
   CmdProcesa.Enabled = False
    Text1.Visible = False
End Sub
Private Sub cmdsalir_Click()
  Set Rst = Nothing
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
 


