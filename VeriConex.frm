VERSION 5.00
Begin VB.Form VeriConex 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "VeriConex.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   8250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Verificar"
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   4695
      Begin VB.TextBox TexFin 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox TexInicio 
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   22
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox TxArchivo 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   6840
      Picture         =   "VeriConex.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cerrar"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "Graba"
      Height          =   615
      Left            =   5160
      Picture         =   "VeriConex.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Grabar"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox TxPwd 
      Height          =   285
      Left            =   5400
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox TxUsr 
      Height          =   285
      Left            =   5400
      TabIndex        =   16
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox TxIcatal 
      Height          =   285
      Left            =   5400
      TabIndex        =   12
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox TxIcatal0 
      Height          =   285
      Left            =   5400
      TabIndex        =   11
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox TxProvider 
      Height          =   285
      Left            =   5400
      TabIndex        =   8
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox TxMac 
      Height          =   285
      Left            =   5400
      TabIndex        =   6
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox TxDsource 
      Height          =   285
      Left            =   5400
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   15
      Left            =   5280
      TabIndex        =   15
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Catalogo"
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Maestra HEP"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Provider"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mac"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contenido"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "VeriConex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim obj_FSO As Object
Dim archivo As Object
Private Sub Form_Load()
   TxArchivo.Text = ""
   'TxArchivo.SetFocus
End Sub

Private Sub Txarchivo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Trim(TxArchivo.Text) = "" Then
         MsgBox ("Indique Archivo de Configuración")
         Exit Sub
      End If
      strExcelFile = TxArchivo.Text
      If Dir(strExcelFile) = "" Then
         MsgBox ("Archivo No Existe")
         Exit Sub
      End If
      Set obj_FSO = CreateObject("Scripting.FileSystemObject")
      Set archivo = obj_FSO.OpenTextFile(TxArchivo.Text, 1)
      Text1 = archivo.ReadAll
      archivo.Close
      Call leer(TxArchivo.Text)
   End If
End Sub
   
Private Sub leer(archivo As String)
'- variable para fso
   Dim Obj_TextStream As Scripting.TextStream
   Dim obj_FSO As New Scripting.FileSystemObject
   Dim Obj_File As File
   Dim Linea_Actual As String
   Set Obj_File = obj_FSO.GetFile(archivo)
   '- Lo abre con OpenAsTextStream
   Set Obj_TextStream = Obj_File.OpenAsTextStream(ForReading, TristateUseDefault)
   '- recorre todo el contenido del fichero
   Do While Not Obj_TextStream.AtEndOfStream
      '- lee la linea
      Linea_Actual = Obj_TextStream.ReadLine
      npos = InStr(Linea_Actual, "=")
      nlen = Len(Linea_Actual)
      If npos > 0 Then
         texto = UCase(Trim(Left(Linea_Actual, npos - 1)))
         Encclave = ""
         WVdesencripta (Right(Linea_Actual, nlen - npos - 1))
         Select Case texto
            Case "2"
               TxDsource.Text = Encclave
            Case "3"
               TxMac.Text = Encclave
            Case "4"
               TxProvider.Text = Right(Linea_Actual, nlen - npos - 1)
            Case "5"
               TxIcatal0.Text = Encclave
            Case "6"
               TxIcatal.Text = Encclave
            Case "7"
               TxUsr.Text = Encclave
            Case "8"
               TxPwd.Text = Encclave
         End Select
      End If
      
'      Printer.Print Linea_Actual ' imprime con
   Loop
'   Printer.EndDoc
End Sub
Private Sub cmdGraba_Click()
  'Dim obj_FSO As Object
  'Dim Archivo As Object
  ' Set obj_FSO = CreateObject("Scripting.FileSystemObject")
  '-Creamos un archivo con el método CreateTextFile
  Set archivo = obj_FSO.CreateTextFile(TxArchivo.Text, True)
  '-ESCRIBIMOS lineas
  ' Archivo.WriteLine "Una linea"
  ' Archivo.WriteLine "Otra linea"
  ' Archivo.WriteLine "Otra mas"
  For i = 1 To 9
      Encclave = ""
      Select Case i
         Case 1
           archivo.WriteLine "[DB_hepdata]"
         Case 2
           WVencripta (TxDsource.Text)
           archivo.WriteLine "2 = " + Trim(Encclave)
         Case 3
           WVencripta (TxMac.Text)
           archivo.WriteLine "3 = " + Trim(Encclave)
         Case 4
           'WVencripta (TxProvider.Text)
           archivo.WriteLine "4 = " + Trim(TxProvider.Text)
         Case 5
           WVencripta (TxIcatal0.Text)
           archivo.WriteLine "5 = " + Trim(Encclave)
         Case 6
           WVencripta (TxIcatal.Text)
           archivo.WriteLine "6 = " + Trim(Encclave)
         Case 7
           WVencripta (TxUsr.Text)
           archivo.WriteLine "7 = " + Trim(Encclave)
         Case 8
           WVencripta (TxPwd.Text)
           archivo.WriteLine "8 = " + Trim(Encclave)
         Case 9
           archivo.WriteLine "[File]"
      End Select
  Next i
   
'-Cerramos el fichero
 archivo.Close

End Sub
Private Sub Texinicio_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Encclave = ""
      WVencripta (TexInicio.Text)
      TexFin.Text = Encclave
   End If
End Sub

Private Sub cmdsalir_Click()
   'Archivo.Close
   Set obj_FSO = Nothing
   Set Obj_TextStream = Nothing
   Set Obj_File = Nothing
   Set archivo = Nothing
   Me.Hide
   Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
   'Archivo.Close
   Set obj_FSO = Nothing
   Set Obj_TextStream = Nothing
   Set Obj_File = Nothing
   Set archivo = Nothing
   Me.Hide
   Unload Me
End Sub



'Dim obj_FSO As Object
'Dim Archivo As Object
' Set obj_FSO = CreateObject("Scripting.FileSystemObject")
'-Creamos un archivo con el método CreateTextFile
' Set Archivo = obj_FSO.CreateTextFile(App.path & "MiArchivo.txt", True)
'-ESCRIBIMOS lineas
' Archivo.WriteLine "Una linea"
' Archivo.WriteLine "Otra linea"
' Archivo.WriteLine "Otra mas"
'-Cerramos el fichero
' Archivo.Close
  

'-Para ABRIR el archivo
' Set Archivo = obj_FSO.OpenTextFile(App.path & "MiArchivo.txt", 1)
  
'-CARGAMOS en el TextBox todo el archivo
' Text1 = Archivo.ReadAll
  
'///Nota: para leer una linea se utiliza ReadLine

' Archivo.Close
'  Set obj_FSO = Nothing
'  Set Archivo = Nothing

'+++++++++++++++++++++++++++++++
'-Otro ejemplo - Método TextStream
'-Hay dos funciones para LEER y GRABAR un fichero de texto utilizando el objeto TextStream de FSO
'-El ejemplo contiene dos funciones de ajuste. Una llamada LeerArchivo y otra guardarArchivo.
'-A la primera se le debe indicar el path , y la función retorna un String con los datos del fichero.
'-La función para guardar o crear , retorna un valor Boolean, si retorna true se pudo guardar el fichero
'-correctamente, en caso contrario la función retorna False

'-Nota: incluir desde el menú de referencia de visual basic la referencia a : Microsft Scripting Runtime.

'Option Explicit
'- Objeto FileSystemObject
'Dim fso As New FileSystemObject
'Public Function GuardarArchivo(Ruta As String, Contenido As String) As Boolean
'    Dim ObjDatos As TextStream     '- Objeto Fso para el Contenido del archivo
'    Dim Datos As String            '- datos del fichero
'    On Error GoTo Error_Function:
'    Datos = Contenido
    '- crea el archivo en la ruta indicada
'   Set ObjDatos = fso.CreateTextFile(Ruta, True, True)
    '- escribe los datos
'    ObjDatos.Write (Datos)
    '- cierra el archivo abierto
'    ObjDatos.Close
'    GuardarArchivo = True
'    Set fso = Nothing
'    Exit Function
'Error_Function:
'    GuardarArchivo = False
'    On Local Error Resume Next
'    Set fso = Nothing
'    MsgBox Err.Description
'End Function
  
'Public Function LeerArchivo(Ruta As String) As String
'    On Error GoTo Error_Function:
'    Dim ObjDatos As TextStream     '- Objeto Fso para el Contenido del archivo
'    Dim Datos As String            '- datos del fichero
'    Datos = ""
'    Set ObjDatos = fso.OpenTextFile(Ruta, ForReading, False, TristateMixed)
'    While Not ObjDatos.AtEndOfStream
'          Datos = Datos & ObjDatos.ReadAll
'    Wend
    '- retorna el contenido leido
'    LeerArchivo = Datos
    '- cierra el archivo abierto
'    ObjDatos.Close
'    Set fso = Nothing
'    Exit Function
'Error_Function:
'    MsgBox Err.Description
'    On Local Error Resume Next
'    Set fso = Nothing
'End Function

'-Luego para utilizar las funciones :
'-Para GRABAR un nuevo archivo se le pasa la ruta y los datos:
'- GuardarArchivo "c:\archivo.txt", "Texto de prueba"
'-Para LEER y visualizarlo en un textBox , la ruta :
'- Text1 = LeerArchivo("c:\archivo.txt")

'+++++++++++++++++++++++++++++++++++++
'-Leer linea por linea un archivo y mandarlo a imprimir
'-En este caso se usa el método ReadLine del objeto TextStream para ir leyendo las líneas
'-del fichero, y luego con printer se imprime

'Private Sub Imprimir(Archivo As String)
'- variable para fso
'Dim Obj_TextStream As Scripting.TextStream
'Dim Obj_Fso As New Scripting.FileSystemObject
'Dim Obj_File As File
'Dim Linea_Actual As String
   '- Referencia al archivo con GetFile
'   Set Obj_File = Obj_Fso.GetFile(Archivo)
   '- Lo abre con OpenAsTextStream
'   Set Obj_TextStream = Obj_File.OpenAsTextStream(ForReading, TristateUseDefault)
   '- recorre todo el contenido del fichero
'   Do While Not Obj_TextStream.AtEndOfStream
      '- lee la linea
'      Linea_Actual = Obj_TextStream.ReadLine
'      Printer.Print Linea_Actual ' imprime con
'   Loop
'   Printer.EndDoc
'End Sub

