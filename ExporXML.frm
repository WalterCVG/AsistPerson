VERSION 5.00
Begin VB.Form ExporXML 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "ExporXML.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   4080
      Picture         =   "ExporXML.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton CmdXML 
      Caption         =   "Exportar datos a XML"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   360
      Picture         =   "ExporXML.frx":803B0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Genera  C:\RPT\*.xml"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   3855
   End
End
Attribute VB_Name = "ExporXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Select Case wTpoQuery
  Case "5.7"
     Label1.Caption = "Genera C\RPT\DirXML.xml"
  End Select
End Sub
Private Sub CmdXML_Click()
  'Dim Base As String
  Dim SQ1 As String
  Dim Ruta_XML As String
  Dim Ret As Boolean
    
   CreaDir ("c:\rpt")
  
  'Base = "C:\Archivos de programa\Microsoft Visual Studio\VB98\Biblio.MDB"
  Ruta_XML = "C:\RPT\DirXML.xml"
  'Ruta_XML = App.path & "\archivo.xml"
  
  If Dir(Ruta_XML) <> "" Then Kill Ruta_XML
  
  SQ1 = "SELECT * FROM CpTmDIRXML"    '
  '-pasa a la función el path de la base de datos, el path del _
    archivo xml y la consulta sql
       
  Ret = Recordset_a_Xml(SQ1, Ruta_XML)
  
  '-Devuelve verdadero SI lo generó correctamente
  If Ret Then
     'MsgBox " Archivo XMl generado en: " & Ruta_XML, vbInformation
     Call cmdsalir_Click
     Exit Sub
     
     
  End If
End Sub
      
Function Recordset_a_Xml(sql As String, _
                         Path_XML As String) As Boolean
  
  On Error GoTo errSub
  
  '-Variables para la conexión ado, el recordset _
  y el objeto para generar el xml
  Dim obj_DOMDocument As DOMDocument
  'Dim cn As ADODB.Connection
  Dim Rst As ADODB.Recordset
    '-Conexión ADO
    'Set cn = New ADODB.Connection
    'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    '                        "Persist Security Info=False;" & _
    '                        "Data Source=" & Path_BD
       
    '-Abre la base de datos
    'cn.Open
  
    '-Llena el recordset
    'Set rst = cn.Execute(SQL)
  
     Set Rst = New ADODB.Recordset
     Rst.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
    
    '-Graba el contenido del Recordset en el Obj DOMDocument.
    Set obj_DOMDocument = New DOMDocument
    Rst.Save obj_DOMDocument, adPersistXML
       
    '-Cierra el recordset y la conexión a la base de datos
    Rst.Close
    'cn.Close
       
    '-Genera el archivo xml
    obj_DOMDocument.Save Path_XML
       
    Recordset_a_Xml = True
     
Exit Function
  
'-Error
errSub:
  
MsgBox "Error:" & Err.Number & vbNewLine & _
       "Descripción:" & Err.Description, vbCritical
          
       Recordset_a_Xml = False
  
End Function

Private Sub cmdsalir_Click()
  Me.Hide
  Unload Me
End Sub





