VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCx"
Begin VB.Form ResPGText 
   Caption         =   "Archivo texto  Pago x Guardia"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "ResPGText.frx":0000
   ScaleHeight     =   3360
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Progreso 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CmdTXTSUNAT 
      Caption         =   "Generar Archivo Plano"
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
      Left            =   2040
      Picture         =   "ResPGText.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   6480
      Picture         =   "ResPGText.frx":80D20
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Salir"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Generar.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label LblFileTxt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1047MMAA.txt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   2400
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "en Carpeta C:\RPT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   3255
   End
End
Attribute VB_Name = "ResPGText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clinea, cano, cmes, mitexto As String
Dim nson As Integer
Dim dfecha As Date
Dim Rsd As ADODB.Recordset
Private Sub Form_Load()
  LblFileTxt.Caption = "1047" + wPer_resol + ".txt"     'ResPgoGuarProy.Txmes + Right(Trim(Str(genaño)), 2)
  cperiodo = "1047" + wPer_resol
  '-Estructura   tabla 1047mmaa.dbf
  'codigo chr(8),    libele chr(8),   Nombre chr(45),    'Monto chr(9)   sin decimales  1,245.40  => 124540
  '13/03/2023  codigo chr(8),    libele chr(9),   Nombre chr(45),    'Monto chr(9)   sin decimales  1,245.40  => 124540
End Sub
Private Sub CmdTxtSunat_Click()
  Dim obj_FSO As Object
  Dim archivo As Object
  Set Rsd = New Recordset
  Sq0 = "SELECT * FROM CpPG1047 ORDER BY nombre"
  Rsd.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  nreg = Rsd.RecordCount
  Progreso.Value = 0
  Progreso.Max = nreg
  Rsd.MoveFirst
  nson = 0
  MiFile = Trim(LblFileTxt.Caption)     ' "1047" + wPer_resol + ".txt"
  Set obj_FSO = CreateObject("Scripting.FileSystemObject")
  '-Creamos un archivo con el método CreateTextFile
  Set archivo = obj_FSO.CreateTextFile("c:\rpt\" & MiFile, True)
  Do While Not Rsd.EOF
     nson = nson + 1
     clinea = ""
     With Progreso
        If .Value < .Max Then
           .Value = .Value + 1
        End If
     End With
     If Rsd!Monto = "00" Then
     Else
        '-Formamos texto y escribimos linea
        clinea = Rsd!codigo + Left(Trim(Rsd!libele) + Space(9), 9) + Left(Trim(Rsd!nombre) + Space(45), 45) + Right(Space(9) + Trim(Str(Rsd!Monto)), 9)
        archivo.WriteLine clinea
     End If
     Rsd.MoveNext
  Loop
  '-Cerramos el fichero
  archivo.Close
  Rsd.Close
  MsgBox ("Proceso Concluido..!")
  Call cmdsalir_Click
End Sub
Private Sub cmdsalir_Click()
  Set Rsd = Nothing
  Me.Hide
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set Rsd = Nothing
  Me.Hide
  Unload Me
End Sub


