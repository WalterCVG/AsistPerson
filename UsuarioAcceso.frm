VERSION 5.00
Begin VB.Form UsuarioAcceso 
   Caption         =   "Accesos"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "UsuarioAcceso.frx":0000
   ScaleHeight     =   5235
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Chk2 
      Caption         =   "Nutrición- Directorio Asisten."
      Height          =   350
      Index           =   8
      Left            =   1920
      TabIndex        =   52
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Pres. PagGuar. Filtrado"
      Height          =   350
      Index           =   8
      Left            =   4680
      TabIndex        =   51
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Horarios Programados"
      Height          =   350
      Index           =   7
      Left            =   4680
      TabIndex        =   50
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Cumplir PROG. XLS"
      Height          =   350
      Index           =   6
      Left            =   4680
      TabIndex        =   49
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "Varios"
      Height          =   255
      Index           =   8
      Left            =   9360
      TabIndex        =   48
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Reporte multiple"
      Height          =   350
      Index           =   5
      Left            =   4680
      TabIndex        =   47
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CheckBox Chk4 
      Caption         =   "Dscto. - Dif Pago Guardia"
      Height          =   350
      Index           =   6
      Left            =   6960
      TabIndex        =   46
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "Actualizar CADENA"
      Height          =   255
      Index           =   7
      Left            =   9360
      TabIndex        =   45
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Cronograma"
      Height          =   350
      Index           =   11
      Left            =   120
      TabIndex        =   44
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "Directorio XML"
      Height          =   255
      Index           =   6
      Left            =   9360
      TabIndex        =   43
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "Actualizar Tarjeta"
      Height          =   255
      Index           =   5
      Left            =   9360
      TabIndex        =   42
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "Actualiza Rem. Total"
      Height          =   255
      Index           =   0
      Left            =   9360
      TabIndex        =   41
      Top             =   600
      Width           =   1815
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Prog Directorio - Excepciones"
      Height          =   350
      Index           =   7
      Left            =   1920
      TabIndex        =   40
      Top             =   3000
      Width           =   2415
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "FERIADOS"
      Height          =   350
      Index           =   8
      Left            =   120
      TabIndex        =   39
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox Chk4 
      Caption         =   "Abono - Dif Pago Guardia"
      Height          =   350
      Index           =   5
      Left            =   6960
      TabIndex        =   38
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CheckBox Chk4 
      Caption         =   "Pago Jefatura de Guardia"
      Height          =   350
      Index           =   4
      Left            =   6960
      TabIndex        =   37
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CheckBox Chk4 
      Caption         =   "Guardias Programadas"
      Height          =   350
      Index           =   3
      Left            =   6960
      TabIndex        =   36
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Nivel Remunerado"
      Height          =   350
      Index           =   10
      Left            =   120
      TabIndex        =   35
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Editor de marcas"
      Height          =   350
      Index           =   6
      Left            =   1920
      TabIndex        =   34
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Cargar marcas de reloj"
      Height          =   350
      Index           =   5
      Left            =   1920
      TabIndex        =   33
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "File Configurador"
      Height          =   255
      Index           =   4
      Left            =   9360
      TabIndex        =   32
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "Acerca de..."
      Height          =   255
      Index           =   3
      Left            =   9360
      TabIndex        =   31
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "Usuarios"
      Height          =   255
      Index           =   2
      Left            =   9360
      TabIndex        =   30
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox Chk5 
      Caption         =   "Tablas nvo. Periodo"
      Height          =   255
      Index           =   1
      Left            =   9360
      TabIndex        =   29
      Top             =   960
      Width           =   1815
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Jefaturas del período"
      Height          =   350
      Index           =   4
      Left            =   4680
      TabIndex        =   27
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Tardanzas/Inasistencia"
      Height          =   350
      Index           =   3
      Left            =   4680
      TabIndex        =   26
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Masivo-record servidor(es)"
      Height          =   350
      Index           =   2
      Left            =   4680
      TabIndex        =   25
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Record de marcación"
      Height          =   350
      Index           =   1
      Left            =   4680
      TabIndex        =   24
      Top             =   840
      Width           =   2055
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Cambio de Turno"
      Height          =   350
      Index           =   4
      Left            =   1920
      TabIndex        =   23
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Autorización, Mot. Pers."
      Height          =   350
      Index           =   3
      Left            =   1920
      TabIndex        =   22
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Tablas "
      Height          =   350
      Index           =   9
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Ocurrencias / MP"
      Height          =   350
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Tarifas"
      Height          =   350
      Index           =   6
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Programa Asistenciales"
      Height          =   350
      Index           =   2
      Left            =   1920
      TabIndex        =   18
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Prog MENSUAL Administrativos"
      Height          =   350
      Index           =   1
      Left            =   1920
      TabIndex        =   17
      Top             =   840
      Width           =   2655
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Horarios"
      Height          =   350
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdsalir 
      Height          =   495
      Left            =   10800
      Picture         =   "UsuarioAcceso.frx":80266
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cerrar"
      Top             =   4680
      Width           =   615
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Grupo Trabajador"
      Height          =   350
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Cargos"
      Height          =   350
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Und. Prog. Rol "
      Height          =   350
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CheckBox Chk4 
      Caption         =   "Dscto. Mot. Personal"
      Height          =   350
      Index           =   2
      Left            =   6960
      TabIndex        =   11
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      TabIndex        =   10
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox Chk4 
      Caption         =   "Dscto. Tardanzas/ Inasist."
      Height          =   350
      Index           =   1
      Left            =   6960
      TabIndex        =   9
      Top             =   840
      Width           =   2175
   End
   Begin VB.CheckBox Chk4 
      Caption         =   "Presu. 100 Pgo Guardias "
      Height          =   350
      Index           =   0
      Left            =   6960
      TabIndex        =   8
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox Chk3 
      Caption         =   "Directorios"
      Height          =   350
      Index           =   0
      Left            =   4680
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.CheckBox Chk2 
      Caption         =   "Trabajadores"
      Height          =   350
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Organización"
      Height          =   350
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox Chk1 
      Caption         =   "Cadena"
      Height          =   350
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "UTILITARIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   28
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "RESOLUCIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6960
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "CONSULTAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "REGISTRO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "TABLAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "UsuarioAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Me.Caption = "USUARIO  : " & Usuarios.Txusuario.Text
   If Usuarios.Txacceso1.Text = "" Then
   Else
      For i = 1 To 12
          Chk1(i - 1).Value = IIf(Mid(Usuarios.Txacceso1.Text, i, 1) = "X", 1, 0)
      Next
   End If
   If Usuarios.Txacceso2.Text = "" Then
   Else
      For i = 1 To 9
         Chk2(i - 1).Value = IIf(Mid(Usuarios.Txacceso2.Text, i, 1) = "X", 1, 0)
      Next
   End If
   If Usuarios.Txacceso3.Text = "" Then
   Else
      For i = 1 To 9
          Chk3(i - 1).Value = IIf(Mid(Usuarios.Txacceso3.Text, i, 1) = "X", 1, 0)
      Next
   End If
   If Usuarios.Txacceso4.Text = "" Then
   Else
      For i = 1 To 7
          Chk4(i - 1).Value = IIf(Mid(Usuarios.Txacceso4.Text, i, 1) = "X", 1, 0)
      Next
   End If
   If Usuarios.Txacceso5.Text = "" Then
   Else
      For i = 1 To 9
          Chk5(i - 1).Value = IIf(Mid(Usuarios.Txacceso5.Text, i, 1) = "X", 1, 0)
      Next
   End If
   
End Sub

Private Sub cmdOK_Click()
   If IsNull(Usuarios.Txacceso1.Text) Then
   Else
     For i = 1 To 12
         Usuarios.Txacceso1.Text = IIf(Chk1(i - 1).Value = 1, Left(Usuarios.Txacceso1.Text, i - 1) + "X" + Right(Usuarios.Txacceso1.Text, 15 - i), Left(Usuarios.Txacceso1.Text, i - 1) + "0" + Right(Usuarios.Txacceso1.Text, 15 - i))
     Next
   End If
   If Usuarios.Txacceso2.Text = "" Then
   Else
      For i = 1 To 9
          Usuarios.Txacceso2.Text = IIf(Chk2(i - 1).Value = 1, Left(Usuarios.Txacceso2.Text, i - 1) + "X" + Right(Usuarios.Txacceso2.Text, 15 - i), Left(Usuarios.Txacceso2.Text, i - 1) + "0" + Right(Usuarios.Txacceso2.Text, 15 - i))
      Next
   End If
   If Usuarios.Txacceso3.Text = "" Then
   Else
      For i = 1 To 9
          Usuarios.Txacceso3.Text = IIf(Chk3(i - 1).Value = 1, Left(Usuarios.Txacceso3.Text, i - 1) + "X" + Right(Usuarios.Txacceso3.Text, 15 - i), Left(Usuarios.Txacceso3.Text, i - 1) + "0" + Right(Usuarios.Txacceso3.Text, 15 - i))
      Next
   End If
   If Usuarios.Txacceso4.Text = "" Then
   Else
      For i = 1 To 7
          Usuarios.Txacceso4.Text = IIf(Chk4(i - 1).Value = 1, Left(Usuarios.Txacceso4.Text, i - 1) + "X" + Right(Usuarios.Txacceso4.Text, 15 - i), Left(Usuarios.Txacceso4.Text, i - 1) + "0" + Right(Usuarios.Txacceso4.Text, 15 - i))
      Next
   End If
   If Usuarios.Txacceso5.Text = "" Then
   Else
      For i = 1 To 9
          Usuarios.Txacceso5.Text = IIf(Chk5(i - 1).Value = 1, Left(Usuarios.Txacceso5.Text, i - 1) + "X" + Right(Usuarios.Txacceso5.Text, 15 - i), Left(Usuarios.Txacceso5.Text, i - 1) + "0" + Right(Usuarios.Txacceso5.Text, 15 - i))
      Next
   End If
   Sw = True
   Unload Me
   Me.Hide
End Sub

Private Sub cmdsalir_Click()
  Me.Hide     '1er esconde
  Unload Me   '2do descarga
End Sub

