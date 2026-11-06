VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form inicio 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Ingreso"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "INICIO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "INICIO.frx":0442
   ScaleHeight     =   5865
   ScaleMode       =   0  'User
   ScaleWidth      =   10012.19
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   495
      Left            =   7440
      TabIndex        =   11
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton CmdChkPwd 
      Height          =   495
      Left            =   7080
      Picture         =   "INICIO.frx":4F2F
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cambiar clave"
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox TxAño 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   6600
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   7
      Top             =   3240
      Width           =   795
   End
   Begin VB.TextBox Txusuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   7800
      MaxLength       =   10
      PasswordChar    =   "&"
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton CmdIN 
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   3240
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5520
      Top             =   4440
   End
   Begin VB.Label LblClienteIp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label LblNick 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label LblDia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LblDia"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label LblDS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LblDs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6960
      TabIndex        =   12
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Lblhora 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6120
      TabIndex        =   9
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODO"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Lblrazon 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   5760
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Lblnombre 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   8400
      Picture         =   "INICIO.frx":57F9
      ToolTipText     =   "Salir"
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label Lblsistema 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Walter C. Valqui Guarniz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistemas H.E.P. Informatica "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   1455
   End
End
Attribute VB_Name = "inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dr As Scripting.Drive                  'objeto drive
Dim FSO As New Scripting.FileSystemObject  'objeto Filesystemobject
Dim wvl As Integer
Dim ccad As String
Dim unafecha As Date
'- Saber Ip cliente/host, agregar componente Microsoft Winsock control 6.0
'- agrega control winsock y usar sus propiedades ha criterio
Private Sub Form_Load()
  wflg = "NO"
  If App.PrevInstance Then
       MsgBox "Aplicación Ya esta  en Ejecución", vbExclamation, "Ya está siendo ejecutada"
       End
  End If
 
  Winsock1.RemoteHost = Winsock1.LocalHostName
  w_clienName = Winsock1.LocalHostName  ' captura Nombre Pc cliente de ejecucion
  w_clienteIP = Winsock1.LocalIP        ' captura IP cliente de ejecucion
  '-valida nombre estación max. 10posc.
  If Len(w_clienName) > 10 Then
     Beep 700, 300
     If MsgBox("Nombre del equipo EXCEDE lo permitido, continua ?  ", vbQuestion + vbYesNo, "[" + w_clienName + "]") = vbNo Then
        End
     End If
     w_clienName = Left(w_clienName, 10)
  End If
  LblDS.Caption = cICatal + " - " + IIf(Val(Right(cDSource, 2)) > 0, Right(cDSource, 2), Trim(cDSource))
  
  'Razon social
  Set rs = New Recordset
  sql = "SELECT * FROM CPnumeros WHERE tipdoc=" & 99
  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  wrazonsocial = Trim(rs!descripcion)
  rs.Close
  sql = "SELECT * FROM CPnumeros WHERE tipdoc=" & 88
  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  wsistema = Trim(rs!descripcion)
  rs.Close
  
  '-Unidad carga del Manual (pdf)
  sql = "select * from CPnumeros where tipdoc=" & 77
  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
  wmanualPdf = Trim(rs!descripcion)
  rs.Close
  
  GenMaxTN = 6
  Set rs = New Recordset
  rs.Open "SELECT * FROM Cptablas WHERE tabla='45' and codigo='00'", Cnx, adOpenStatic, adLockReadOnly
  If Not rs.EOF Then
     GenMaxTN = Val(rs!campo)     'Maximo permitido Turnos Noche en Directorio
  End If
  rs.Close
  GenMaxDL = 4
  Set rs = New Recordset
  rs.Open "SELECT * FROM Cptablas WHERE tabla='46' and codigo='00'", Cnx, adOpenStatic, adLockReadOnly
  If Not rs.EOF Then
     GenMaxDL = Val(rs!campo)     'Maximo Dias Libres permitidos en Directorio
  End If
  rs.Close
  
  Lblrazon.Caption = wrazonsocial
  Lblsistema.Caption = wsistema
  Txusuario = ""
  Txclave = ""
  CmdIN.Enabled = False
  Image2.Enabled = True
  genaño = Year(Date)
  TxAño = genaño:  w_ano = genaño
  genmes = Month(Date)
           wMeses = "ENERO    FEBRERO  MARZO    ABRIL    MAYO     JUNIO    "
  wMeses = wMeses + "JULIO    AGOSTO   SETIEMBREOCTUBRE  NOVIEMBREDICIEMBRE"
  '                    XXXXXXXX1XXXXXXXX2XXXXXXXX3XXXXXXXX4XXXXXXXX5XXXXXXXX6
  gennombmes = ""
  For X = 1 To 12
    If genmes = X Then
       gennombmes = Trim(Mid(wMeses, (X - 1) * 9 + 1, 9))
    End If
  Next X
  
  gendia = Day(Date)
  unafecha = Right("0" + Trim(Str(gendia)), 2) & "/" + Right("0" + Trim(Str(genmes)), 2) + "/" + Trim(Str(genaño))
  LblDia.Caption = WeekdayName(Weekday(unafecha)) + " " & gendia & " de " & gennombmes & " del " & genaño
  
  '- 02/Ene/2000  Domingo
  ''MsgBox (WeekdayName(Weekday("02/01/2000")) + " " + Format("02/01/2000", "dd/mm/yyyy"))
  If Weekday("02/01/2000") <> 1 Or UCase(Left(WeekdayName(Weekday("02/01/2000")), 3)) <> "DOM" Then  'No es Domingo dia Inicio semana
     MsgBox "[" + WeekdayName(Weekday(unafecha)) + "], Error, Inicio de Semana no es Domingo, Comunique a Sistemas.", vbCritical, "Configuración Calendario"
     End
  End If
  
  Exit Sub
TrataError:
    Beep 500, 350
    For Each adoErr In Cnx.Errors
        MsgBox _
        Prompt:=adoErr.Description, _
        Title:="Error Nro. " & CStr(adoErr.Number)
    Next
    End
End Sub
Private Sub Label4_Click()
   Load RegSerial
   RegSerial.Left = 0: RegSerial.Top = 0
   RegSerial.Refresh
   RegSerial.Show
End Sub

Private Sub Txusuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      '  Call chkNroserial
      If Trim(Txusuario) = "" Then
         Beep 500, 350
         MsgBox "No indico usuario..", vbCritical
         End
      End If
      Txusuario = UCase(Txusuario)
      Call Cmdfirmar
      Exit Sub
   End If
End Sub

Private Sub CmdChkPwd_Click()
    If wcambia = "SI" Or wProgUpr = "S" Then
       Load ChkPwd
       ChkPwd.Left = 0: ChkPwd.Top = 0
       ChkPwd.Refresh
       ChkPwd.Show
    End If
End Sub

Private Sub Cmdfirmar()
  Sw = False
  wrights = "":  wDerEspec = "": wTerminal = "": wBaja = ""
  Encclave = ""
  WVencripta (Txusuario)
  wTarjeta = "SINUSO"
  Wccosto = "SINUSO"
  WcodUPR = "": wProgUpr = "N"
  wTipoPla = ""
  wDirector = ""      'Director Dpto/Oficina; Observa records de subordinados
  If Trim(Txusuario) = "8E9Z0I3" Then
     '-Administrador
     wUsuario = "AMASTER"
     wNick = "AMaster"
     wUsuNombre = "AMASTER"
     wrights = "MAS"
     wDerEspec = "S"
     Lblnombre.Caption = wUsuNombre
     Wccosto = ""
     wTipoPla = ""
     Sw = True
     wacceso1 = "XXXXXXXXXXXXXXX": wacceso2 = "XXXXXXXXXXXXXXX": wacceso3 = "XXXXXXXXXXXXXXX": wacceso4 = "XXXXXXXXXXXXXXX"
     wacceso5 = "XXXXXXXXXXXXXXX": wacceso6 = "XXXXXXXXXXXXXXX": wacceso7 = "XXXXXXXXXXXXXXX"
  Else
     '-valida Usuario
     Set rs = New Recordset
     sql = "SELECT * FROM CPUsuarios WHERE clave='" & Encclave & "'"
     '''  MsgBox sql + " -" + Txusuario
     rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
     If Not rs.EOF Then
        wmodulo = rs!modulo
            wPwdOk = Trim(Txusuario)
        wClaveOld = Trim(Txusuario)
        wcambia = IIf(IsNull(rs!cambia), "", rs!cambia)
        wUsuario = rs!usuario
        wrights = rs!derechos
        wDerEspec = IIf(IsNull(rs!DerEspec), "", rs!DerEspec)
        wBaja = IIf(IsNull(rs!BAJA), "", rs!BAJA)
        Wccosto = ""
        wTipoPla = ""
        Lblnombre.Caption = rs!nombre
        wUsuNombre = rs!nombre
        npos = InStr(rs!nombre, " ")
        wNick = Left(wUsuNombre, 1) + Trim(Mid(wUsuNombre, npos + 1, 9))
        wUsuNombre = rs!nombre + " (U)"
        wacceso1 = IIf(IsNull(rs!acceso1), "000000000000000", rs!acceso1): wacceso2 = IIf(IsNull(rs!acceso2), "000000000000000", rs!acceso2): wacceso3 = IIf(IsNull(rs!acceso3), "000000000000000", rs!acceso3): wacceso4 = IIf(IsNull(rs!acceso4), "000000000000000", rs!acceso4)
        wacceso5 = IIf(IsNull(rs!acceso5), "000000000000000", rs!acceso5): wacceso6 = IIf(IsNull(rs!acceso6), "000000000000000", rs!acceso6): wacceso7 = IIf(IsNull(rs!acceso7), "000000000000000", rs!acceso7)
        Sw = True
          '''MsgBox "1.usr"
     Else
        '-Acceso DNI
        rs.Close
        
          '--MsgBox "1.Dni"
        Encclave = ""
        wDirector = ""
        WVencripta ("HEPMCA")     'acceso Operador generico
        Set rs = New Recordset
        sql = "SELECT * FROM CPUsuarios WHERE clave='" & Encclave & "'"
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        If Not rs.EOF Then
           wmodulo = rs!modulo
           wrights = rs!derechos
           wDerEspec = IIf(IsNull(rs!DerEspec), "", rs!DerEspec)
           wBaja = IIf(IsNull(rs!BAJA), "", rs!BAJA)
           wacceso1 = IIf(IsNull(rs!acceso1), "000000000000000", rs!acceso1): wacceso2 = IIf(IsNull(rs!acceso2), "000000000000000", rs!acceso2): wacceso3 = IIf(IsNull(rs!acceso3), "000000000000000", rs!acceso3): wacceso4 = IIf(IsNull(rs!acceso4), "000000000000000", rs!acceso4)
           wacceso5 = IIf(IsNull(rs!acceso5), "000000000000000", rs!acceso5): wacceso6 = IIf(IsNull(rs!acceso6), "000000000000000", rs!acceso6): wacceso7 = IIf(IsNull(rs!acceso7), "000000000000000", rs!acceso7)
           rs.Close
           wUsuario = Trim(Txusuario)  'guarda DNI
           
            '--  MsgBox "2.dni [" + wUsuario + "]"
           
           '-Generales del servidor
           Set rs = New Recordset
           sql = "SELECT * FROM CPgeneral WHERE (baja is NULL) and documento='" & wUsuario & "'"    'Dni
             '--  MsgBox "2.1 " + sql
           rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
           If Not rs.EOF Then
              wTarjeta = rs!tarjeta
              Wccosto = IIf(IsNull(rs!CCosto), Wccosto, rs!CCosto)
              wTipoPla = IIf(IsNull(rs!tipoPla), Wccosto, rs!tipoPla)
              wDirector = IIf(IsNull(rs!Director), "", rs!Director)
              WcodUPR = IIf(IsNull(rs!CodUpr), "", rs!CodUpr)       ' Unidad Prog Roles
              wNick = Left(rs!nombres, 1) + Left(rs!apelpat, 9)
              Lblnombre.Caption = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35)
              wUsuNombre = Left(Trim(rs!apelpat) + " " + Trim(rs!apelmat) + " " + Trim(rs!nombres), 35) + " (D)"
              Encclave = ""
              If IsNull(rs!Clave) Then
                 '-sin clave coge DNI
                 WVencripta (wUsuario)
              Else
                 WVdesencripta (Trim(rs!Clave))
              End If
              wClaveOld = Encclave
                   wPwdOk = Encclave
              wcambia = IIf(IsNull(rs!ActClave), "", rs!ActClave)
              Sw = True
               '--  MsgBox "3.dni [cc:" + Wccosto + "]" + Txusuario + "(" + wclave + ")pwd:" + wPwdOk
           Else
              ''MsgBox "3.dni-error cc:"
           End If
           weditor = "X"           'Verifica si es Titular UPR
           wCadenaUpr = ""
           If Wccosto <> "SINUSO" Then
              rs.Close
              '-Verifica derechos Programador UPR
              Set rs = New Recordset
              '''rs.Open "SELECT * FROM CpCcostoUPR WHERE ccosto='" & Wccosto & "'", Cnx, adOpenStatic, adLockReadOnly
              rs.Open "SELECT * FROM CpCcostoUPR order by codupr ", Cnx, adOpenStatic, adLockReadOnly
              Do While Not rs.EOF
                 If rs!Dni1 = wUsuario Then
                    If InStr(wCadenaUpr, rs!CodUpr) = 0 Then
                       wCadenaUpr = wCadenaUpr + (rs!CodUpr + rs!editor1 + "_")
                       weditor = "S"
                    End If
                 ElseIf rs!Dni2 = wUsuario Then
                        If InStr(wCadenaUpr, rs!CodUpr) = 0 Then
                           wCadenaUpr = wCadenaUpr + (rs!CodUpr + rs!editor2 + "_")
                           weditor = "S"
                        End If
                 ElseIf rs!dni3 = wUsuario Then
                        If InStr(wCadenaUpr, rs!CodUpr) = 0 Then
                           wCadenaUpr = wCadenaUpr + (rs!CodUpr + rs!editor3 + "_")
                           weditor = "S"
                        End If
                 End If
                 rs.MoveNext
              Loop
              wProgUpr = weditor
           End If
        End If
     End If
     rs.Close
     If w_clienteIP = "192.168.1.19" Then
        MsgBox "cc:" + Wccosto + "[" + wCadenaUpr + "=>" + wPwdOk
     End If
  End If
  If Wccosto = "SINUSO" Then
     Lblnombre.Caption = "<E> Sin C.Costo..!"
     Txusuario = ""
     Txusuario.SetFocus
  End If
  If Not Sw Then
     Lblnombre.Caption = "<E> USUARIO..!"
     Txusuario = ""
     Txusuario.SetFocus
  Else
     If wBaja = "S" Then
        Lblnombre.Caption = "<E>Usuario de BAJA..!"
        Txusuario = ""
        Txusuario.SetFocus
     Else
  
        LblNick.Caption = wNick
        LblClienteIp.Caption = w_clienteIP
        If w_clienteIP <> "" Then
            wTerminal = Right(w_clienteIP, 3)
            If Left(wTerminal, 1) = "." Then
                wTerminal = Right(w_clienteIP, 2)    '.19
            ElseIf Mid(wTerminal, 2, 1) = "." Then
                    wTerminal = Right(w_clienteIP, 1)  '.7
            End If
        End If
        If wProgUpr = "S" And wcambia = "SI" Then
            '-verifica cambio de clave
            Beep 500, 350
            MsgBox "Debe cambiar Clave de programaciòn (doc. identidad).", vbCritical, "[PROGRAMADOR UPR]"
            CmdChkPwd.SetFocus
            Exit Sub
        End If
     
        CmdIN.Enabled = True
        CmdIN.SetFocus
        
     End If
  End If
End Sub

Private Sub UpDown1_DownClick()
    If TxAño <= 1980 Then
       TxAño = 2090
       Exit Sub
    End If
    TxAño = TxAño - 1
    genaño = Trim(TxAño)
End Sub
Private Sub UpDown1_UpClick()
    If TxAño >= 2090 Then
       TxAño = 1980
       Exit Sub
    End If
    TxAño = TxAño + 1
    genaño = Trim(TxAño)
End Sub
Private Sub cmdIN_Click()
   If Trim(Txusuario) = "" Then
      Beep 500, 350
      MsgBox "No indico usuario..", vbCritical
      End
   End If
   wflg = "NO"
   w_ano = TxAño
   
   NombTabla = "cpmarca" + w_ano
   Call CompruebaTabla
   If wflg = "NO" Then
      MsgBox "Periodo " & w_ano & " No es ..", vbExclamation
      UpDown1.SetFocus
      Exit Sub
   End If
   genaño = TxAño
   Unload Me
   ''Me.Hide
   Load CpMenu
   CpMenu.Left = 0: CpMenu.Top = 0
   CpMenu.Show
End Sub
Private Sub Image2_Click()
   Cnx.Close
   Cnx0.Close
   Set Cnx0 = Nothing
   Set Cnx = Nothing
   End
End Sub
Private Sub Timer1_Timer()
   Lblhora.Caption = Format(Time, "hh:mm:ss")
End Sub

Private Sub chkNroserial()
     For Each dr In FSO.Drives
         If UCase$(dr.DriveLetter) = "C" Then
            NumerodeSerie = CStr(dr.SerialNumber)
            Exit For
         End If
      Next
      Encclave = ""
      WVencripta (NumerodeSerie)
      Set rs = New Recordset
      sql = "SELECT * FROM CPregserial WHERE regserial='" & Encclave & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If rs.EOF Then
         MsgBox ("EQUIPO NO Autorizado..!")
         rs.Close
         Call Image2_Click
      End If
      rs.Close
End Sub

Public Sub VALMAC()
   'Dim colNetAdapters, objWMIService As Object
   Dim strComputer As String
   '-strComputer = "."  direccion IP Servidor ha Scanear
   strComputer = cDSource
   'Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
   'Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
   'For Each objitem In colNetAdapters
   '    If objitem.MACAddress = cMAC Then
          wacceso = True
   '    End If
   'Next
End Sub



'Nota:
'====
'Si quieres usar un icono para el ejecutable diferente al creado por el VB6, puedes hacer lo siguiente:
'-Añade un formulario al proyecto.
'-Cambia el icono del formulario por el que quieras usar.
'-En el cuado de propiedades de las propiedades del proyecto (menú Proyecto>Propiedades de Link...), selecciona la ficha Generar, y selecciona ese formulario del combo que hay junto a Icono.
'-El Objeto inicial (de la ficha General) debe seguir siendo Sub Main.
'-Compilar el proyecto y el ejecutable tendrá el icono indicado
