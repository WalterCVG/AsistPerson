VERSION 5.00
Begin VB.Form ChkPwd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1950
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ChkPwd.frx":0000
   ScaleHeight     =   1152.124
   ScaleMode       =   0  'User
   ScaleWidth      =   3323.868
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxPwdC 
      ForeColor       =   &H80000006&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox TxPwd 
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   10
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox txPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   1
      Top             =   240
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&NUEVA Contraseña:"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Contraseña Actual:"
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "ChkPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean
Dim rs As New ADODB.Recordset
Dim wclave, sql, FlgUser As String
Private Sub Form_Load()
  TxPwd = "": TxPwdC = ""
  TxPwd.Enabled = False
  TxPwdC.Enabled = False
  cmdOK.Enabled = False
End Sub
Private Sub TxPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Encclave = ""
    WVencripta (UCase(txPassword))
    FlgUser = "N"
    wcambia = "SI"
    Set rs = New Recordset
    sql = "select * from CPUsuarios where clave='" & Encclave & "'"
    rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.EOF Then
       FlgUser = "S"
       '--wPwdOk = Trim(txPassword)
       wUsuario = rs!usuario
       wcambia = rs!cambia
       ChkPwd.Caption = "Login : " + wUsuario
       rs.Close
       If wcambia <> "SI" Then
          MsgBox ("Usuario NO Autorizado..")
          Call cmdCancel_Click
       End If
       TxPwd.Enabled = True
       TxPwdC.Enabled = True
       TxPwd.SetFocus
    Else
       rs.Close
       sql = "select * from CPGeneral where (baja is NULL) and Documento='" & Trim(wUsuario) & "'"
       rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
       If Not rs.EOF Then
          Encclave = ""
          WVdesencripta (Trim(rs!Clave))
          wPwdOk = Encclave
          wClaveOld = Encclave
          rs.Close
          TxPwd.Enabled = True
          TxPwdC.Enabled = True
          TxPwd.SetFocus
       Else
          MsgBox ("Usuario NO Existe..")
          rs.Close
          Call cmdCancel_Click
       End If
    End If
  End If
  Exit Sub
End Sub

Private Sub TxPwd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxPwd = UCase(TxPwd)
      Encclave = ""
      WVencripta (Trim(TxPwd))
      If FlgUser = "S" Then
         sql = "select * from CPUsuarios where clave='" & Encclave & "'"
         rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
         If Not rs.EOF Then
            If wUsuario <> rs!usuario Then
               MsgBox ("Clave no disponible..!")
               TxPwd = ""
               Exit Sub
            End If
         End If
         rs.Close
      Else
         If TxPwd = wUsuario Then
            MsgBox "NO se permite Doc. Identidad..!", vbExclamation
            TxPwd = ""
            Exit Sub
         End If
      End If
      TxPwdC.SetFocus
      Exit Sub
   End If
End Sub
Private Sub TxPwdC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TxPwdC = UCase(TxPwdC)
      If TxPwd = TxPwdC Then
         cmdOK.Enabled = True
         cmdOK.SetFocus
         Exit Sub
      Else
         MsgBox "Contraseña no coincide. Vuelva a intentar"
         TxPwdC.Text = ""
      End If
   End If
End Sub
Private Sub cmdOK_Click()
    'comprobar si la contraseña es correcta
    If TxPwd = TxPwdC And TxPwd <> "" Then
        'colocar código aquí para pasar al sub
        'que llama si la contraseña es correcta
        'lo más fácil es establecer una variable global
        Encclave = ""
        WVencripta (Trim(TxPwd))
        If FlgUser = "S" Then
           sql = "UPDATE CPusuarios SET clave='" + Encclave + "' WHERE usuario='" + wUsuario + "'"
        Else
           sql = "UPDATE CPGeneral SET clave='" + Encclave + "', ActClave='NO'  WHERE  (baja is NULL) and Documento='" + wUsuario + "'"
        End If
        Cnx.Execute sql
        
        Set rs = New Recordset
        sql = "SELECT * FROM CpAudClaves WHERE documento='" & Trim(wUsuario) & "'"
        rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
        rs.AddNew
        '-rs!tarjeta = Txtarjeta
        rs!Documento = wUsuario
        rs!claveold = wClaveOld
        rs!claveNew = Trim(TxPwd)
        rs!cCosto = Wccosto
        rs!CodUpr = WcodUPR
        rs!fecha = Date
        rs!hora = Format(Time, "HH:mm:ss")
        rs!usuario = wNick         'Firma Usuario
        rs!estacion = w_clienName
        rs.Update
        rs.Close
        
        wPwdOk = Trim(TxPwd)
        LoginSucceeded = True
        Me.Hide
        End
    Else
        MsgBox "Contraseña no coincide. Vuelva a intentarlo"
        TxPwd = "":        TxPwdC = ""
        TxPwd.SetFocus
        Exit Sub
     End If
End Sub
Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    Me.Hide
End Sub

