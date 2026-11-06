VERSION 5.00
Begin VB.Form RegSerial 
   Caption         =   "Registro Serial"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3675
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "RegSerial.frx":0000
   ScaleHeight     =   1605
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtContra 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   7
      PasswordChar    =   "&"
      TabIndex        =   1
      Top             =   360
      Width           =   900
   End
   Begin VB.CommandButton Cmd 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Height          =   450
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CLAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "RegSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dr As Scripting.Drive                  'objeto drive
Dim FSO As New Scripting.FileSystemObject  'objeto Filesystemobject

Private Sub Cmd_Click()
    Unload Me        'Antes de salir ejecuta el evento QueryUnload
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Trim(TxtContra.Text) = "wcvg" Then
       For Each dr In FSO.Drives
         If UCase$(dr.DriveLetter) = "C" Then
            NumerodeSerie = CStr(dr.SerialNumber)
            Exit For
         End If
       Next
       Encclave = ""
       WVencripta (NumerodeSerie)
       Set rs = New Recordset
       sql = "select * from CPregserial where regserial='" & Encclave & "'"
       rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
       If rs.EOF Then
          rs.AddNew
          rs!RegSerial = Encclave
          rs!fecha = Date
          rs.Update
          MsgBox ("Serial registrado ")
       Else
          MsgBox ("Es Valido ")
       End If
       Encclave = ""
       rs.Close
       Call Cmd_Click
    Else
       For i = 1 To 3   ' Realiza un bucle 3 veces.
            Beep 500, 350      ' Emite un tono.
       Next i
       MsgBox "Contraseña equivocada.  " & vbCr _
            & "- Esta opción es de uso restringido.   ", vbCritical
    End If
End Sub
Private Sub TxtContra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Unload Me       'Antes de salir ejecuta el evento QueryUnload
    End If
End Sub


