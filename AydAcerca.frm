VERSION 5.00
Begin VB.Form AydAcerca 
   Caption         =   "Acerca de..."
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form2"
   Picture         =   "AydAcerca.frx":0000
   ScaleHeight     =   3585
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   240
      MaxLength       =   1500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "AydAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim sql  As String
Private Sub Form_Load()
  Set rs = New Recordset
  rs.Open "SELECT * FROM CpAyudas WHERE grupo='" & c_grupo & "' And secuencia='" & c_secuencia & "'", Cnx, adOpenStatic, adLockReadOnly
  If rs.EOF Then
  Else
     Text1.Text = rs!descripcion
  End If
  rs.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set rs = Nothing
   Me.Hide
   Unload Me
End Sub

