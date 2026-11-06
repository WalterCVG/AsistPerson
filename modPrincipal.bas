Attribute VB_Name = "modPrincipal"
Option Explicit
'Indispensable para browse del combo
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Long) As Long

'-Sonido Beep   frecuencia, Duracion
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
'-Ejecutar comando
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Cnx As ADODB.Connection
Public Cnx0 As ADODB.Connection
Public Cnx1 As ADODB.Connection
Public cDSource, cMAC, cProvider As String
Public cICatal0, cICatal, cUsr, cPwd As String
Public CString, CString0 As String
Public CString1, PathLeg As String
Public genaño, genmes, gendia, GenMaxTN, GenMaxDL As Integer   'max. turno N / Max. Dia Libre
Public w_clienteIP, w_clienName, wTerminal, wrights, wBaja, wDerEspec, wClaveOld, wcambia, wProgUpr, weditor, wLstHC As String
Public NombTabla, wrazonsocial, wsistema, wmodulo, wUsuario, wPwdOk, wUsuNombre, wNick, wTarjeta, Wccosto, wTipoPla, wDirector, WcodUPR, wTituRpt, wCadenaUpr As String
Public nTipRpt, wvlen, wpages As Integer
Public Encclave, numeroserie As String
Public wMeses, wnombrep, W_ccosto, W_upr, W_RegPlh, W_lbldh, wMesdeAno, wNroResolu, wmsgpie, wNroReso0, WfechaR0, wPer_resol, Wonomastico As String
Public wForma, wHorasP, wmTarda, wmDscto, wHorasFJ, wHorasMP, wHorasFI, wHorasT As String
Public wmanualPdf, c_grupo, c_secuencia As String
Public wflg, wTpoQuery, wNameXLS, wNombFile, wcampo1, wcampo2, c_Periodo, cObjeto As String
Public nnanos, nnmeses, nndias, nMinTot, nLapso As Integer
Public w_ano, MSG, FecFIN, FecINICIO, cHorsMints As String
Public wacceso As Boolean
Public wSigueOk As Boolean
Public wacceso1, wacceso2, wacceso3, wacceso4, wacceso5, wacceso6, wacceso7 As String
Public carpeta, cPlantilla0, cDirectHoy As String
Public unafecha, primer, ultimo, pridia, ultdia As Date

'Variable para acceder al Registro mediante Wsh - Windows Scripting Host (referencia)
Dim O_Registro As WshShell
Private Const LocalMachine As String = "HKEY_LOCAL_MACHINE"
Private Const CurrentUser As String = "HKEY_CURRENT_USER"

Public Sub Main()
   Set O_Registro = New WshShell
   If Leer_Dato(CurrentUser, "sCurrency") = "S/." Or Leer_Dato(CurrentUser, "sCurrency") = "S/" Then
   Else
      MsgBox "Verificar, Configuración Regional " + Leer_Dato(CurrentUser, "sCurrency"), vbInformation
      Set O_Registro = Nothing
      End
   End If
   Set O_Registro = Nothing

   cDSource = "C:\rpt\cnxheper.ini"
   If Dir(cDSource) = "" Then
     MsgBox " No existe Configurador..!"
     End
   End If
   
   If InStr(UCase(App.path), "ASISTPERSON") = 0 And InStr(UCase(App.path), "CNTRLPER") = 0 Then
      Beep 700, 300
      MsgBox "Error de acceso, Ruta de ejecución incorrecta..!", vbCritical, "DIRECTORIO ORIGEN"
      End
   End If
   
    
   '-Lee file .INI y  validar parametros para conexion
   wacceso = False
   LeerINI
   '    MsgBox "1"
   
   Call inicio.VALMAC   'valida MAC
   If Not wacceso Then
      MsgBox "Error Critico: Conexion impropia", vbExclamation
      End
   End If
         
   Dim adoErr As ADODB.error
   Set Cnx = New ADODB.Connection
   Set Cnx0 = New ADODB.Connection
   Set Cnx1 = New ADODB.Connection    '-Legajos    MDB
   On Error GoTo TrataError
   
   '-CString0 = "Provider=SQLOLEDB;Data Source='192.168.1.07';" & _
   '-   "initial catalog=HepMaestSql;user id='sa';password='pwd';"
   '-----DATA Source=hepdatos   esta en DOMINIO
   CString0 = "Provider=" + cProvider + ";Data Source='" + cDSource + "';" & _
       "initial catalog=" + cICatal0 + "; user id='" + cUsr + "'; password='" + cPwd + "';"
   Cnx0.ConnectionTimeout = 30
   Cnx0.CommandTimeout = 30
   With Cnx0
     .ConnectionString = CString0
     .Open
   End With
   'MsgBox "2"
      
   '-CString = "Provider=SqlOleDb;Data Source=192.168.1.07,1433;network library=DBMSSOCN;" & _
   '-   "initial catalog=hepControlP; user id='sa'; password='pwd';"
   CString = "Provider=" + cProvider + ";Data Source='" + cDSource + "';" & _
      "initial catalog=" + cICatal + "; user id='" + cUsr + "'; password='" + cPwd + "';"
   Cnx.ConnectionTimeout = 30
   Cnx.CommandTimeout = 30
   With Cnx
     .ConnectionString = CString
     .Open
   End With
   'MsgBox "3"
      
     
   '-----DATA Legajos MDB
   If InStr(UCase(App.path), "ASISTPERSON") > 0 Then
      PathLeg = Replace(UCase(App.path), "ASISTPERSON", "LEGAJOS")
   ElseIf InStr(UCase(App.path), "CNTRLPER") > 0 Then
           PathLeg = Replace(UCase(App.path), "CNTRLPER", "LEGAJOS")
   End If
     'MsgBox "3.5 " + PathLeg
   
   CString1 = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" & PathLeg & "\Personal.Mdb" & ";" & _
       "Jet OLEDB:Database Password='hep'"
    'MsgBox "4.0 " + CString1
       
   With Cnx1
      .ConnectionString = CString1
      .Open
   End With
   'MsgBox "4.."
   
   inicio.Show
   Exit Sub
TrataError:
Beep 500, 350
   For Each adoErr In Cnx.Errors
        MsgBox _
            Prompt:=adoErr.Description, _
            Title:="Error Nro. " & CStr(adoErr.Number)
   Next
End Sub

Public Function Leer_Dato(Principal As String, Valor As String) As String
    Leer_Dato = O_Registro.RegRead(Principal & "\Control Panel\International\" & Valor)
End Function

Public Sub precentworm(picture As Control, ByVal percent)
    Dim num$
    If Not picture.AutoRedraw Then
       picture.AutoRedraw = -1
    End If
    picture.Cls
    picture.ScaleWidth = 100
    picture.DrawMode = 10
    num$ = Format$(percent, "###") + "%"
    picture.CurrentX = 50 - picture.TextWidth(num$) / 2
    picture.CurrentY = (picture.ScaleHeight - picture.TextHeight(num$)) / 2
    picture.Print num$
    picture.Line (0, 0)-(percent, picture.ScaleHeight), , BF
    picture.Refresh
End Sub

'-Export using Microsoft Excel Object Library, This code copies data from temp table to an excel file(temp.xls),
'-delete the generated xml file at the end of the program as its just an temp file.
'-Add a project reference to ADODB & Excel Object library before using this code,
'-crear tabla temporal (nombre varchar (10), int edad) create table temp ( name varchar(10), age int)
'-insertar en los valores temporales ('XYZ', 20) insert into temp values ( 'xyz', 20)
'-insertar en los valores temporales ('ABC', 30) insert into temp values ( 'abc', 30)
Public Sub ExportTableToXLS_click()    '    wTpoQuery As String)
 On Error GoTo Control
 Dim SqlXls As String
 Dim strExcelFile, strXmlFile As String
 'Dim cnx As ADODB.Connection
 Dim rs As ADODB.Recordset
 Dim eApp As Excel.Application
 Dim wk As New Excel.Workbook
 'Set cnx = New ADODB.Connection
 Set rs = New ADODB.Recordset
 
 CreaDir ("c:\rpt")
 '-Si ficheros excel y xm; existen, aqui se borran
 Select Case wTpoQuery
    Case "0"
       strExcelFile = "C:\RPT\RESDSCTS.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\RESDSCTS.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM CpTmResDscto"
       wNombFile = "RESDSCTS"
    Case "1"
       strExcelFile = "C:\RPT\RESMP.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\RESMP.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM CpTmResMP"
       wNombFile = "RESMP"
    Case "1.1"
       strExcelFile = "C:\RPT\QMULTIPLE.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\QMULTIPLE.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM CpTmgeneral"
       wNombFile = "QMULTIPLE"
       
    Case "1.5"
       strExcelFile = "C:\RPT\PGFiltrado.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\PGFiltrado.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM #TabPGFilt"
       wNombFile = "PGFiltrado"
       
    Case "2"
       strExcelFile = "C:\RPT\RESJGUAR.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\RESJGUAR.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM CpTmResJGuar"
       wNombFile = "RESJGUAR"
    Case "3"
       strExcelFile = "C:\RPT\1047" + ResPgoGuarProy.Txmes + Right(Trim(Str(genaño)), 2) + ".xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\1047" + ResPgoGuarProy.Txmes + Right(Trim(Str(genaño)), 2) + ".xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM CpPG1047"
       wNombFile = "1047" + ResPgoGuarProy.Txmes + Right(Trim(Str(genaño)), 2)
    Case "4"
       strExcelFile = "C:\RPT\RESPGABO.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\RESPGABO.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM CpTmDifPgoGuar"
       wNombFile = "RESPGABO"
    Case "5"
       strExcelFile = "C:\RPT\RESPGDSC.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\RESPGDSC.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM CpTmDifPgoGuar"
       wNombFile = "RESPGDSC"
    Case "5.7"
       strExcelFile = "C:\RPT\CpDirXML.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\cPDirXML.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM CpTmDirXML"
       wNombFile = "CpDirXML"
       
    Case "7.0"
       strExcelFile = "C:\RPT\SUMRecord.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\SumRecord.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM #Tabla5"
       wNombFile = "#Tabla5"
    Case "7.7"
       strExcelFile = "C:\RPT\DirPrograma.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\DirPrograma.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM #Tabla77"
       wNombFile = "#Tabla77"
    Case "7.8"
       strExcelFile = "C:\RPT\DirProgALL.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\DirProgALL.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM #Tabla78"
       wNombFile = "#Tabla78"
      
    Case "9"
       strExcelFile = "C:\RPT\CpCumplimiento.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\CpCumplimiento.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM #tabla9"
       wNombFile = "CpCumplimiento"
       
    Case "10"
       strExcelFile = "C:\RPT\SUMAutori.xls"
       If Dir(strExcelFile) <> "" Then Kill strExcelFile
       strXmlFile = "C:\RPT\SUMAutori.xml"
       If Dir(strXmlFile) <> "" Then Kill strXmlFile
       SqlXls = "SELECT * FROM #tablaMI"
       wNombFile = "SUMAutori"
       
 End Select
 
 '-cnx.Open "Provider=SQLOLEDB.1;Password=password;Persist Security Info=True;User ID=sa;Initial Catalog=test;Data Source=TDL06B21"
 'rs.Open "SELECT * FROM tablatmp", Cnx, adOpenStatic
 rs.Open SqlXls, Cnx, adOpenStatic
 rs.Save strXmlFile, adPersistXML
 Set eApp = New Excel.Application
 Set wk = eApp.Workbooks.Add
 With wk.ActiveSheet.QueryTables.Add(Connection:="FINDER;" & strXmlFile, Destination:=wk.ActiveSheet.Range("A1"))
   .Name = wNombFile
   .FieldNames = True
   .RowNumbers = False
   .FillAdjacentFormulas = False
   .PreserveFormatting = True
   .RefreshOnFileOpen = False
   .BackgroundQuery = True
   .RefreshStyle = xlInsertDeleteCells
   .SavePassword = False
   .SaveData = True
   .AdjustColumnWidth = True
   .RefreshPeriod = 0
   .WebSelectionType = xlAllTables
   .WebFormatting = xlWebFormattingNone
   .WebPreFormattedTextToColumns = True
   .WebConsecutiveDelimitersAsOne = True
   .WebSingleBlockTextImport = False
   .WebDisableDateRecognition = False
   .WebDisableRedirections = False
   .Refresh BackgroundQuery:=False
 End With
 wk.SaveAs strExcelFile
 wk.Close
 eApp.Quit
  Exit Sub
Control:
  MsgBox Err.Description & " Error: " & Err.Number & vbCr
End Sub

Public Function Exportar_ADO_Excel()
      On Error GoTo Control
        '-Dim cn          As New ADODB.Connection
        Dim rec         As New ADODB.Recordset
        Dim Excel       As Object
        Dim Libro       As Object
        Dim Hoja        As Object
        Dim ArrData     As Variant
        Dim iRec        As Long
        Dim iCol        As Integer
        Dim iRow        As Integer
                
        Dim sqlTabla, strExcelFile      As String
        
        Select Case wTpoQuery
        Case "0"
          strExcelFile = wNameXLS           ' "C:\RPT\QMULTIPLE.xls"
          sqlTabla = "SELECT * FROM #Tabla0"
          wNombFile = "QMULTIPLE"
          
        Case "0.1"
          strExcelFile = wNameXLS           ' "C:\RPT\RESDSCTS.xls"
          sqlTabla = "SELECT * FROM CpTmResDscto"
          wNombFile = "RESDSCTS"
          
        Case "0.5"
          strExcelFile = wNameXLS            '  "C:\RPT\PROGUPR.xls"
          sqlTabla = "SELECT * FROM XProgUPR"
          wNombFile = "ProgUPR"
        Case "1"
          strExcelFile = wNameXLS           ' "C:\RPT\QBOLTALTA.xls"
          sqlTabla = "SELECT * FROM #Tabla5"
          wNombFile = "QBOLTALTA"
          
        Case "1.0"
          strExcelFile = wNameXLS            '"C:\RPT\RESMP.xls"
          sqlTabla = "SELECT * FROM CpTmResMP"
          wNombFile = "RESMP"
          
        Case "1.1"
          strExcelFile = wNameXLS          ' "C:\RPT\QMULTIPLE.xls"
          sqlTabla = "SELECT * FROM CpTmgeneral"
          wNombFile = "QMULTIPLE"
          
        Case "1.5"
          strExcelFile = wNameXLS               '"C:\RPT\PGFiltrado.xls"
          sqlTabla = "SELECT * FROM #TabPGFilt"
          wNombFile = "PGFiltrado"
         
        Case "2"
          strExcelFile = wNameXLS            ' "C:\RPT\QEMERGENC.xls"
          sqlTabla = "SELECT * FROM #Tabla "         'RelRpeTmp  =>  Atenc Emergenc
          wNombFile = "QEMERGENC"
        
        Case "2.0"
          strExcelFile = wNameXLS            '  "C:\RPT\RESJGUAR.xls"
          sqlTabla = "SELECT * FROM CpTmResJGuar"
          wNombFile = "RESJGUAR"
        
        Case "2.1"
          strExcelFile = wNameXLS           ' "C:\RPT\QCONSESPE.xls"
          sqlTabla = "SELECT * FROM #Tabla "
          wNombFile = "QCONSESPE"
        Case "2.2"
          strExcelFile = wNameXLS           ' "C:\RPT\QUERYUCI.xls"
          sqlTabla = "SELECT * FROM #TablaUCI "     'RelUCITmp  =>  Cateterismo / Atenc UCI
          wNombFile = "QUERYUCI"
        Case "2.9"
          strExcelFile = wNameXLS           ' "C:\RPT\QICONSULTA.xls"
          sqlTabla = "SELECT * FROM #Tabla70 "      'RelICTMP  =>  InterConsulta
          wNombFile = "QICONSULTA"
          
        Case "3"
          strExcelFile = wNameXLS           ' "C:\RPT\AUDITAREME.xls"
          sqlTabla = "SELECT * FROM #Tabla6 "      'Auditar EMer
          wNombFile = "AUDITAREME"
        Case "4"
          strExcelFile = wNameXLS           ' "C:\RPT\AUDITAHOSP.xls"
          sqlTabla = "SELECT * FROM #Tabla4 "      ' RelAHtmp
          wNombFile = "AUDITAHOSP"
          
        Case "4.0"
          strExcelFile = wNameXLS          ' "C:\RPT\RESPGABO.xls"
          sqlTabla = "SELECT * FROM CpTmDifPgoGuar"
          wNombFile = "RESPGABO"
          
        Case "4.3"
          strExcelFile = wNameXLS          ' "C:\RPT\SUMARIOtot.xls"
          sqlTabla = "SELECT * FROM #TablaST "
          wNombFile = "SUMARIOTOT"
        Case "5.2"
          strExcelFile = wNameXLS          ' "C:\RPT\SUMARIOdet.xls"
          sqlTabla = "SELECT * FROM TablaSD "
          wNombFile = "SUMARIODET"
        
        Case "5"
          strExcelFile = wNameXLS           ' "C:\RPT\AUDITAREPI.xls"
          sqlTabla = "SELECT * FROM #Tabla7 "      ' RelEpiTmp  =>  Auditar EPIC
          wNombFile = "AUDITAREPI"
          
        Case "5.0"
          strExcelFile = wNameXLS           '  "C:\RPT\RESPGDSC.xls"
          sqlTabla = "SELECT * FROM CpTmDifPgoGuar"
          wNombFile = "RESPGDSC"
          
        Case "5.1"
          strExcelFile = wNameXLS           ' "C:\RPT\AUDCHGEPI.xls"
          sqlTabla = "SELECT * FROM #Tabla4 "      ' RelEpiTmp  =>  Auditar EPIC
          wNombFile = "AUDCHGEPI"
        Case "6"
          strExcelFile = wNameXLS           ' "C:\RPT\EnvioServ.xls"
          sqlTabla = "SELECT * FROM #Tabla3 "      '   RelBuzTmp
          wNombFile = "EnvioServ"
        Case "7"
          strExcelFile = wNameXLS           ' "C:\RPT\PACHOSP.xls"
          sqlTabla = "SELECT * FROM #Tabla7 "      '  RelAEtmp
          wNombFile = "PACHOSP"
          
        Case "7.0"
          strExcelFile = wNameXLS            '"C:\RPT\SUMRecord.xls"
          sqlTabla = "SELECT * FROM #Tabla5"
          wNombFile = "#Tabla5"
          
        Case "7.7"
            strExcelFile = wNameXLS         '"C:\RPT\DirPrograma.xls"
            sqlTabla = "SELECT * FROM #Tabla77"
            wNombFile = "DirPrograma"
        Case "7.8"
            strExcelFile = wNameXLS         '"C:\RPT\DirProgALL.xls"
            sqlTabla = "SELECT * FROM #Tabla78"
            wNombFile = "DirProgALL"
          
        Case "8"
          strExcelFile = wNameXLS           ' "C:\RPT\AUDCHGPWD.xls"
          sqlTabla = "SELECT * FROM #TablaPWD "      '  RelUserPWD => #TablaPWD
          wNombFile = "AUDCHGPWD"
        Case "9"
          strExcelFile = wNameXLS           ' "C:\RPT\Egreso_TM.xls"
          sqlTabla = "SELECT * FROM #TablaTM "      ' RefTelMonTMP (modificado) => #TablaTM
          wNombFile = "Egreso_TM"
          
        Case "9.0"
           strExcelFile = wNameXLS          ' "C:\RPT\CpAnalisis.xls"
           sqlTabla = "SELECT * FROM #tabla9"
           wNombFile = "CpAnalisis"
           
        Case "10"
           strExcelFile = wNameXLS          ' "C:\RPT\SUMAutori.xls"
           sqlTabla = "SELECT * FROM #tablaMI"
           wNombFile = "SUMAutori"
  
        End Select
        
        '-Si fichero excel existe, aqui es borrado
        If Dir(strExcelFile) <> "" Then Kill strExcelFile
       
        'Exportar_ADO_Excel(sPathDB, Consulta, "c:\libro.xLS")
        Set rec = New ADODB.Recordset
        ' -- Abrir el Recordset pasándole la cadena sql
        rec.Open sqlTabla, Cnx, adOpenStatic
        ' -- Crear los objetos para utilizar el Excel
        Set Excel = CreateObject("Excel.Application")
        Set Libro = Excel.Workbooks.Add
          
        ' -- Hacer referencia a la hoja
        Set Hoja = Libro.Worksheets(1)
          
        Excel.Visible = True: Excel.UserControl = True
        iCol = rec.Fields.Count
        For iCol = 1 To rec.Fields.Count
            Hoja.Cells(1, iCol).Value = rec.Fields(iCol - 1).Name
        Next
          
        If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
            Hoja.Cells(2, 1).CopyFromRecordset rec
        Else
      
            ArrData = rec.GetRows
      
            iRec = UBound(ArrData, 2) + 1
              
            For iCol = 0 To rec.Fields.Count - 1
                For iRow = 0 To iRec - 1
      
                    If IsDate(ArrData(iCol, iRow)) Then
                        ArrData(iCol, iRow) = Format(ArrData(iCol, iRow))
      
                    ElseIf IsArray(ArrData(iCol, iRow)) Then
                        ArrData(iCol, iRow) = "Array Field"
                    End If
                Next iRow
            Next iCol
                  
            ' -- pasa los datos a la hoja de Excel
            Hoja.Cells(2, 1).Resize(iRec, rec.Fields.Count).Value = GetData(ArrData)
        End If
      
        Excel.Selection.CurrentRegion.Columns.AutoFit
        Excel.Selection.CurrentRegion.Rows.AutoFit
      
        ' -- Cierra el recordset y la base de datos y los objetos ADO
        rec.Close
        'Cnx.Close
          
        Set rec = Nothing
        'Set Cnx = Nothing
        ' -- graba el libro XLS
        Libro.SaveAs strExcelFile
        Libro.Close
        ' -- Elimina las referencias Xls
        Set Hoja = Nothing
        Set Libro = Nothing
        Excel.Quit
        Set Excel = Nothing
    
      Exit Function
Control:
  MsgBox Err.Description & " Error: " & Err.Number & vbCr

End Function

 Public Function GetData(vValue As Variant) As Variant
        Dim X As Long, Y As Long, xMax As Long, yMax As Long, T As Variant
          
        xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
          
        ReDim T(xMax, yMax)
        For X = 0 To xMax
            For Y = 0 To yMax
                T(X, Y) = vValue(Y, X)
            Next Y
        Next X
          
        GetData = T
 End Function


Public Function CreaDir(valdir As String)
   carpeta = Dir(valdir & "\", vbDirectory)    'comprueba si ya existe la carpeta "fautos"
   If carpeta = "" Then
      MkDir valdir
      MsgBox "Carpeta creada " & valdir
   End If
  ' Dim AttrDev%
  ' On Error Resume Next
  ' AttrDev = GetAttr(valdir)
  ' If Err.Number Then
  '    Err.Clear
  '    MkDir valdir
  '    MsgBox "Creo Dir " & valdir
  '  End If
End Function

Public Sub acceso(X)
   wacceso = False
   Select Case X
   Case "11"
     If Mid(wacceso1, 1, 1) = "X" Then wacceso = True
    Case "12"
     If Mid(wacceso1, 2, 1) = "X" Then wacceso = True
    Case "13"
     If Mid(wacceso1, 3, 1) = "X" Then wacceso = True
   Case "14"
     If Mid(wacceso1, 4, 1) = "X" Then wacceso = True
   Case "15"
     If Mid(wacceso1, 5, 1) = "X" Then wacceso = True
   Case "16"
     If Mid(wacceso1, 6, 1) = "X" Then wacceso = True
   Case "17"
     If Mid(wacceso1, 7, 1) = "X" Then wacceso = True
   Case "18"
     If Mid(wacceso1, 8, 1) = "X" Then wacceso = True
   Case "19"
     If Mid(wacceso1, 9, 1) = "X" Then wacceso = True
   Case "110"
     If Mid(wacceso1, 10, 1) = "X" Then wacceso = True
   Case "111"
     If Mid(wacceso1, 11, 1) = "X" Then wacceso = True
   Case "112"
     If Mid(wacceso1, 12, 1) = "X" Then wacceso = True
   Case "113"
     If Mid(wacceso1, 13, 1) = "X" Then wacceso = True
   Case "114"
     If Mid(wacceso1, 14, 1) = "X" Then wacceso = True
   Case "115"
     If Mid(wacceso1, 15, 1) = "X" Then wacceso = True
   
   Case "21"
     If Mid(wacceso2, 1, 1) = "X" Then wacceso = True
    Case "22"
     If Mid(wacceso2, 2, 1) = "X" Then wacceso = True
    Case "23"
     If Mid(wacceso2, 3, 1) = "X" Then wacceso = True
    Case "24"
     If Mid(wacceso2, 4, 1) = "X" Then wacceso = True
    Case "25"
     If Mid(wacceso2, 5, 1) = "X" Then wacceso = True
    Case "26"
     If Mid(wacceso2, 6, 1) = "X" Then wacceso = True
    Case "27"
     If Mid(wacceso2, 7, 1) = "X" Then wacceso = True
    Case "28"
     If Mid(wacceso2, 8, 1) = "X" Then wacceso = True
    Case "29"
     If Mid(wacceso2, 9, 1) = "X" Then wacceso = True
    
    Case "31"
     If Mid(wacceso3, 1, 1) = "X" Then wacceso = True
    Case "32"
     If Mid(wacceso3, 2, 1) = "X" Then wacceso = True
    Case "33"
     If Mid(wacceso3, 3, 1) = "X" Then wacceso = True
    Case "34"
     If Mid(wacceso3, 4, 1) = "X" Then wacceso = True
    Case "35"
     If Mid(wacceso3, 5, 1) = "X" Then wacceso = True
    Case "36"
     If Mid(wacceso3, 6, 1) = "X" Then wacceso = True
    Case "37"
     If Mid(wacceso3, 7, 1) = "X" Then wacceso = True
    Case "38"
     If Mid(wacceso3, 8, 1) = "X" Then wacceso = True
    Case "39"
     If Mid(wacceso3, 9, 1) = "X" Then wacceso = True
     
    Case "41"
     If Mid(wacceso4, 1, 1) = "X" Then wacceso = True
    Case "42"
     If Mid(wacceso4, 2, 1) = "X" Then wacceso = True
    Case "43"
     If Mid(wacceso4, 3, 1) = "X" Then wacceso = True
    Case "44"
     If Mid(wacceso4, 4, 1) = "X" Then wacceso = True
    Case "45"
     If Mid(wacceso4, 5, 1) = "X" Then wacceso = True
    Case "46"
     If Mid(wacceso4, 6, 1) = "X" Then wacceso = True
    Case "47"
     If Mid(wacceso4, 7, 1) = "X" Then wacceso = True
    Case "48"
     If Mid(wacceso4, 8, 1) = "X" Then wacceso = True
    Case "49"
     If Mid(wacceso4, 9, 1) = "X" Then wacceso = True
     
    Case "51"
     If Mid(wacceso5, 1, 1) = "X" Then wacceso = True
    Case "52"
     If Mid(wacceso5, 2, 1) = "X" Then wacceso = True
    Case "53"
     If Mid(wacceso5, 3, 1) = "X" Then wacceso = True
    Case "54"
     If Mid(wacceso5, 4, 1) = "X" Then wacceso = True
    Case "55"
     If Mid(wacceso5, 5, 1) = "X" Then wacceso = True
    Case "56"
     If Mid(wacceso5, 6, 1) = "X" Then wacceso = True
    Case "57"
     If Mid(wacceso5, 7, 1) = "X" Then wacceso = True
    Case "58"
     If Mid(wacceso5, 8, 1) = "X" Then wacceso = True
    Case "59"
     If Mid(wacceso5, 9, 1) = "X" Then wacceso = True
   
   End Select
End Sub
Public Sub WVencripta(cclave As String)
   Dim i As Integer
   Dim wchr, wvasc As String
   wvlen = Len(cclave)
   For i = 1 To wvlen
      wchr = Mid(cclave, i, 1)
      wvasc = Asc(wchr) + (10 + i)
      Encclave = Encclave + Chr(wvasc)
   Next
End Sub
Public Sub WVdesencripta(cclave As String)
   Dim i As Integer
   Dim wchr, wvasc As String
   wvlen = Len(cclave)
   For i = 1 To wvlen
      wchr = Mid(cclave, i, 1)
      wvasc = Asc(wchr) - (10 + i)
      Encclave = Encclave + Chr(wvasc)
   Next
End Sub

Public Sub LeerINI()
 'Cnxheper.ini
 '===========
 '2 = 192.168.1.19           DSource
 '3 = 00:1C:C0:50:8A:7F      MAC
 '4 = SQLOLEDB               Provider
 '5 = HepMaestSql            Icatal0
 '6 = hepControlP            Icatal
 '7 = usr_hep                usr
 '8 = hep                    pwd
 
   'Las variable que no están definidas dentro de la función, son globales definidas en otra parte del código
     On Error GoTo errSub
      Dim n_File, npos, nlen As Integer
      Dim linea, texto As String
      Dim lectura As String
      Dim arr() As String
      Dim path As String
      n_File = FreeFile 'Asignamos el numero de archivo
      'path = CurrentProject.path & "\cnxheper.ini"  'Indicamos en una cadena el path donde se encuentra el .ini
      path = "C:\rpt\cnxheper.ini"
      Open path For Input As n_File 'Abrimos el archivo
     'Montamos un bucle para cada línea del archivo hasta que este se acabe.
      Do While Not EOF(n_File)
         Line Input #n_File, lectura
         'linea = linea & ";" & lectura 'Guardo variable todas las líneas separadas por '='
         npos = InStr(lectura, "=")
         nlen = Len(lectura)
         If npos > 0 Then
            texto = UCase(Trim(Left(lectura, npos - 1)))
            Encclave = ""
            WVdesencripta (Right(lectura, nlen - npos - 1))
            Select Case texto
               Case "2"
                 cDSource = Encclave
               Case "3"
                 cMAC = Encclave
               Case "4"
                 cProvider = Right(lectura, nlen - npos - 1)
               Case "5"
                 cICatal0 = Encclave
               Case "6"
                 cICatal = Encclave
               Case "7"
                 cUsr = Encclave
               Case "8"
                 cPwd = Encclave
            End Select
         End If
      Loop
      'arr = Split(linea, "=") 'Monto array con función Split       'Guardo en cada variable todos los datos que hay en cada   variable del .ini
      'cDsource = arr(2)
      'Cmac = arr(4)
      'cProvider = arr(6)
      Exit Sub
errSub:
    MsgBox "Número de error: " & Err.Number & vbNewLine & _
    "Descripción del error: " & Err.Description, vbCritical
End Sub
      
   
Public Sub CalcTiempo_click()
    ''On Error GoTo Error
    MSG = ""
    nndias = 0
    nnmeses = 0
    nnanos = 0
    Call ValidarFecha
    If MSG <> "" Then
        MsgBox MSG, vbInformation, App.Title
        Exit Sub
    End If
    FecINICIO = Format(FecINICIO, "dd/mm/yyyy")
    FecFIN = Format(FecFIN, "dd/mm/yyyy")
    nnanos = Year(FecFIN) - Year(FecINICIO)
    If Month(FecFIN) = Month(FecINICIO) Then
       nnmeses = 0
    ElseIf Month(FecFIN) > Month(FecINICIO) Then
        nnmeses = Month(FecFIN) - Month(FecINICIO)
    ElseIf Month(FecFIN) < Month(FecINICIO) Then
        nnmeses = 12 + (Month(FecFIN) - Month(FecINICIO))
        nnanos = nnanos - 1
    End If
    
    If Day(FecFIN) >= Day(FecINICIO) Then
        nndias = Day(FecFIN) - Day(FecINICIO)
    ElseIf Day(FecFIN) < Day(FecINICIO) Then
        nnmeses = nnmeses - 1
        If nnmeses < 0 Then
            nnanos = nnanos - 1
            nnmeses = nnmeses + 12
        End If
        If Month(FecINICIO) = 1 Or Month(FecINICIO) = 3 Or Month(FecINICIO) = 5 Or Month(FecINICIO) = 7 Or Month(FecINICIO) = 8 Or Month(FecINICIO) = 10 Or Month(FecINICIO) = 12 Then
           nndias = Day(FecFIN) - Day(FecINICIO) + 31
        ElseIf Month(FecINICIO) = 2 Then
            nndias = Day(FecFIN) - Day(FecINICIO) + 28
            If Year(FecINICIO) Mod 4 = 0 Then nndias = nndias + 1
        Else
            nndias = Day(FecFIN) - Day(FecINICIO) + 30
        End If
    End If
Exit Sub
''Error:
''MsgBox "Error para Calcular.. " & Error, vbCritical, App.Title
End Sub
Public Sub ValidarFecha()
  If FecFIN = "" Then
    MSG = "Verifique la fecha Final *" & FecFIN & "*"
    Exit Sub
  End If
  If FecINICIO = "" Then
    MSG = "Verifique la fecha de Inicio *" & FecINICIO & "*"
    Exit Sub
  End If
  If DateValue(FecINICIO) > DateValue(FecFIN) Then
     Exit Sub
  End If
  MSG = ""
End Sub

Public Sub ObtenerPrimerUltimoDia(fecha As Date)
    'Pasarle a esta rutina la fecha que queremos averiguar
    'Dim primer As Date
    'Dim ultimo As Date
    'Usamos la funcion DAteSerial para obtener el primero y el ultimo dia
    primer = DateSerial(Year(fecha), Month(fecha) + 0, 1)
    ultimo = DateSerial(Year(fecha), Month(fecha) + 1, 0)
    
   'MsgBox WeekdayName(Weekday(fecha))       muestra Nombre de dia

    'MsgBox " Primer día : " & WeekdayName(Weekday(fecha)) & "," & primer & vbNewLine & _
    '       " Último día : " & WeekdayName(Weekday(fecha)) & "," & ultimo, vbInformation
End Sub

Public Sub VerHrsMint(nvalor)
  Dim nhors, nmints As Integer
  nhors = Int(nvalor)
  nmints = (nvalor - nhors) * 60
  MSG = (Trim(Str(nhors)) + ":" + Trim(Str(nmints)))
End Sub
Public Function cHorsMint(nvalor) As String
  Dim nhors, nmints As Integer
  nhors = Int(nvalor)
  nmints = (nvalor - nhors) * 60
  cHorsMint = ""
  If nhors + nmints > 0 Then
     cHorsMint = (Trim(Str(nhors)) + ":" + Trim(Str(nmints)))
  End If
End Function
Public Sub ConvertDecimalenHora(d As Single)
  Dim e As Single
  e = CStr(Round((d - Int(d)) / 100 * 60, 2)) & "0"
  MSG = CStr(Int(d)) + ":" + Mid(e, 3, 2)
  Exit Sub
End Sub

Public Function nTotMints(cInicio, cFin) As Integer
  nTotMints = 0
  If Val(Left(cFin, 2)) - Val(Left(cInicio, 2)) >= 0 Then
     nTotMints = (Val(Left(cFin, 2)) - Val(Left(cInicio, 2))) * 60 + Val(Right(cFin, 2)) - Val(Right(cInicio, 2))
  End If
End Function

Public Function CalcMints(cInicio, cFin) As Integer
   CalcMints = 0
   If Val(Left(cFin, 2)) - Val(Left(cInicio, 2)) >= 0 Then
      CalcMints = (Val(Left(cFin, 2)) - Val(Left(cInicio, 2))) * 60 + Val(Right(cFin, 2)) - Val(Right(cInicio, 2))
   Else
      CalcMints = (24 + Val(Left(cFin, 2)) - Val(Left(cInicio, 2))) * 60 + Val(Right(cFin, 2)) - Val(Right(cInicio, 2))
   End If
End Function
Public Function CalcMTard(p_ing, m_ing) As Integer
   CalcMTard = 0
   CalcMTard = (Val(Left(m_ing, 2)) - Val(Left(p_ing, 2))) * 60 + Val(Right(m_ing, 2)) - Val(Right(p_ing, 2))
End Function
Public Function FinOk(cFin, cInicio, choras) As Boolean
   FinOk = False
   nLapso = (Val(Left(cInicio, 2)) + Val(Left(choras, 2)) - Val(Left(cFin, 2))) * 60 + Val(Right(cInicio, 2)) + Val(Right(choras, 2)) - Val(Right(cFin, 2))
   FinOk = IIf(nLapso >= 0, True, False)
End Function
Public Function EsTexto(ctexto) As Boolean
   ctexto = Trim(ctexto)
   EsTexto = True
   nLapso = Len(ctexto)
   Dim i As Integer
   For i = 1 To nLapso
       If InStr("0123456789", Mid$(ctexto, i, 1)) > 0 Then
          EsTexto = False
       End If
   Next i
End Function

Public Function IniciaClave()
    '-Aplica solo Programadores UPR
    Dim rt As New ADODB.Recordset
    Dim rt1 As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim W_ok, Sq0, sql, sq2, Sqt, c_Costo, c_upr, cDocumento, cTarjeta, c_claveOld, fecha1 As String
    Dim i As Integer
    nLapso = 0
    If MsgBox("Reinicia Claves ", vbOKCancel + vbQuestion, "Programadores UPR") <> vbCancel Then
        fecha1 = Format(Date, "yyyymmdd")
        Set rt = New Recordset
        Sq0 = "SELECT * FROM CpAudClaves WHERE fecha='" & fecha1 & "'"
        rt.Open Sq0, Cnx, adOpenStatic, adLockOptimistic, adCmdText
        
        nLapso = 0
        sql = "SELECT * FROM CpCCostoUpr ORDER BY CodUpr"
        Set rs = New Recordset
        rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not rs.EOF
           c_Costo = rs!CCosto
           c_upr = rs!CodUpr
           For i = 1 To 3
              W_ok = "N"
              Select Case i
                Case 1
                  If IsNull(rs!Dni1) Or Trim(rs!Dni1) = "" Then
                  Else
                     If rs!editor1 = "S" Then
                        W_ok = "S"
                        nLapso = nLapso + 1
                        cDocumento = Trim(rs!Dni1)
                        c_claveOld = Trim(rs!claveu1)
                        Encclave = ""
                        WVencripta (Trim(rs!Dni1))
                        Sqt = "UPDATE CPGeneral SET clave='" + Encclave + "', ActClave='SI' WHERE (baja is NULL) and Documento='" + Trim(rs!Dni1) + "'"
                        Cnx.Execute Sqt
                     End If
                  End If
                Case 2
                  If IsNull(rs!Dni2) Or Trim(rs!Dni2) = "" Then
                  Else
                     If rs!editor2 = "S" Then
                         W_ok = "S"
                        nLapso = nLapso + 1
                        cDocumento = Trim(rs!Dni2)
                        c_claveOld = Trim(rs!claveu2)
                        Encclave = ""
                        WVencripta (Trim(rs!Dni2))
                        Sqt = "UPDATE CPGeneral SET clave='" + Encclave + "', ActClave='SI'  WHERE  (baja is NULL) and Documento='" + Trim(rs!Dni2) + "'"
                        Cnx.Execute Sqt
                     End If
                  End If
                Case 3
                  If IsNull(rs!dni3) Or Trim(rs!dni3) = "" Then
                  Else
                     If rs!editor3 = "S" Then
                         W_ok = "S"
                        nLapso = nLapso + 1
                        cDocumento = Trim(rs!dni3)
                        c_claveOld = Trim(rs!claveu3)
                        Encclave = ""
                        WVencripta (Trim(rs!dni3))
                        Sqt = "UPDATE CPGeneral SET clave='" + Encclave + "', ActClave='SI' WHERE  (baja is NULL) and Documento='" + Trim(rs!dni3) + "'"
                        Cnx.Execute Sqt
                     End If
                  End If
              End Select
              
              If W_ok = "S" Then
                 cTarjeta = ""
                 sq2 = "SELECT * FROM Cpgeneral WHERE  (baja is NULL) and documento='" + cDocumento + "'"
                 Set rt1 = New Recordset
                 rt1.Open sq2, Cnx, adOpenStatic, adLockReadOnly, adCmdText
                 If Not rt1.EOF() Then
                    cTarjeta = rt1!tarjeta
                 End If
                 rt1.Close
                 rt.AddNew
                 rt!tarjeta = cTarjeta
                 rt!documento = cDocumento
                 rt!claveold = c_claveOld
                 rt!claveNew = cDocumento
                 rt!CCosto = c_Costo
                 rt!CodUpr = c_upr
                 rt!fecha = Date
                 rt!hora = Format(Time, "HH:mm:ss")
                 rt!usuario = wNick         'Firma Usuario
                 rt!estacion = w_clienName
                 rt.Update
              End If
           Next
           rs.MoveNext
        Loop
        rs.Close
        rt.Close
        MsgBox "Inicializo " + Trim(Str(nLapso)) + " claves.", vbInformation
   End If
   
   nLapso = 0
   If MsgBox("Genera Plantilla, Serv. asistenciales. ", vbOKCancel + vbQuestion, "Plantilla-Directorio") <> vbCancel Then
        Dim unafecha As Date
        unafecha = "01/" + Trim(Str(genmes)) + "/" + Trim(Str(genaño))
        Call ObtenerPrimerUltimoDia(unafecha)
        
        sql = "SELECT * FROM CpDirectorio  WHERE ano ='" & genaño & "' AND mes='" & Trim(Str(genmes)) & "' and flag='0' order by tarjeta"
        Set rs = New Recordset
        rs.Open sql, Cnx, adOpenStatic, adLockOptimistic, adCmdText
        Dim z As String
        'MsgBox "Proceso plantilla " + Format(primer, "dd/mm/yyyy") + " al " + Format(ultimo, "dd/mm/yyyy")
        nLapso = 0
        Do While Not rs.EOF
           '-Imagen hasta primer dia < Hoy
           cPlantilla0 = ""
           For i = Day(primer) To Day(ultimo)
              z = i - 1
              cPlantilla0 = cPlantilla0 + IIf(IsNull(rs.Fields(z + 4)), "==", Left(Trim(rs.Fields(z + 4)) + "==", 2))
              '-D N D MTD*V   N*...
              ' 1234567890123456
           Next i
           rs!plantilla0 = cPlantilla0
           rs.Update
           nLapso = nLapso + 1
           'If nLapso < 6 Then
           '   MsgBox rs!tarjeta + " van " + Trim(Str(nLapso))
           'Else
           '   rs.Close
           '   End
           'End If
           rs.MoveNext
        Loop
        rs.Close
        MsgBox "Completo plantillas 0, serv. asistenciales.." + Trim(Str(nLapso)), vbInformation
    End If
    
End Function

Public Sub Ver_xls()         '----- VERSION   XLS ------------
Dim OExcel As Object
Dim TempVer As Variant
Set OExcel = CreateObject("Excel.Application")

On Error Resume Next

TempVer = OExcel.Version

If Err = 0 Then
   '''Ver = TempVer
    MsgBox TempVer
    
   'La version 12, se refiere al Office 2007
   'La version 11, se refiere al Office 2003
   'La version 10, se refiere al Office XP
   'La version 9, se refiere al Office 2000
    
End If

OExcel.Quit
Set OExcel = Nothing

Err.Clear

End Sub


Public Sub CrxToPrinter()
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As CRAXDRT.Report
    Dim crxTables As CRAXDRT.DatabaseTables
    Dim crxTable As CRAXDRT.DatabaseTable
    
    Dim strServerOrDSNName As String
    Dim strDBNameOrPath As String
    Dim strUserID As String
    Dim strPassword As String
    Dim strSelectionfrm As String
    Dim Rst As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    Dim crxParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crxParamDef As CRAXDRT.ParameterFieldDefinition
   ''' On Error GoTo ErrHandler
  
    strServerOrDSNName = cDSource    '"192.168.1.19,1433" replace servername with your SQL connection server
    strDBNameOrPath = cICatal        '"hepasistencial" replace databasename with your SQL connection databasename
    strUserID = cUsr                 '"sa"  replace userid with your SQL connection userid
    strPassword = cPwd               '"hep" replace password with your SQL connection password
    '-Valor asignar en Parametros
    Dim sql, wimpreso, wTitulo, wdetalle, wdetalle1, wcmp0, wdscCmp0 As String
    Dim n_copia As Integer
    wimpreso = wUsuNombre
    If nTipRpt = 5.5 Then
       wdetalle = "Reporte Sumario Asistencia del " & RecMarcasQuery.Txpridia & "  Al " & RecMarcasQuery.TxUltDia
       '--   wdetalle = "Reporte Sumario Asistencia del " & RecMasRS.Txpridia & "  Al " & RecMasRS.TxUltDia
       wnombrep = "(CpRecMasA4RS)"
    ElseIf nTipRpt = 5.8 Then
        wdetalle = "Reporte mensual de asistencia del " & RecMarcasQuery.Txpridia & "  Al " & RecMarcasQuery.TxUltDia
    End If
 
    '-Open the report
    Set crxApp = New CRAXDRT.Application
    If nTipRpt = 5.5 Then
       Set crxRpt = crxApp.OpenReport(App.path & "\CpRecMasRS.rpt", 1)
    ElseIf nTipRpt = 5.8 Then
        Set crxRpt = crxApp.OpenReport(App.path & "\CpRecUprTarjeta.rpt", 1)
    End If
    
    '-Remove any saved data
    crxRpt.DiscardSavedData
    '-Oculta pantalla de ingreso de valores Parametro del reporte
    crxRpt.EnableParameterPrompting = False
   
    '-Set the connection for the report.
    '-SetLogOnInfo is a deprecated method in CR 9, but still works.
    '-The suggested method for CR9 is to use the ConnectionProperty object.
    crxRpt.Database.Tables(1).SetLogOnInfo strServerOrDSNName, strDBNameOrPath, strUserID, strPassword
    
    '-This removes the schema from the Database Table's Location property.
    Set crxTables = crxRpt.Database.Tables
    For Each crxTable In crxTables
        With crxTable
            .Location = .Name
        End With
    Next
  
    '-CR Seleccionar  registros Distintivos entre 2 tablas Padre-hijo
    If nTipRpt = 5.5 Then
          'CpTmRecord -> CpTmGeneral
          sql = "SELECT DISTINCT " & _
          "CpTmRecord2.tarjeta, CpTmRecord2.fecha, CpTmRecord2.NombDia, CpTmRecord2.turno, CpTmRecord2.ming, CpTmRecord2.marca2, CpTmRecord2.marca3, CpTmRecord2.marca4, CpTmRecord2.tardanza, CpTmRecord2.Descta, CpTmRecord2.hrProg, CpTmRecord2.hrTrab, " & _
          "CpTmRecord2.Autorizac, CpTmRecord2.DatoMp, CpTmRecord2.hrini, CpTmRecord2.hrfin, CpTmRecord2.Comision, CpTmRecord2.GRUPO, " & _
          "CpTmGeneral.trabajador, CpTmGeneral.dscactividad, CpTmGeneral.dscccosto, CpTmGeneral.dsccargo, CpTmGeneral.horasP, CpTmGeneral.htarda, CpTmGeneral.hdscto, CpTmGeneral.horasmP, CpTmGeneral.horasFJ, " & _
          "CpTmGeneral.horasFI,CpTmGeneral.horasFmp, CpTmGeneral.horast " & _
          "FROM  CpTmRecord2 INNER JOIN CptmGeneral ON CpTmRecord2.tarjeta=CptmGeneral.tarjeta " & _
          "ORDER BY CpTmRecord2.tarjeta,CpTmRecord2.fecha"
    ElseIf nTipRpt = 5.8 Then
          'Oficina-servidor-PERIODO   CpTmRecord -> CpTmGeneral
          sql = "SELECT DISTINCT " & _
          "CpTmrecord2.tarjeta,CpTmrecord2.MES,CpTmrecord2.tarjetaMes, CpTmrecord2.fecha, CpTmrecord2.NombDia, CpTmrecord2.turno, CpTmrecord2.ming, CpTmrecord2.marca2, CpTmrecord2.marca3, CpTmrecord2.marca4, CpTmrecord2.tardanza, CpTmrecord2.Descta, CpTmrecord2.hrProg, CpTmrecord2.hrTrab, " & _
          "CpTmrecord2.Autorizac, CpTmrecord2.DatoMp, CpTmrecord2.hrini, CpTmrecord2.hrfin, CpTmrecord2.Comision, CpTmrecord2.GRUPO, " & _
          "CpTmGeneral.trabajador, CpTmGeneral.dscactividad, CpTmGeneral.dscccosto, CpTmGeneral.dsccargo, CpTmGeneral.horasP, CpTmGeneral.htarda, CpTmGeneral.hdscto, CpTmGeneral.horasmP, CpTmGeneral.horasFJ, " & _
          "CpTmGeneral.horasFI,CpTmGeneral.horasFmp, CpTmGeneral.horast " & _
          "FROM  CpTmrecord2 INNER JOIN CptmGeneral ON CpTmrecord2.tarjetaMes=CptmGeneral.tarjetaMES  " & _
          "ORDER BY CpTmrecord2.tarjetaMES,CpTmrecord2.fecha"
     End If
   
    '-Define Scope RecordSet
    '- Set rst = Cnx.Execute(SQL)
    Set Rst = New Recordset
    Rst.Open sql, Cnx, adOpenStatic, adLockReadOnly
    
    '-Carga Recordset
    crxRpt.Database.SetDataSource Rst
    '-Passing a SQL Query string
    '-The select query to return the data required
    'crxRpt.SQLQueryString = sql   'Input your SELECT query here
 
    '-Asigna valor a Parametros del reporte
    Set crxParamDefs = crxRpt.ParameterFields
    For Each crxParamDef In crxParamDefs
      crxParamDef.ClearCurrentValueAndRange
      Select Case crxParamDef.ParameterFieldName
        Case "Razonsocial"
            crxParamDef.AddCurrentValue wrazonsocial
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wrazonsocial
        Case "Sistema"
            crxParamDef.AddCurrentValue wsistema
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wsistema
        Case "Titulo"
            crxParamDef.AddCurrentValue wTitulo
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wTitulo
        Case "Detalle"
            crxParamDef.AddCurrentValue wdetalle
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wdetalle
        Case "Detalle1"
            crxParamDef.AddCurrentValue wdetalle1
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wdetalle
        Case "NombRep"
            crxParamDef.AddCurrentValue wnombrep
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wnombrep
        Case "Impreso"
            crxParamDef.AddCurrentValue wUsuNombre
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wusunombre
        Case "Campo1"
            crxParamDef.AddCurrentValue wcampo1
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wcampo1
        Case "Campo2"
            crxParamDef.AddCurrentValue wcampo2
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wcampo1
        Case "CCosto"
            crxParamDef.AddCurrentValue W_ccosto
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue w_ccosto
        Case "Upr"
            crxParamDef.AddCurrentValue W_upr
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue w_upr
        Case "Regplh"
            crxParamDef.AddCurrentValue W_RegPlh
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue w_RegPlh
        Case "Lbldh"
            crxParamDef.AddCurrentValue W_lbldh
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue w_lbldh
        Case "MesdeAno"
            crxParamDef.AddCurrentValue wMesdeAno
            '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wMesdeAno
      
      End Select
    Next
    
    '-Set the Paper source to the lower bin
    'crxRpt.PaperSource = crPRBinLower
    
    '-Set DUPLEX printing to horizontal
    'crxRpt.PrinterDuplex = crPRDPHorizontal
  
    '-Aplica formula para seleccionar registro
    'crxRpt.RecordSelectionFormula = strSelectionfrm
    
    If nTipRpt = 9.34 Then
       '-muestra dialogo para elegir print
       crxRpt.PrinterSetup (0)
    End If
    If nTipRpt = 5.5 Or nTipRpt = 5.8 Then
       '-Export to PDF
       crxRpt.ExportOptions.DiskFileName = "C:\Rpt\" + cObjeto + IIf(nTipRpt = 5.5, "S", "") + ".PDF"
       crxRpt.ExportOptions.DestinationType = crEDTDiskFile
       'crxRpt.ExportOptions.FormatType = crEFTWordForWindows
       '-Set the ExportOptions FormatType to PDF Format
       crxRpt.ExportOptions.FormatType = crEFTPortableDocFormat
       crxRpt.Export False
       
       '-MsgBox "Se Genero.. PDF "
    End If
    If nTipRpt = 99 Then
       '-Aplica formula para seleccionar registro... no funcionó!
       crxRpt.RecordSelectionFormula = strSelectionfrm
    End If
    '-Envia directo a impre
    '-Nro de copias
    'If nTipRpt = 2 Then
    '   n_copia = 2
    '   crxRpt.PrintOut False, n_copia
    'Else
    '--   crxRpt.PrintOut False
    'End If
     
    Screen.MousePointer = vbDefault
    ''Set crxParamFields = Nothing
    ''Set crxParamField = Nothing
    Rst.Close
    Set Rst = Nothing
    Exit Sub
'''ErrHandler:
'''  If Err.Number = -2147206461 Then
'''     MsgBox "El archivo de reporte no se encuentra, restáurelo ", vbCritical + vbOKOnly
'''  Else
'''     MsgBox "<E>RROR, " + Err.Description, vbCritical + vbOKOnly
'''  End If
  Screen.MousePointer = vbDefault
End Sub

Public Sub CompruebaTabla()
    
Dim sql As String
Dim rs As New ADODB.Recordset
Set rs = New Recordset
    
'-Para saber si existe una tabla escribe lo siguiente:
'sql = " SELECT * From INFORMATION_SCHEMA.Tables" & _
'" Where TABLE_TYPE = 'BASE TABLE' And TABLE_NAME = 'marca'" + w_ano
    
'-Si por ejemplo, quieres añadir una columna a un tabla sólo si ésta no existe haz lo siguiente:
'  IF not exists
'(
'SELECT *
'From INFORMATION_SCHEMA.Columns
'WHERE COLUMN_NAME = 'NOMBRE_COLUMNA' AND TABLE_NAME = 'NOMBRE_TABLA'
')
'BEGIN
'  ALTER TABLE NOMBRE_TABLA ADD NUEVO_CAMPO int
'End
    
   sql = "SELECT * From INFORMATION_SCHEMA.Tables  Where TABLE_TYPE = 'BASE TABLE' And TABLE_NAME = '" + NombTabla + "'"
   rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs.EOF Then
      wflg = "SI"
   End If
   rs.Close
   Set rs = Nothing
   Exit Sub
    
End Sub


Public Function CHKDias(cTarjeta As String, cCodigo As String)

  Dim rs As New Recordset
  Set rs = New Recordset
  Dim sql, ccadena, c_ano, cfecI, cFecF As String
  Dim i, nini, nfin As Integer
  ccadena = ""
  '-5años post a año Vac.
  nini = Val(Right(cCodigo, 2)) + 2000
  If nini + 5 > Year(Date) Then
     nfin = Year(Date)
  Else
     nfin = nini + 5
  End If
  For i = nini To nfin
      Set rs = New Recordset
      c_ano = Trim(Str(i))
      
'SELECT     DATEDIFF(day, FECHAINI, FECHAFIN) + 1 AS dias, TARJETA, CODIGO, FECHAINI, FECHAFIN, anulado
'From CpAutori2023
'WHERE     (TARJETA = '02520') AND (CODIGO = 'VA21') AND (anulado IS NULL)
      
      sql = "SELECT DATEDIFF(day, FECHAINI, FECHAFIN) + 1 AS dias, TARJETA, CODIGO, FECHAINI, FECHAFIN, anulado FROM CpAutori" & c_ano & " WHERE (anulado is NULL) and tarjeta ='" & Trim(cTarjeta) & "' AND codigo='" & cCodigo & "'"
      rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      Do While Not rs.EOF
          If InStr(ccadena, Format(rs!fechaIni, "yyyymmdd")) = 0 Then
             'nvo lapso
             ccadena = ccadena + (Format(rs!fechaIni, "yyyymmdd") + "_")
             cfecI = Format(rs!fechaIni, "DD/mm/yyyy")
             cFecF = Format(rs!fechaFin, "dd/mm/yyyy")
             nndias = nndias + rs!dias
          End If
          rs.MoveNext
      Loop
      '-- MsgBox (Trim(Str(nndias)) + "->" + c_ano)
      rs.Close
  Next
   
End Function


Public Sub ProcProg3RO(c_Tarjeta As String, c_Periodo As String)
    '-MsgBox "llego " + c_Tarjeta + ": " + c_Periodo
    '---08.09.2026 retirar carga BDCOLAS
    Dim cnx2 As ADODB.Connection
    Dim cDSource2, cICatal2, CString2, cUsr2, cPwd2 As String
    Dim cICatalC, CStringC, cUsrC, cPwdC As String

    '---ACTUALIZA Programa-Comedor desde dia = date (sujeto que propietario no sea "Nutrición") en adelante; No dias anteriores que fueron realizados.
    Dim rs As New ADODB.Recordset
    Dim r_d As New ADODB.Recordset
    Dim rsr As New ADODB.Recordset
    Dim rs_d As New ADODB.Recordset
    Dim rsrC As New ADODB.Recordset
    Dim rt As New ADODB.Recordset
    
    Dim cMesActual, c_mes, c_ano, c_DIA, c_criterio, cIniLac, cFinLac, dftope As String
    Dim nLacta, nSonLacta, nposicion, nvalor, nIdHORARIO As Integer
    Dim sql, Sq0, sqr, sqg, SqC As String
    Dim cTarjeta, cDni, capelPat, capelMat, cnombres, csexo, cAplica, wdscUpr As String
    Dim c_Costo, cCodUpr, cGrpGen, cCargo, cColas, cDscCargo As String
    Dim cHoyDia, xDia, CriterioDia, cvalor, p_ing, p_ing1, p_sal, p_sal1, p_horas, phora, pmints As String
    Dim cTurno, FechaB, cHoraI, cHoraF, cHoraI1, cHoraF1 As String
    Dim nHrProg As Double
    
    '------OJO: SERVIDOR de COMEDOR
    cDSource2 = "192.168.1.198"    '"192.168.1.198"       '"192.168.1.15"     '
    cICatal2 = "Heptrabajadores"
    cUsr2 = "sa"
    cPwd2 = "P4sscomedor"       '"P4sscomedor"             '"P4ssw0rd"   '
    
    On Error GoTo ErrProc    'caso error en parametros conex
    CString2 = "Provider=" + cProvider + ";Data Source='" + cDSource2 + "';" & _
       "initial catalog=" + cICatal2 + "; user id='" + cUsr2 + "'; password='" + cPwd2 + "';"
    Set cnx2 = New ADODB.Connection
    With cnx2
       .ConnectionString = CString2
       .Open
    End With
    cnx2.CommandTimeout = 200
      '--MsgBox "OK " + CString2
       
    '--Verifica Servidor SQL responde
    'If cnx2.State = 1 Then
    '   MsgBox "Servidor en Linea..", vbInformation, "[SQL " + cDSource2 + "]"
    'Else
    '   MsgBox "Servidor SQL NO responde..", vbCritical
    '   Exit Sub
    'End If
    
    c_ano = Left(c_Periodo, 4)
    c_mes = Right(c_Periodo, 2)
    c_criterio = " ano='" + c_ano + "' and mes='" + c_mes + "' and tarjeta='" + c_Tarjeta + "'"
      '--    MsgBox "1-aplica  " + c_criterio
     
    cHoyDia = "00"
    cMesActual = "N"
    If c_ano + c_mes = Left(Format(Date, "YYYYMMDD"), 6) Then
       '-Fecha desde donde aplica cambios directorio comedor, sujeto no sea propiedad nutricion
       cMesActual = "S"
       cHoyDia = Right(Format(Date, "YYYYMMDD"), 2)
    End If
      '--    MsgBox "2-mes act " + cMesActual + " hoydia " + cHoyDia
         
   Dim OpenForms
   Dim v_an, i, z, k, diai, diaf, nvez As Integer
   Dim lgrabo As Boolean
   v_an = 0
   
   '-OJO- cMesActual="S" -> Hoy -> hasta fin de mes
   diai = IIf(cMesActual = "S", Val(cHoyDia), Day(primer))  'Dias de inicio y cierre del periodo consultado
   diaf = Day(ultimo)     'Day(ultimo)
         '--MsgBox "3-periodo del: " + Trim(Str(diai)) + " al: " + Trim(Str(diaf))
         
   '-Directorio mensual  del SERVIDOR
   sqg = "SELECT * FROM CpDirectorio WHERE " & c_criterio
          '--MsgBox "3.1-busca " + sqg
   Set r_d = New Recordset
   r_d.Open sqg, Cnx, adOpenStatic, adLockReadOnly, adCmdText
   If r_d.EOF Then
      '-check  cuando CREA dir servidor
         '--MsgBox "3.2-No ubico " + sqg, vbCritical, "[ERROR]"
      Exit Sub
   End If
   ''---Do While Not r_d.EOF         '=====> Servidores
        
      '-dar control a windows y permitir multihilo
      v_an = v_an + 1
      If v_an Mod 10 = 0 Then     'if loop has repeated
         OpenForms = DoEvents     'Yield to operating system
      End If
      cColas = "N": cDscCargo = "": wdscUpr = ""
      capelPat = "": capelMat = "": cnombres = "": csexo = ""
      cTarjeta = Trim(r_d!tarjeta)     '- servidor
      c_Costo = IIf(IsNull(r_d!CCosto), "", Trim(r_d!CCosto))   '06.10.2025
      cCodUpr = IIf(IsNull(r_d!CodUpr), "", Trim(r_d!CodUpr))
      cCargo = IIf(IsNull(r_d!Cargo), "", Trim(r_d!Cargo))
      Set rt = New Recordset
      SqC = "SELECT * FROM CpUndPrgRol WHERE codupr='" + cCodUpr + "'"
      rt.Open SqC, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rt.EOF Then
         '----cColas = rt!colas
         wdscUpr = Trim(rt!descripcion)
      End If
      rt.Close
      
      Set rs_d = New Recordset
      sql = "SELECT * FROM CpGeneral WHERE tarjeta='" + cTarjeta + "'"
      rs_d.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
      If Not rs_d.EOF Then
         capelPat = IIf(IsNull(rs_d!apelpat), "", Trim(rs_d!apelpat))
         capelMat = IIf(IsNull(rs_d!apelmat), "", Trim(rs_d!apelmat))
         cnombres = IIf(IsNull(rs_d!nombres), "", Trim(rs_d!nombres))
         cDni = Trim(rs_d!documento)
         csexo = IIf(IsNull(rs_d!sexo), "", Trim(rs_d!sexo))
         
      End If
      rs_d.Close
         '--MsgBox "3.3-" + capelPat + " " + cnombres + "->" + cColas
           
    If cDni <> "" Then  '<====  servidor identificado para comedor
         '--If cColas = "S" Then
         '--   MsgBox "4-" + cnombres + " " + capelPat + ": " + cColas
         '--End If
        
       If cCodUpr <> "" And wdscUpr = "" Then
          Set rs_d = New Recordset
          rs_d.Open "SELECT * FROM CpUndPrgRol WHERE codupr='" & cCodUpr & "'", Cnx, adOpenStatic, adLockReadOnly
          If Not rs_d.EOF Then
             wdscUpr = Trim(rs_d!descripcion)   ' Trim(rs_d!servicioSusalud)
          End If
          rs_d.Close
       End If
       
       If cMesActual = "S" Then
           '-Retira registros editables de Mes ACTUAL  programa-comedor del Servidor
           sql = "DELETE FROM CpDirTrabajadores WHERE ano='" + c_ano + "' AND mes='" + c_mes + "' AND dia>='" + cHoyDia + "' AND tarjeta='" + cTarjeta + "' AND  (Usuarioreg<>'mgomez') AND (left(Usuarioreg,2)<>'PG') AND (left(Usuarioreg,2)<>'NU') "
               '--  MsgBox "retiro mes actual->S: " + sql
           cnx2.Execute (sql)
       Else
           '-Retira programa-comedor del Servidor
               '--MsgBox "No: Delete FROM CpDirTrabajadores  WHERE " + c_criterio
           cnx2.Execute "Delete FROM CpDirTrabajadores  WHERE " + c_criterio
       End If
                  
       For i = diai To diaf      '<<===dias calendario mes analizar, max 31
            z = i - 1
            FechaB = Format(primer + z, "yyyymmdd")
            cTurno = "": cHoraI = "": cHoraF = "":: cHoraI1 = "": cHoraF1 = "": p_horas = ""
            p_ing = "": p_sal = "": p_ing1 = "": p_sal1 = ""
            '-horas programadas
            nHrProg = 0
            
            '-NO APLICA  LACTANCIA
            '==========
            cAplica = "N": nSonLacta = 0
            cIniLac = "": cFinLac = ""
            '-Tolerancia x LACTANCIA (I)ngreso/(S)alida
            '=======================
            nLacta = 0
            
            cTurno = IIf(IsNull(r_d.Fields(z + 4)), "", r_d.Fields(z + 4))
            '-HORARIO prog en dia.
            sql = "SELECT * FROM CpHor" + c_ano + "  WHERE tarjeta='" & cTarjeta & "' AND fecha ='" & FechaB & "'"
            Set rs = New Recordset
            rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
            If rs.EOF Then
               If Trim(cGrpGen) <> "" Then
                  rs.Close
                  sql = "SELECT * FROM CpHor" + c_ano + "  WHERE grupo='" & cGrpGen & "' AND fecha ='" & FechaB & "'"
                  Set rs = New Recordset
                  rs.Open sql, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               End If
            End If
            If Not rs.EOF Then
               cvalor = IIf(IsNull(rs!Horario), "", rs!Horario)
               Sq0 = "SELECT * FROM CpHorarioNew  WHERE Codigo='" & cvalor & "'"
               Set rs_d = New Recordset
               rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               If rs_d.EOF Then
                  rs_d.Close
                  Sq0 = "SELECT * FROM CpHorarios  WHERE Codigo='" & cvalor & "'"
                  Set rs_d = New Recordset
                  rs_d.Open Sq0, Cnx, adOpenStatic, adLockReadOnly, adCmdText
               End If
               If Not rs_d.EOF Then
                  p_ing = IIf(IsNull(rs_d!ingreso), "", rs_d!ingreso)
                  p_sal = IIf(IsNull(rs_d!salida), "", rs_d!salida)
                  p_ing1 = IIf(IsNull(rs_d!Ingreso1), "", rs_d!Ingreso1)
                  p_sal1 = IIf(IsNull(rs_d!salida1), "", rs_d!salida1)
                  nHrProg = Val(Left(rs_d!Horas, 2)) + Val(Right(rs_d!Horas, 2)) / 60
                  p_horas = IIf(IsNull(rs_d!Horas), "0000", rs_d!Horas)
               End If
               rs_d.Close
            End If
            rs.Close
            
            '-Actualiza horario del Servidor
            c_DIA = Right("0" + Trim(Str(i)), 2)
            If nHrProg >= 12 Or (cCodUpr = "23" And cTurno = "MX") Then '
              '-Aplica 26.01.2026 incluye PASANTES Solo Almuerzo (cCodUpr = "23" And Trim(cTurno) = "MX")
                '''If cCodUpr = "23" And cTurno = "MX" Then
                '''    MsgBox "pasante " + c_DIA + " con turno: " + cTurno
                '''End If
               
               nvez = 1
               cHoraI = Left(p_ing, 2) + ":" + Right(p_ing, 2)
               cHoraF = Left(p_sal, 2) + ":" + Right(p_sal, 2)
               If Val(p_ing1) > 0 Then
                  cHoraI1 = Left(p_ing1, 2) + ":" + Right(p_ing1, 2)
                  cHoraF1 = Left(p_sal1, 2) + ":" + Right(p_sal1, 2)
                  nvez = 2
               End If
               CriterioDia = c_criterio + " AND dia='" + c_DIA + "'"
                  'MsgBox "6-aplica nutri  " + CriterioDia
                  
               '-Aplica programa-COMEDOR
               For k = 1 To nvez
                  lgrabo = False
                   ' If c_DIA = "9" Then
                   ' MsgBox c_DIA + " vez:" + Trim(Str(k))
                   ' End If
                  
                  If k = 1 Then
                     nIdHORARIO = 0
                     If Val(Left(cHoraI, 2)) < 12 Then
                        nIdHORARIO = 1
                     Else
                        nIdHORARIO = 2
                     End If
                     '''   If cCodUpr = "23" And cTurno = "MX" Then
                     '''     MsgBox c_DIA + " pasante, vez:" + Trim(Str(k)) + "->" + Trim(Str(nIdHORARIO))
                     '''   End If
                        
                     nvalor = Val(Left(cHoraF, 2)) - Val(Left(cHoraI, 2))
                     If nvalor > 0 Then
                          ' If c_DIA = "9" Then
                          ' MsgBox "vez:" + Trim(Str(k)) + "=>" + cHoraI + "-" + cHoraF + "; hrs:" + Trim(Str(nvalor))
                          ' End If
                        
                        If nvalor >= 12 Or (cCodUpr = "23" And cTurno = "MX") Then    '¿PASANTES solo ALMUERZO ==> (cCodUpr = "23" And Trim(cTurno) = "MX")
                           If cMesActual = "N" Or (cMesActual = "S" And c_DIA >= cHoyDia) Then
                              Set rsr = New Recordset
                              sqr = "SELECT * FROM CpDirTrabajadores  WHERE " & CriterioDia & " AND IdHorario=" & CStr(nIdHORARIO)
                              rsr.Open sqr, cnx2, adOpenStatic, adLockOptimistic, adCmdText
                              If rsr.EOF() Then         'restablece dia programa-comedor
                                rsr.AddNew
                                rsr!ano = c_ano
                                rsr!mes = c_mes
                                rsr!Dia = c_DIA
                                rsr!Dni = cDni
                                rsr!apellidos = Trim(capelPat) + " " + Trim(capelMat)
                                rsr!nombres = Trim(cnombres)
                                rsr!tarjeta = cTarjeta
                                rsr!CCosto = Trim(c_Costo)     '06.10.2025
                                rsr!Servicio = Trim(wdscUpr)
                                rsr!fechareg = r_d!fecha
                                rsr!horareg = Format(Time, "HH:mm:ss")
                                rsr!usuarioReg = r_d!usuario
                                rsr!estacionREg = r_d!estacion
                                rsr!entrada = cHoraI
                                rsr!salida = cHoraF
                                rsr!IdHorario = nIdHORARIO
                                lgrabo = True
                                    'nSonLacta = nSonLacta + 1
                                '''    If cCodUpr = "23" And cTurno = "MX" Then
                                '''       MsgBox "Creo1 pasante " + cTarjeta + ": " + c_ano + c_mes + c_DIA + "=>" + Trim(Str(nIdHORARIO))
                                '''    End If
                                 
                              End If
                           End If
                        End If
                     Else
                          'If c_DIA = "27" Then
                          'MsgBox c_DIA + " vez:" + Trim(Str(k)) + "->" + Trim(Str(nIdHORARIO))
                          'End If
                   
                        nvalor = nvalor + 24
                        If nvalor >= 12 Then
                           If cMesActual = "N" Or (cMesActual = "S" And c_DIA >= cHoyDia) Then
                            Set rsr = New Recordset
                              sqr = "SELECT * FROM CpDirTrabajadores  WHERE " & CriterioDia & " AND IdHorario=" & CStr(nIdHORARIO)
                              rsr.Open sqr, cnx2, adOpenStatic, adLockOptimistic, adCmdText
                              If rsr.EOF() Then         'restablece dia programa-comedor
                                rsr.AddNew
                                rsr!ano = c_ano
                                rsr!mes = c_mes
                                rsr!Dia = c_DIA
                                rsr!Dni = cDni
                                rsr!apellidos = Trim(capelPat) + " " + Trim(capelMat)
                                rsr!nombres = Trim(cnombres)
                                rsr!tarjeta = cTarjeta
                                rsr!CCosto = Trim(c_Costo)     '06.10.2025
                                rsr!Servicio = Trim(wdscUpr)
                                rsr!fechareg = r_d!fecha
                                rsr!horareg = Format(Time, "HH:mm:ss")
                                rsr!usuarioReg = r_d!usuario
                                rsr!estacionREg = r_d!estacion
                                rsr!entrada = cHoraI
                                rsr!salida = cHoraF
                                rsr!IdHorario = nIdHORARIO
                                lgrabo = True
                                    'nSonLacta = nSonLacta + 1
                                    'If c_DIA = "27" Then
                                    '   MsgBox "Creo2 " + c_ano + c_mes + c_DIA + "=>" + Trim(Str(nIdHORARIO))
                                    'End If
                                 
                              End If
                           End If
                        End If
                     End If
                  Else
                     nvalor = Val(Left(cHoraF1, 2)) - Val(Left(cHoraI1, 2)) + 24
                     nIdHORARIO = 0
                     If Val(Left(cHoraI1, 2)) < 12 Then
                        nIdHORARIO = 1
                     Else
                        nIdHORARIO = 2
                     End If
                           'If c_DIA = "27" Then
                           'MsgBox "vez:" + Trim(Str(k)) + "=>" + cHoraI1 + "-" + cHoraF1 + "; hrs:" + Trim(Str(nvalor))
                           'End If
                     
                     If nvalor >= 12 Then
                        If cMesActual = "N" Or (cMesActual = "S" And c_DIA >= cHoyDia) Then
                           Set rsr = New Recordset
                           sqr = "SELECT * FROM CpDirTrabajadores  WHERE " & CriterioDia & " AND IdHorario=" & CStr(nIdHORARIO)
                           rsr.Open sqr, cnx2, adOpenStatic, adLockOptimistic, adCmdText
                           If rsr.EOF() Then         'restablece dia programa-servidor
                              rsr.AddNew
                              rsr!ano = c_ano
                              rsr!mes = c_mes
                              rsr!Dia = c_DIA
                              rsr!Dni = cDni
                              rsr!apellidos = Trim(capelPat) + " " + Trim(capelMat)
                              rsr!nombres = Trim(cnombres)
                              rsr!tarjeta = cTarjeta
                              rsr!CCosto = Trim(c_Costo)     '06.10.2025
                              rsr!Servicio = Trim(wdscUpr)
                              rsr!fechareg = r_d!fecha
                              rsr!horareg = Format(Time, "HH:mm:ss")
                              rsr!usuarioReg = r_d!usuario
                              rsr!estacionREg = r_d!estacion
                              rsr!entrada = cHoraI1
                              rsr!salida = cHoraF1
                              rsr!IdHorario = nIdHORARIO
                              lgrabo = True
                                 'nSonLacta = nSonLacta + 1
                                 'If c_DIA >= "24" Then
                                 '   MsgBox "Creo3 " + c_ano + c_mes + c_DIA + "=>" + Trim(Str(nIdHORARIO))
                                 'End If
                                 
                           End If
                        End If
                     End If
                  End If
                  If lgrabo Then
                     rsr.Update
                     rsr.Close
                  End If
               Next k
                  'If c_DIA = "27" Then
                  'MsgBox "GRABO..." + Str(nSonLacta)
                  'End If
               
            End If
            
       Next i            '<<=== sgt Dia calendario mes analizar, max 31
       
       '-UPDATE caso 08:00 y IdHorario=2
       sqr = "UPDATE CpDirTrabajadores SET Idhorario=1 WHERE " & c_criterio & " AND ENTRADA='08:00' AND IdHorario<>1"
       cnx2.Execute (sqr)
       
     End If   '====>  fin servidor identificado
   ''---  r_d.MoveNext     'Sgt Servidor
   ''---Loop
   r_d.Close
   cnx2.Close
   Set cnx2 = Nothing
   Beep 700, 300
   Beep 700, 300
   Exit Sub
ErrProc:
   MsgBox "ERROR " & Err.Number & vbCrLf & Err.Description, vbCritical, "ProcProg3RO"
   Resume Next
End Sub


