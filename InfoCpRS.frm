VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form InfoCpRS 
   Caption         =   "Informes"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12465
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "InfoCpRS.frx":0000
   ScaleHeight     =   9090
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows Default
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   8415
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   11295
      lastProp        =   600
      _cx             =   19923
      _cy             =   14843
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "InfoCpRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private crxApp As New CRAXDRT.Application
Private crxRpt As CRAXDRT.Report
Dim crxTables As CRAXDRT.DatabaseTables
Dim crxTable As CRAXDRT.DatabaseTable
Dim strServerOrDSNName As String
Dim strDBNameOrPath As String
Dim strUserID As String
Dim strPassword As String
Dim strSelectionfrm As String
Dim Rst As ADODB.Recordset
Dim cOrdenado As String
'----Set de datos  es Recordset NO SQL query

Private Sub Form_Load()
  Dim crxParamDefs As CRAXDRT.ParameterFieldDefinitions
  Dim crxParamDef As CRAXDRT.ParameterFieldDefinition
  On Error GoTo ErrHandler
  strServerOrDSNName = cDSource    '"192.168.1.19,1433" replace servername with your SQL connection server
  strDBNameOrPath = cICatal        '"hepasistencial" replace databasename with your SQL connection databasename
  strUserID = cUsr                 '"sa"  replace userid with your SQL connection userid
  strPassword = cPwd               '"hep" replace password with your SQL connection password
'-Valor asignar en Parametros
  wimpreso = wUsuNombre
  wTitulo = "PROGRAMACIÓN MENSUAL DE TURNOS Y HORARIOS DEL SERVICIO ASISTENCIAL" + IIf(ProgASISTENCIAL.TxRegPlh.Text = "3RO", " - TERCEROS", "")
           
  If nTipRpt = 4.3 Then
         wdetalle = "SUMARIO DE FALTAS y TARDANZAS - periodo del: " & Format(TardaInasis.Txfecini, "dd/mm/yyyy") + " al: " & Format(TardaInasis.TxFecFin, "dd/mm/yyyy")
  
  ElseIf nTipRpt = 5 Then
         wdetalle = "Reporte de Asistencia del " & RecMarcasRS.Txpridia & "  Al " & RecMarcasRS.TxUltDia
         wnombrep = "(CpRecordRS)"
         Wonomastico = RecMarcasRS.LblONOMASTICO.Caption
         Wtrabajador = RecMarcasRS.Txtrabajador.Text
         wtHr_Prog = "": whr_tard = "": whr_tm130 = "": wtHr_Fi = "": wtHr_Mp = "": wtHr_trab = ""
         wtHr_Prog = RecMarcasRS.TxhorasP.Text
         wm_tard = RecMarcasRS.TxmTarda.Text
         whr_tard = cHorsMint(Val(RecMarcasRS.TxmTarda.Text) / 60)
         wm_tm130 = RecMarcasRS.TxmDscto.Text
         whr_tm130 = cHorsMint(Val(RecMarcasRS.TxmDscto.Text) / 60)
         wtHr_Mp = "" + RecMarcasRS.Txhorasmp.Text
         wtHr_FMp = "" + RecMarcasRS.Txhorasfmp.Text
         wtHr_Fj = "" + RecMarcasRS.TxhorasFj.Text
         wtHr_Fi = "" + RecMarcasRS.TxhorasFI.Text
         wtHr_trab = "" + RecMarcasRS.TxHorast.Text
         
  ElseIf nTipRpt = 5.5 Then
         If wForma = "RECMARCASQUERY" Then
            wdetalle = "Reporte Sumario Asistencia del " & RecMarcasQuery.Txpridia & "  Al " & RecMarcasQuery.TxUltDia
         ElseIf wForma = "RECPERTARJETA" Then
            wdetalle = "Reporte Sumario Asistencia del " & RecPerTARJETA.Txpridia & "  Al " & RecPerTARJETA.TxUltDia
         End If
         wnombrep = "(CpRecMasRS)"
         
  ElseIf nTipRpt = 5.7 Then
         wdetalle = "Reporte mensual de asistencia del " & RecPerTARJETA.Txpridia & "  Al " & RecPerTARJETA.TxUltDia
         '--   wdetalle = "Reporte mensual de asistencia del " & RecPerTARJETA.Txpridia & "  Al " & RecPerTARJETA.TxUltDia
         wnombrep = "(CpRecPerTarjeta)"
  ElseIf nTipRpt = 5.8 Then
         wdetalle = "Reporte mensual de asistencia del " & RecMarcasQuery.Txpridia & "  Al " & RecMarcasQuery.TxUltDia
        
  ElseIf nTipRpt = 7 Then
         wdetalle = "Directorio " + IIf(RptDirectorio.Check2.Value = 1, " - TERCEROS", "") + " del " & RptDirectorio.TxDia.Text
         wnombrep = "(CpDirectorio)"
         
  ElseIf nTipRpt = 7.5 Then
         wdetalle = "Servidores NO PROGRAMADOS el " & RptDirectorio.TxDia.Text
         wnombrep = "(CpDirectorio)"
         
  ElseIf nTipRpt = 9 Then
            wdetalle = "Analisis Cumplimiento del: " & CPcumplimientoProg.Txfecini.Text & "  Al: " & CPcumplimientoProg.TxFecFin.Text
            wdetalle1 = "Horas Programadas: " + CPcumplimientoProg.TxHrsP.Text
            wcampo0 = CPcumplimientoProg.TxhrsT.Text
            wcampo1 = Format(CPcumplimientoProg.TxPorcEjecuta.Text, "##.#0")
         wnombrep = "(CpCumplimientoProg)"
         
  ElseIf nTipRpt = 11 Then
         cOrdenado = IIf(ProgASISTENCIAL.Ordenado(0).Value = True, "A", "C")
         If ProgASISTENCIAL.TxRegPlh.Text = "RES" Then
            wdetalle = "Correspondiente al Mes de " & ProgASISTENCIAL.TxNombMes.Text & " " & genaño + IIf(InStr("23_91", ProgASISTENCIAL.Tx_codupr.Text) > 0, IIf(ProgASISTENCIAL.Tx_codupr.Text = "23", " - RESIDENTES", " - INTERNOS"), "")
         Else
            wdetalle = "Correspondiente al Mes de " & ProgASISTENCIAL.TxNombMes.Text & " " & genaño + IIf(ProgASISTENCIAL.Tx_codupr.Text = "23", " - RESIDENTES", "")
         End If
         wnombrep = "(CpDirCosto)"
         W_upr = ProgASISTENCIAL.TxdscUpr.Text
  ElseIf nTipRpt = 11.1 Then
         cOrdenado = IIf(ProgASISTENCIAL.Ordenado(0).Value = True, "A", "C")
         wdetalle = "Correspondiente al Mes de " & ProgASISTENCIAL.TxNombMes.Text & " " & genaño
         wnombrep = "(CpDirCosto11)"
         W_upr = ProgASISTENCIAL.TxdscUpr.Text
  ElseIf nTipRpt = 112 Then
         wdetalle = "HORAS COMPLEMENTARIAS del Mes de " & ProgComplementarios.TxNombMes.Text & " " & genaño
         wnombrep = "(CpDirHC)"
         W_upr = ProgComplementarios.TxdscUpr.Text
  ElseIf nTipRpt = 112.1 Then
         wdetalle = "HORAS COMPLEMENTARIAS del Mes de " & ProgComplementarios.TxNombMes.Text & " " & genaño
         wnombrep = "(CpDirHC11)"
         W_upr = ProgComplementarios.TxdscUpr.Text
         
  ElseIf nTipRpt = 11.2 Then
         wdetalle = "Correspondiente al Mes de " & ProgAdministrativo.TxNombMes.Text & " " & genaño
         wnombrep = "(CpDirCostoAdm)"
         W_ccosto = ProgAdministrativo.Txdscccosto.Text
  ElseIf nTipRpt = 11.3 Then
         wdetalle = "Correspondiente al Mes de " & ProgEspecial.TxNombMes.Text & " " & genaño
         W_ccosto = ProgEspecial.Txdscccosto.Text
         wnombrep = "(CpDirectCosto)"
  ElseIf nTipRpt = 1011 Then
         wdetalle = "Cambios en Directorio - " + ProgASISTENCIAL.TxdscUpr.Text
         wdetalle1 = "Correspondiente al Mes de " & ProgASISTENCIAL.TxNombMes.Text & " " & genaño
         W_upr = ProgASISTENCIAL.Tx_codupr.Text
  ElseIf nTipRpt = 1011.2 Then
         wdetalle = "Cambios en Directorio - " + ProgComplementarios.TxdscUpr.Text
         wdetalle1 = "HORAS COMPLEMENTARIAS del Mes de " & ProgComplementarios.TxNombMes.Text & " " & genaño
         W_upr = ProgComplementarios.Tx_codupr.Text
  ElseIf nTipRpt = 1011.3 Then
         wdetalle = "Cambios en Directorio - " + ProgEspecial.Txdscccosto.Text
         wdetalle1 = "Correspondiente al Mes de " & ProgEspecial.TxNombMes.Text & " " & genaño
         W_upr = ProgEspecial.Tx_ccosto.Text
  End If

'-Open the report
  Set crxApp = New CRAXDRT.Application
    
  If nTipRpt = 4.3 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpSumarioFaltTard.rpt", 1)
      
  ElseIf nTipRpt = 5 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpRecordRS.rpt", 1)
         
  ElseIf nTipRpt = 5.5 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpRecMasRS.rpt", 1)
  ElseIf nTipRpt = 5.7 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpRecPerTarjeta.rpt", 1)
  ElseIf nTipRpt = 5.8 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpRecUprTarjeta.rpt", 1)
 
  ElseIf nTipRpt = 7 Or nTipRpt = 7.5 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpDirectorio.rpt", 1)
  ElseIf nTipRpt = 9 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpCumplimientoProg.rpt", 1)
         
  ElseIf nTipRpt = 11 Or nTipRpt = 112 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpDirCosto.rpt", 1)
  ElseIf nTipRpt = 11.1 Or nTipRpt = 112.1 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpDirCosto11.rpt", 1)
  ElseIf nTipRpt = 11.2 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpDirCostoAdm.rpt", 1)
  ElseIf nTipRpt = 11.3 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpDirectCosto.rpt", 1)
  ElseIf nTipRpt = 1011 Or nTipRpt = 1011.2 Or nTipRpt = 1011.3 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpAudDirect.rpt", 1)

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
If nTipRpt = 4.3 Then
         sql = " SELECT DISTINCT " & _
 "#tablaST.dni,#tablaST.servidor,#tablaST.cargo,#tablaST.falta1,#tablaST.tarda1,#tablaST.falta2,#tablaST.tarda2,#tablaST.falta3,#tablaST.tarda3,#tablaST.falta4,#tablaST.tarda4,#tablaST.falta5, " & _
 "#tablaST.tarda5,#tablaST.falta6,#tablaST.tarda6,#tablaST.falta7,#tablaST.tarda7,#tablaST.falta8,#tablaST.tarda8,#tablaST.falta9,#tablaST.tarda9,#tablaST.falta10,#tablaST.tarda10,#tablaST.falta11, " & _
 "#tablaST.tarda11,#tablaST.falta12,#tablaST.tarda12,#tablaST.TOTFaltas,#tablaST.TOTTarda FROM  #tablaST order by servidor "
         

ElseIf nTipRpt = 5 Then
         'CpTmRecord     field(grupo)->"LACT"
         sql = "SELECT DISTINCT " & _
         "#tabla1.DscPlh, #tabla1.NombDia, #tabla1.fecha, #tabla1.turno, #tabla1.ming, #tabla1.marca2, #tabla1.marca3, #tabla1.marca4, #tabla1.hrProg, #tabla1.hrTrab, #tabla1.tardanza, #tabla1.Descta, #tabla1.Autorizac, " & _
         "#tabla1.DatoMp, #tabla1.hrini, #tabla1.hrfin, #tabla1.Comision, #tabla1.motpersonal, #tabla1.tarjeta, #tabla1.PLH, #tabla1.dscactividad, #tabla1.dscCCosto, #tabla1.dscCarEst, #tabla1.hrasTrab, #tabla1.GRUPO " & _
         "FROM   #tabla1 WHERE #tabla1.tarjeta='" & RecMarcasRS.Txtarjeta & "'  ORDER BY #tabla1.fecha"
         '-Replace the value of strSelectionfrm variable with your selection formula
         'strSelectionfrm = "{tablatmp.orden}='" & QMultiple.Txorden.Text & "'"
  
  ElseIf nTipRpt = 5.5 Then
          'CpTmRecord -> CpTmGeneral
          sql = "SELECT DISTINCT " & _
          "#tabla1.tarjeta, #tabla1.fecha, #tabla1.NombDia, #tabla1.turno, #tabla1.ming, #tabla1.marca2, #tabla1.marca3, #tabla1.marca4, #tabla1.tardanza, #tabla1.Descta, #tabla1.hrProg, #tabla1.hrTrab, " & _
          "#tabla1.Autorizac, #tabla1.DatoMp, #tabla1.hrini, #tabla1.hrfin, #tabla1.Comision, #tabla1.GRUPO, " & _
          "CpTmGeneral.trabajador, CpTmGeneral.dscactividad, CpTmGeneral.dscccosto, CpTmGeneral.dsccargo, CpTmGeneral.horasP, CpTmGeneral.htarda, CpTmGeneral.hdscto, CpTmGeneral.horasmP, CpTmGeneral.horasFJ, " & _
          "CpTmGeneral.horasFI,CpTmGeneral.horasFmp, CpTmGeneral.horast " & _
          "FROM  #tabla1 INNER JOIN CptmGeneral ON #tabla1.tarjeta=CptmGeneral.tarjeta " & _
          "ORDER BY #tabla1.tarjeta,#tabla1.fecha"
  
  ElseIf nTipRpt = 5.7 Then
          'SERVIDOR-PERIODO (mensual)  CpTmRecord -> CpTmGeneral
          sql = "SELECT DISTINCT " & _
          "#tabla1.tarjeta,#tabla1.MES, #tabla1.fecha, #tabla1.NombDia, #tabla1.turno, #tabla1.ming, #tabla1.marca2, #tabla1.marca3, #tabla1.marca4, #tabla1.tardanza, #tabla1.Descta, #tabla1.hrProg, #tabla1.hrTrab, " & _
          "#tabla1.Autorizac, #tabla1.DatoMp, #tabla1.hrini, #tabla1.hrfin, #tabla1.Comision, #tabla1.GRUPO, " & _
          "CpTmGeneral.trabajador, CpTmGeneral.dscactividad, CpTmGeneral.dscccosto, CpTmGeneral.dsccargo, CpTmGeneral.horasP, CpTmGeneral.htarda, CpTmGeneral.hdscto, CpTmGeneral.horasmP, CpTmGeneral.horasFJ, " & _
          "CpTmGeneral.horasFI,CpTmGeneral.horasFmp, CpTmGeneral.horast " & _
          "FROM  #tabla1 INNER JOIN CptmGeneral ON #tabla1.Mes=CptmGeneral.Mes " & _
          "ORDER BY #tabla1.mes,#tabla1.fecha"
  ElseIf nTipRpt = 5.8 Then
          'Oficina-servidor-PERIODO   CpTmRecord -> CpTmGeneral
          sql = "SELECT DISTINCT " & _
          "#tabla1.tarjeta,#tabla1.MES,#tabla1.tarjetaMes, #tabla1.fecha, #tabla1.NombDia, #tabla1.turno, #tabla1.ming, #tabla1.marca2, #tabla1.marca3, #tabla1.marca4, #tabla1.tardanza, #tabla1.Descta, #tabla1.hrProg, #tabla1.hrTrab, " & _
          "#tabla1.Autorizac, #tabla1.DatoMp, #tabla1.hrini, #tabla1.hrfin, #tabla1.Comision, #tabla1.GRUPO, " & _
          "CpTmGeneral.trabajador, CpTmGeneral.dscactividad, CpTmGeneral.dscccosto, CpTmGeneral.dsccargo, CpTmGeneral.horasP, CpTmGeneral.htarda, CpTmGeneral.hdscto, CpTmGeneral.horasmP, CpTmGeneral.horasFJ, " & _
          "CpTmGeneral.horasFI,CpTmGeneral.horasFmp, CpTmGeneral.horast " & _
          "FROM  #tabla1 INNER JOIN CptmGeneral ON #tabla1.tarjetaMes=CptmGeneral.tarjetaMES  " & _
          "ORDER BY #tabla1.tarjetaMES,#tabla1.fecha"
          
  ElseIf nTipRpt = 7 Or nTipRpt = 7.5 Then
         sql = "SELECT DISTINCT " & _
         "#tabla2.BLOQUE, #tabla2.Orden, #tabla2.CAMPO1, #tabla2.CAMPO2, #tabla2.CAMPO3, #tabla2.CAMPO4, #tabla2.CAMPO5, #tabla2.CAMPO6, " & _
         "#tabla2.DSCbLOQUE, #tabla2.DSCCARGO, #tabla2.CAMPO7 , #tabla2.CAMPO8 , #tabla2.CAMPO9 " & _
         "FROM  #tabla2 " & _
         "ORDER BY #tabla2.BLOQUE, #tabla2.Orden"
         
  ElseIf nTipRpt = 9 Then
         sql = "SELECT DISTINCT " & _
         "#tabla9.fecha, #tabla9.Dni, #tabla9.CCosto, #tabla9.DscCCosto, #tabla9.CodUpr, #tabla9.DscUpr, #tabla9.tarjeta, CpGeneral.ApelMAT, CpGeneral.ApelPAT, CpGeneral.NOMBRES, " & _
         "#tabla9.cargo, #tabla9.Dsccargo,#tabla9.Plh,#tabla9.categoria,#tabla9.Dsccategoria,  #tabla9.HrasProg, #tabla9.hrasfalt, #tabla9.hrasTrab " & _
         "FROM  #tabla9 INNER JOIN CpGeneral ON #tabla9.tarjeta=CpGeneral.Tarjeta " & _
         "ORDER BY #tabla9.fecha, CpGeneral.ApelPAT, CpGeneral.ApelMAT, CpGeneral.NOMBRES"
         
  ElseIf nTipRpt = 11 Or nTipRpt = 112 Then
         sql = "SELECT DISTINCT " & _
         "#Tabla.d1, #Tabla.d2, #Tabla.d3, #Tabla.d4, #Tabla.d5, #Tabla.d6, #Tabla.d7, #Tabla.d8, #Tabla.d9, #Tabla.d10, #Tabla.d11, #Tabla.d12, " & _
         "#Tabla.d13, #Tabla.d14, #Tabla.d15, #Tabla.d16, #Tabla.d17, #Tabla.d18, #Tabla.d19, #Tabla.d20, #Tabla.d21, #Tabla.d22, #Tabla.d23, #Tabla.d24, " & _
         "#Tabla.d25, #Tabla.d26, #Tabla.d27, #Tabla.d28, #Tabla.d29, #Tabla.d30, #Tabla.d31, #Tabla.n1, #Tabla.n2, #Tabla.n3, #Tabla.n4, #Tabla.n5, " & _
         "#Tabla.n6, #Tabla.n7, #Tabla.n8, #Tabla.n9, #Tabla.n10, #Tabla.n11, #Tabla.n12, #Tabla.n13, #Tabla.n14, #Tabla.n15, #Tabla.n16, #Tabla.n17, " & _
         "#Tabla.n18, #Tabla.n19, #Tabla.n20, #Tabla.n21, #Tabla.n22, #Tabla.n23, #Tabla.n24, #Tabla.n25, #Tabla.n26, #Tabla.n27, #Tabla.n28, #Tabla.n29, " & _
         "#Tabla.n30, #Tabla.n31, #Tabla.d, #Tabla.n, #Tabla.M, #Tabla.T, #Tabla.MT, #Tabla.NA, #Tabla.PA, #Tabla.horas, #Tabla.dscCargo , #Tabla.nombre " & _
         "FROM   #Tabla  ORDER BY " + IIf(cOrdenado = "A", " #Tabla.nombre", " #Tabla.dscCargo, #Tabla.nombre")
  ElseIf nTipRpt = 11.1 Or nTipRpt = 112.1 Then
         sql = "SELECT DISTINCT " & _
         "#Tabla.d1, #Tabla.d2, #Tabla.d3, #Tabla.d4, #Tabla.d5, #Tabla.d6, #Tabla.d7, #Tabla.d8, #Tabla.d9, #Tabla.d10, #Tabla.d11, #Tabla.d12, " & _
         "#Tabla.d13, #Tabla.d14, #Tabla.d15, #Tabla.d16, #Tabla.d17, #Tabla.d18, #Tabla.d19, #Tabla.d20, #Tabla.d21, #Tabla.d22, #Tabla.d23, #Tabla.d24 ," & _
         "#Tabla.d25, #Tabla.d26, #Tabla.d27, #Tabla.d28, #Tabla.d29, #Tabla.d30, #Tabla.d31, #Tabla.n1, #Tabla.n2, #Tabla.n3, #Tabla.n4, #Tabla.n5 , " & _
         "#Tabla.n6, #Tabla.n7, #Tabla.n8, #Tabla.n9, #Tabla.n10, #Tabla.n11, #Tabla.n12, #Tabla.n13, #Tabla.n14, #Tabla.n15, #Tabla.n16, #Tabla.n17, " & _
         "#Tabla.n18, #Tabla.n19, #Tabla.n20, #Tabla.n21, #Tabla.n22, #Tabla.n23, #Tabla.n24, #Tabla.n25, #Tabla.n26, #Tabla.n27, #Tabla.n28, #Tabla.n29, " & _
         "#Tabla.n30, #Tabla.n31, #Tabla.d, #Tabla.n, #Tabla.M, #Tabla.T, #Tabla.MT, #Tabla.NA, #Tabla.PA, #Tabla.horas, #Tabla.dsccargo, #Tabla.NOMBRE " & _
         "FROM   #Tabla  ORDER BY " + IIf(cOrdenado = "A", " #Tabla.nombre", " #Tabla.dscCargo, #Tabla.nombre")
  ElseIf nTipRpt = 11.2 Then
         sql = "SELECT DISTINCT " & _
         "#Tabla.d1, #Tabla.d2, #Tabla.d3, #Tabla.d4, #Tabla.d5, #Tabla.d6, #Tabla.d7, #Tabla.d8, #Tabla.d9, #Tabla.d10, #Tabla.d11, " & _
         "#Tabla.d12, #Tabla.d13, #Tabla.d14, #Tabla.d15, #Tabla.d16, #Tabla.d17, #Tabla.d18, #Tabla.d19, #Tabla.d20, #Tabla.d21, #Tabla.d22, " & _
         "#Tabla.d23, #Tabla.d24, #Tabla.d25, #Tabla.d26, #Tabla.d27, #Tabla.d28, #Tabla.d29, #Tabla.d30, #Tabla.d31, #Tabla.horas, #Tabla.nombre, " & _
         "#Tabla.ccosto, #Tabla.tarjeta, #Tabla.d, #Tabla.n, #Tabla.M, #Tabla.T, #Tabla.MT, #Tabla.n1, #Tabla.n2, #Tabla.n3, #Tabla.n4, #Tabla.n5, #Tabla.n6, " & _
         "#Tabla.n7, #Tabla.n8, #Tabla.n9, #Tabla.n10, #Tabla.n11, #Tabla.n12, #Tabla.n13, #Tabla.n14, #Tabla.n15, #Tabla.n16, #Tabla.n17, #Tabla.n18, " & _
         "#Tabla.n19, #Tabla.n20, #Tabla.n21, #Tabla.n22, #Tabla.n23, #Tabla.n24, #Tabla.n25, #Tabla.n26, #Tabla.n27, #Tabla.n28, #Tabla.n29, #Tabla.n30, " & _
         "#Tabla.n31, #Tabla.d, #Tabla.n, #Tabla.M, #Tabla.T, #Tabla.MT, #Tabla.NA, #Tabla.PA, #Tabla.horas " & _
         "FROM  #Tabla   ORDER BY #Tabla.nombre"
  ElseIf nTipRpt = 11.3 Then
         sql = "SELECT DISTINCT " & _
         "#Tabla.d1, #Tabla.d2, #Tabla.d3, #Tabla.d4, #Tabla.d5, #Tabla.d6, #Tabla.d7, #Tabla.d8, #Tabla.d9, #Tabla.d10, #Tabla.d11, #Tabla.d12, " & _
         "#Tabla.d13, #Tabla.d14, #Tabla.d15, #Tabla.d16, #Tabla.d17, #Tabla.d18, #Tabla.d19, #Tabla.d20, #Tabla.d21, #Tabla.d22, #Tabla.d23, #Tabla.d24 ," & _
         "#Tabla.d25, #Tabla.d26, #Tabla.d27, #Tabla.d28, #Tabla.d29, #Tabla.d30, #Tabla.d31, #Tabla.n1, #Tabla.n2, #Tabla.n3, #Tabla.n4, #Tabla.n5 , " & _
         "#Tabla.n6, #Tabla.n7, #Tabla.n8, #Tabla.n9, #Tabla.n10, #Tabla.n11, #Tabla.n12, #Tabla.n13, #Tabla.n14, #Tabla.n15, #Tabla.n16, #Tabla.n17, " & _
         "#Tabla.n18, #Tabla.n19, #Tabla.n20, #Tabla.n21, #Tabla.n22, #Tabla.n23, #Tabla.n24, #Tabla.n25, #Tabla.n26, #Tabla.n27, #Tabla.n28, #Tabla.n29, " & _
         "#Tabla.n30, #Tabla.n31, #Tabla.d, #Tabla.n, #Tabla.M, #Tabla.T, #Tabla.MT,#Tabla.NA, #Tabla.PA, #Tabla.horas, #Tabla.dscCargo , #Tabla.nombre " & _
         "FROM  #Tabla   ORDER BY  #Tabla.nombre"
  ElseIf nTipRpt = 1011 Or nTipRpt = 1011.2 Then
         sql = "SELECT DISTINCT " & _
         "CpAudDirect.fecha, CpAudDirect.tarjeta, CpAudDirect.deTurno, CpAudDirect.aTurno, CpAudDirect.detalle, CpAudDirect.registro, CpAudDirect.estacion, CpAudDirect.fecregis, CpAudDirect.horregis, CpGeneral.apelMAT, CpGeneral.apelPAT, CpGeneral.NOMBREs " & _
         "FROM   CpAudDirect INNER JOIN CpGeneral ON CpAudDirect.tarjeta=CpGeneral.Tarjeta " & _
         "WHERE  CpAudDirect.ano='" + genaño + "' and CpAudDirect.mes='" + ProgASISTENCIAL.Txmes.Text + "' and CpAudDirect.Codupr='" + W_upr + "'  ORDER BY  CpAudDirect.Fecregis,CpAudDirect.horregis,CpAudDirect.Tarjeta"
           
  ElseIf nTipRpt = 1011.3 Then
          sql = "SELECT DISTINCT " & _
          "CpAudDirect.fecha, CpAudDirect.tarjeta, CpAudDirect.deTurno, CpAudDirect.aTurno, CpAudDirect.detalle, CpAudDirect.registro, CpAudDirect.estacion, CpAudDirect.fecregis, CpAudDirect.horregis, CpGeneral.apelMAT, CpGeneral.apelPAT, CpGeneral.NOMBREs " & _
          "FROM   CpAudDirect INNER JOIN CpGeneral ON CpAudDirect.tarjeta=CpGeneral.Tarjeta " & _
          "WHERE  CpAudDirect.ano='" + genaño + "' and CpAudDirect.mes='" + ProgEspecial.Txmes.Text + "' and CpAudDirect.cCosto='" + W_upr + "'  ORDER BY  CpAudDirect.Fecregis,CpAudDirect.horregis,CpAudDirect.Tarjeta"
           
  End If
    
     '-Define Scope
     '- Set rst = Cnx.Execute(SQL)
     Set Rst = New Recordset
     Rst.Open sql, Cnx, adOpenStatic, adLockReadOnly
      
  '+++++++++++++++++++++
  'Query = .....;
  'adapter.fill(datasetname,"table name of the dataset");
  'rpt.setdatasource(datasetname);
  'rpviewer.reportsource = rpt;
  'rpviewer.refresh();
  '-La idea es agregar esta tabla temporar a algun dataset y luego
  '-invocar el metodo setdatasource() para pasar el dataset al conjunto de datos del informes.
  '++++++++++++++++++++++
  
'-Record set as DATASOURCE of Crystal Report
'Dim cnn As ADODB.Connection
'Dim rst As ADODB.Recordset
'Set cnn = New ADODB.Connection
'cnn.Open "File name=c:Product.udl"
'Set rst = cnn.Execute("select item_no,item_category,Item_desc,Item_unit,item_price from tblitems where item_Category ='" & dcCAtegory.Text & "'")
'Dim crystal As CRAXDRT.Application
'Dim treport As CRAXDRT.Report
'CRViewer1.DisplayBorder = False
'CRViewer1.DisplayTabs = False
'CRViewer1.EnableDrillDown = False
'CRViewer1.EnableRefreshButton = False
'Dim crxApplication As New CRAXDRT.Application
'Set crystal = New CRAXDRT.Application
'Set treport = crystal.OpenReport("d:applicationlist.rpt")
'treport.DiscardSavedData
'treport.Database.SetDataSource rst
'CRViewer1.ReportSource = treport
'CRViewer1.ViewReport
'++++++++++++++++++++++++++++++
  
'-Passing a SQL Query string
'-The select query to return the data required
'OUT =>  crxRpt.SQLQueryString = SQL   'Input your SELECT query here
    
   '-Carga Recordset
   crxRpt.Database.SetDataSource Rst
    
'- Oculta Arbol de grupos
  CRViewer1.DisplayGroupTree = False
 
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
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wdetalle1
       Case "NombRep"
         crxParamDef.AddCurrentValue wnombrep
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wnombrep
       Case "Impreso"
         crxParamDef.AddCurrentValue wUsuNombre
       Case "Campo0"
         crxParamDef.AddCurrentValue wcampo0
       Case "Campo1"
         crxParamDef.AddCurrentValue wcampo1
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wcampo1
       Case "Campo2"
         crxParamDef.AddCurrentValue wcampo2
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wcampo2
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
       Case "MesAno"
         crxParamDef.AddCurrentValue wMesAno
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wMesdeAno
       Case "Ano"
         crxParamDef.AddCurrentValue wAno
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wAno
       Case "NroResolu"
         crxParamDef.AddCurrentValue wNroResolu
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wNroresolu
       Case "Onomastico"
         crxParamDef.AddCurrentValue Wonomastico
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wonomastico
       Case "NroReso0"
         crxParamDef.AddCurrentValue wNroReso0
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wNroreso0
       Case "FechaR0"
         crxParamDef.AddCurrentValue WfechaR0
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wFechaR0
         
       Case "Trabajador"
         crxParamDef.AddCurrentValue Wtrabajador
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wtrabajador
       Case "Thr_Prog"
         crxParamDef.AddCurrentValue wtHr_Prog
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wtHr_Prog
       Case "M_Tard"
         crxParamDef.AddCurrentValue wm_tard
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wm_tard
       Case "Hr_Tard"
         crxParamDef.AddCurrentValue whr_tard
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue whr_tard
       Case "M_Tm130"
         crxParamDef.AddCurrentValue wm_tm130
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wHr_tm130
       Case "Hr_Tm130"
         crxParamDef.AddCurrentValue whr_tm130
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wHr_tm130
       Case "Thr_Fmp"
         crxParamDef.AddCurrentValue wtHr_FMp
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wtHr_FMp
       Case "Thr_Mp"
         crxParamDef.AddCurrentValue wtHr_Mp
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wtHr_Mp
       Case "Thr_Fj"
         crxParamDef.AddCurrentValue wtHr_Fj
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wtHr_Fj
       Case "Thr_Fi"
         crxParamDef.AddCurrentValue wtHr_Fi
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wtHr_Fi
       Case "Thr_Trab"
         crxParamDef.AddCurrentValue wtHr_trab
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wtHr_trab
       End Select
  Next
  
 '==== DUPLEX  ====
 '-Set the Paper source to the lower bin
  'crxRpt.PaperSource = crPRBinLower
 '-Set Duplex printing to horizontal
  'crxRpt.PrinterDuplex = crPRDPHorizontal
 '================
 
 '-Muestra dialogo para seleccion PRINTER
 crxRpt.PrinterSetup (0)
 crxRpt.PaperSize = crPaperA4
  
 '-Set the Report to the Report Viewer
 CRViewer1.ReportSource = crxRpt
 If nTipRpt = 90 Then
    '-Aplica formula para seleccionar registro
    crxRpt.RecordSelectionFormula = strSelectionfrm
 End If
 CRViewer1.ViewReport
 '-Set the Zoom on the report
 CRViewer1.Zoom (100)
  
  
  Screen.MousePointer = vbDefault
  Set crxParamFields = Nothing
  Set crxParamField = Nothing
      Rst.Close
      Set Rst = Nothing
  Exit Sub
ErrHandler:
If Err.Number = -2147206461 Then
   MsgBox "El archivo de reporte no se encuentra, restáurelo ", vbCritical + vbOKOnly
Else
   MsgBox Err.Description, vbCritical + vbOKOnly
End If
  Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set crxRpt = Nothing
    Set crxApp = Nothing
End Sub



