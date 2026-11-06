VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Begin VB.Form InfoCpSQL 
   Caption         =   "Informes"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12465
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   Picture         =   "InfoCpSQL.frx":0000
   ScaleHeight     =   10035
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows Default
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRViewer1 
      Height          =   9375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   11415
      lastProp        =   600
      _cx             =   20135
      _cy             =   16536
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
Attribute VB_Name = "InfoCpSQL"
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
  wTitulo = "PROGRAMACION MENSUAL DE TURNOS Y HORARIOS DEL SERVICIO ASISTENCIAL" + IIf(ProgASISTENCIAL.TxRegPlh.Text = "3RO", " - TERCEROS", "")
  wdetalle1 = ""
  If nTipRpt = 0.1 Then
     wdetalle = "Servidores " + wTituRpt
     wnombrep = "(CpCadHep)"
  ElseIf nTipRpt = 1 Then
         wdetalle = "Servidores " + wTituRpt
         wnombrep = "(CpRegistro)"
  ElseIf nTipRpt = 1.01 Then
         wdetalle = "Servidores " + wTituRpt
         wnombrep = "(CpRegistro)"
         
  ElseIf nTipRpt = 1.1 Then
         wdetalle = "Servidores " + wTituRpt
         wnombrep = "(CpQMultiple)"
         wcampo1 = QMultiple.Txcampo1.Text
         wcampo2 = QMultiple.Txcampo2.Text
  ElseIf nTipRpt = 2 Then
         wdetalle = "Tarjeta Nº " & "'" & Trim(General.Txtarjeta.Text) & "'"
         wnombrep = "(CpFicha)"
  ElseIf nTipRpt = 3 Then
         wdetalle = "Reporte de Autorizaciones " + c_periodo
         wnombrep = "(CpAutoriza)"
  ElseIf nTipRpt = 3.01 Then
         wdetalle = "Producción de Autorizaciones "
         wnombrep = "(CpBAutReg)"
  ElseIf nTipRpt = 3.02 Then
         wdetalle = "Autorizaciones del Concepto "
         wnombrep = "(CpAutLst)"
         
  ElseIf nTipRpt = 3.1 Then
         wdetalle = "Reporte de Cambios de Turno "
         wnombrep = "(CpCambiosT)"
  ElseIf nTipRpt = 4 Then
         wdetalle = "Marcación del Dia : " & pridia & "  Al : " & ultdia
         wnombrep = "(CpRegRelj)"
         
  ElseIf nTipRpt = 6 Then
         wdetalle = "Kardex de Asistencia Mensual "
         wnombrep = "(CpDirKar)"
  ElseIf nTipRpt = 8 Then
         wdetalle = "Record de Tardanzas  del: " & TardaInasis.Txfecini.Text & "  Al: " & TardaInasis.TxFecFin.Text
         wdetalle1 = IIf(Trim(TardaInasis.Txdetalle1) = "", "", "(" & TardaInasis.Txdetalle1.Text & ")")
         wnombrep = "(CpTardanzas)"
  ElseIf nTipRpt = 9 Then
         wdetalle = "Record de Inasistencia del: " & TardaInasis.Txfecini.Text & "  Al: " & TardaInasis.TxFecFin.Text
         wdetalle1 = IIf(Trim(TardaInasis.Txdetalle1) = "", "", "(" & TardaInasis.Txdetalle1.Text & ")")
         wnombrep = "(CpInasistencia)"
  ElseIf nTipRpt = 10 Then
         wdetalle = "Jefaturas programadas en : " & JGuardias.TxNombMes.Text & " " & genaño
         wnombrep = "(CpJGuardia)"
  ElseIf nTipRpt = 12 Then
         wrplh = IIf(ResDscto.TxRegPlh.Text = "HEP", "<NOMB./CONTR.>", IIf(ResDscto.TxRegPlh.Text = "CAS", "<C.A.S.>", IIf(ResDscto.TxRegPlh.Text = "3RO", "<TERCEROS>", "")))
         wdetalle = ResDscto.TxFecActual.Text
         wMesdeAno = ResDscto.TxMesdeAno.Text
         wMesAno = ResDscto.TxMesAno.Text
         wAno = Right(ResDscto.TxMesAno.Text, 4)
         wNroResolu = ResDscto.TxNroResolu.Text
  ElseIf nTipRpt = 12.1 Then
         wrplh = IIf(ResMP.TxRegPlh.Text = "HEP", "<NOMB./CONTR.>", IIf(ResMP.TxRegPlh.Text = "CAS", "<C.A.S.>", IIf(ResMP.TxRegPlh.Text = "3RO", "<TERCEROS>", "")))
         wdetalle = ResMP.TxFecActual.Text
         wMesdeAno = ResMP.TxMesdeAno.Text
         wMesAno = ResMP.TxMesAno.Text
         wAno = Right(ResMP.TxMesAno.Text, 4)
         wNroResolu = ResMP.TxNroResolu.Text
  ElseIf nTipRpt = 13 Or nTipRpt = 13.01 Then
         wdetalle = ResPgoGuarProy.TxFecActual.Text
         wdetalle1 = "(Proyectado)"
         wMesAno = ResPgoGuarProy.TxMesAno.Text
         wMesdeAno = ResPgoGuarProy.TxMesdeAno.Text
         wNroResolu = ResPgoGuarProy.TxNroResolu.Text
  ElseIf nTipRpt = 13.5 Or nTipRpt = 13.51 Then
         wdetalle = PresPgoGuar.TxFecActual.Text
         wMesAno = PresPgoGuar.TxMesAno.Text
         wMesdeAno = PresPgoGuar.TxMesdeAno.Text
         wNroResolu = PresPgoGuar.TxNroResolu.Text
         wmsgpie = "(Cálculo basado en tarifas al " + Format(Date, "dd/mm/yyyy") + ")"
  ElseIf nTipRpt = 14 Then
         wdetalle = ResPagoJGuar.TxFecActual.Text
         wMesdeAno = ResPagoJGuar.TxMesdeAno.Text
         wNroResolu = ResPagoJGuar.TxNroResolu.Text
  ElseIf nTipRpt = 15 Then
         wdetalle = ResPgoGuarDscto.TxFecActual.Text
         wMesAno = ResPgoGuarDscto.TxMesAno.Text
         wMesdeAno = ResPgoGuarDscto.TxMesdeAno.Text
         wNroResolu = ResPgoGuarDscto.TxNroResolu.Text
         wNroReso0 = ResPgoGuarDscto.TxNroReso0.Text
         WfechaR0 = ResPgoGuarDscto.TxFechaR0.Text
  ElseIf nTipRpt = 16 Then
         wdetalle = ResPgoGuarAbon.TxFecActual.Text
         wMesAno = ResPgoGuarAbon.TxMesAno.Text
         wMesdeAno = ResPgoGuarAbon.TxMesdeAno.Text
         wNroResolu = ResPgoGuarAbon.TxNroResolu.Text
         wNroReso0 = ResPgoGuarAbon.TxNroReso0.Text
         WfechaR0 = ResPgoGuarAbon.TxFechaR0.Text
  End If
'-Open the report
  Set crxApp = New CRAXDRT.Application
  If nTipRpt = 0.1 Then
     Set crxRpt = crxApp.OpenReport(App.path & "\CpCadHep.rpt", 1)
  ElseIf nTipRpt = 1 Or nTipRpt = 1.01 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpRegistr0.rpt", 1)
  'ElseIf nTipRpt = 1.01 Then
  '       Set crxRpt = crxApp.OpenReport(App.path & "\CpRegistro.rpt", 1)
  ElseIf nTipRpt = 1.1 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpQMultiple.rpt", 1)
  ElseIf nTipRpt = 2 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpFicha.rpt", 1)
         
  ElseIf nTipRpt = 3 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpAutoriza.rpt", 1)
  ElseIf nTipRpt = 3.01 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpBAutReg.rpt", 1)
  ElseIf nTipRpt = 3.02 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpAutLst.rpt", 1)
         
  ElseIf nTipRpt = 3.1 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpCambiosT.rpt", 1)
         
  ElseIf nTipRpt = 4 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpRegRelj.rpt", 1)
  ElseIf nTipRpt = 6 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpDirKar.rpt", 1)
  ElseIf nTipRpt = 8 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpTardanzas.rpt", 1)
  ElseIf nTipRpt = 9 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpInasistencia.rpt", 1)
         
  ElseIf nTipRpt = 10 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpJGuardia.rpt", 1)
  ElseIf nTipRpt = 12 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpResDscto.rpt", 1)
  ElseIf nTipRpt = 12.1 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpResMp.rpt", 1)
  ElseIf nTipRpt = 13 Or nTipRpt = 13.5 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpResPgoGuar.rpt", 1)
  ElseIf nTipRpt = 13.01 Or nTipRpt = 13.51 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpResPgoGuaA.rpt", 1)
  ElseIf nTipRpt = 14 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpResJGuard.rpt", 1)
  ElseIf nTipRpt = 15 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpResPgoGuarDscto.rpt", 1)
  ElseIf nTipRpt = 16 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpResPgoGuarAbon.rpt", 1)
  ElseIf nTipRpt = 17 Then
         Set crxRpt = crxApp.OpenReport(App.path & "\CpDirXML.rpt", 1)
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
  If nTipRpt = 0.1 Then
     sql = " SELECT DISTINCT " & _
     "CpTmGeneral.cadHep, CpTmGeneral.apelpat, CpTmGeneral.apelmat, CpTmGeneral.nombres, CpTmGeneral.tarjeta, CpTmGeneral.documento, " & _
     "CpTmGeneral.dsccadena, CpTmGeneral.ccosto, CpTmGeneral.dscccosto, CpTmGeneral.codupr, CpTmGeneral.dscupr " & _
     "FROM  CpTmGeneral " & _
     "ORDER BY CpTmGeneral.ccosto,  CpTmGeneral.codupr, CpTmGeneral.apelpat, CpTmGeneral.apelmat, CpTmGeneral.nombres "
       
  ElseIf nTipRpt = 1 Or nTipRpt = 1.01 Then
     sql = " SELECT DISTINCT " & _
    "CpTmGeneral.ccosto, CpTmGeneral.codUpr, CpTmGeneral.apelpat, CpTmGeneral.apelmat, CpTmGeneral.nombres, CpTmGeneral.tarjeta, CpTmGeneral.fecingr, CpTmGeneral.grpPer, " & _
    "CpTmGeneral.grpGen, CpTmGeneral.TipoPla, CpTmGeneral.dscccosto, CpTmGeneral.dsccarpla, CpTmGeneral.nivrem " & _
    "FROM  CpTmGeneral " & _
    "ORDER BY CpTmGeneral.ccosto, CpTmGeneral.codUpr, CpTmGeneral.apelpat, CpTmGeneral.apelmat"
  'ElseIf nTipRpt = 1.01 Then
  '   sql = " SELECT DISTINCT " & _
  '  "CpTmGeneral.ccosto, CpTmGeneral.codupr, CpTmGeneral.apelpat, CpTmGeneral.apelmat, CpTmGeneral.nombres, CpTmGeneral.tarjeta, CpTmGeneral.fecingr, CpTmGeneral.grpPer, " & _
  '  "CpTmGeneral.grpGen, CpTmGeneral.TipoPla, CpTmGeneral.dscccosto, CpTmGeneral.dsccarpla, CpTmGeneral.nivrem " & _
  '  "FROM  CpTmGeneral " & _
  '  "ORDER BY CpTmGeneral.ccosto, CpTmGeneral.codupr, CpTmGeneral.apelpat, CpTmGeneral.apelmat"
  ElseIf nTipRpt = 1.1 Then
         sql = " SELECT DISTINCT " & _
         "CpTmGeneral.ccosto, CpTmGeneral.codupr, CpTmGeneral.apelpat, CpTmGeneral.apelmat, CpTmGeneral.nombres, CpTmGeneral.tarjeta, CpTmGeneral.TipoPla, CpTmGeneral.documento, " & _
         "CpTmGeneral.Plh, CpTmGeneral.categoria, CpTmGeneral.dscCategoria, CpTmGeneral.dscccosto, CpTmGeneral.dscupr, CpTmGeneral.campo1, CpTmGeneral.campo2 " & _
         "FROM  CpTmGeneral " & _
         "ORDER BY CpTmGeneral.ccosto, CpTmGeneral.codupr, CpTmGeneral.apelpat, CpTmGeneral.apelmat"
  ElseIf nTipRpt = 2 Then
         sql = "SELECT DISTINCT " & _
         "CpTmGeneral.fecingr, CpTmGeneral.estado, CpTmGeneral.documento, CpTmGeneral.apelpat, CpTmGeneral.apelmat, CpTmGeneral.nombres, CpTmGeneral.actividad, CpTmGeneral.dscactividad, CpTmGeneral.cadena, " & _
         "CpTmGeneral.dsccadena, CpTmGeneral.grupo, CpTmGeneral.dscgrupo, CpTmGeneral.ccosto, CpTmGeneral.dscccosto, CpTmGeneral.categoria, CpTmGeneral.dsccategoria, CpTmGeneral.dsccarpla, CpTmGeneral.dsccarest, " & _
         "CpTmGeneral.nivrem, CpTmGeneral.remtot, CpTmGeneral.cargo, CpTmGeneral.dsccargo,   " & _
         "FROM  CpTmGeneral"
  ElseIf nTipRpt = 3 Then
         sql = "SELECT DISTINCT " & _
         "CpTmpTabla.tarjeta, CpTmpTabla.codigo, CpTmpTabla.documento, CpTmpTabla.Diahora, CpTmpTabla.fechaini, CpTmpTabla.fechafin, CpTmpTabla.valor, CpTmpTabla.horaIni, CpTmpTabla.horafin, CpTmpTabla.trabajador " & _
         "FROM  CpTmpTabla "
  ElseIf nTipRpt = 3.01 Then
         sql = "SELECT DISTINCT " & _
         "CpTmpTabla.Usuario,CpTmpTabla.fecha,CpTmpTabla.tarjeta, CpTmpTabla.trabajador, CpTmpTabla.codigo, CpTmpTabla.documento, CpTmpTabla.Diahora, CpTmpTabla.fechaini, CpTmpTabla.fechafin, CpTmpTabla.valor, CpTmpTabla.horaIni, CpTmpTabla.horafin " & _
         "FROM  CpTmpTabla  ORDER BY CpTmpTabla.usuario, CpTmpTabla.fecha, CpTmpTabla.trabajador"
  ElseIf nTipRpt = 3.02 Then
         sql = "SELECT DISTINCT " & _
         "CpTmpTabla.tarjeta, CpTmpTabla.codigo, CpTmpTabla.documento, CpTmpTabla.Diahora, CpTmpTabla.fechaini, CpTmpTabla.fechafin, CpTmpTabla.valor, CpTmpTabla.horaIni, CpTmpTabla.horafin, CpTmpTabla.trabajador " & _
         "FROM  CpTmpTabla ORDER BY CpTmpTabla.trabajador, CpTmpTabla.fechaini "
         
  ElseIf nTipRpt = 3.1 Then
         sql = "SELECT DISTINCT " & _
         "CpTmpTabla.tarjeta, CpTmpTabla.documento, CpTmpTabla.trabajador, CpTmpTabla.tipo, CpTmpTabla.fechaini, CpTmpTabla.fechafin, CpTmpTabla.detur0, CpTmpTabla.atur0, CpTmpTabla.detur, CpTmpTabla.atur, CpTmpTabla.usuario " & _
         "FROM  CpTmpTabla " & _
         "ORDER BY CpTmpTabla.trabajador, CpTmpTabla.fechaini"
  ElseIf nTipRpt = 4 Then
         sql = " SELECT DISTINCT " & _
         "CpTmRecord.fecha, CpTmRecord.tarjeta, CpGeneral.ApelPAT, CpGeneral.ApelMAT, CpGeneral.NOMBRES, CpTmRecord.ming, CpTmRecord.msal, CpTmRecord.tardanza, CpTmRecord.hrasProg" & _
         " FROM  CpTmRecord INNER JOIN CpGeneral ON CpTmRecord.tarjeta=CpGeneral.Tarjeta and (cpgeneral.baja is NULL)"
         
         
  ElseIf nTipRpt = 6 Then
         sql = "SELECT DISTINCT " & _
         "CpTmpDirKar.d1, CpTmpDirKar.d2, CpTmpDirKar.d3, CpTmpDirKar.d4, CpTmpDirKar.d5, CpTmpDirKar.d6, CpTmpDirKar.d7, CpTmpDirKar.d8, CpTmpDirKar.d9, CpTmpDirKar.d10, CpTmpDirKar.d11, " & _
         "CpTmpDirKar.d12, CpTmpDirKar.d13, CpTmpDirKar.d14, CpTmpDirKar.d15, CpTmpDirKar.d16, CpTmpDirKar.d17, CpTmpDirKar.d18, CpTmpDirKar.d19, CpTmpDirKar.d20, CpTmpDirKar.d21, CpTmpDirKar.d22, " & _
         "CpTmpDirKar.d23, CpTmpDirKar.d24, CpTmpDirKar.d25, CpTmpDirKar.d26, CpTmpDirKar.d27, CpTmpDirKar.d28, CpTmpDirKar.d29, CpTmpDirKar.d30, CpTmpDirKar.d31,  CpTmpDirKar.tarjeta, " & _
         "CpGeneral.ApelPAT, CpGeneral.ApelMAT, CpGeneral.NIVREM, CpTmpDirKar.DscCargo, CpActividad.Descripcion, CpTmpDirKar.mes,CpTmpDirKar.ano " & _
         "FROM  (CpGeneral INNER JOIN CpTmpDirKar ON CpGeneral.Tarjeta=CpTmpDirKar.tarjeta) INNER JOIN CpActividad " & _
         "ON CpGeneral.ACTIVIDAD=CpActividad.CODIGO ORDER BY CpTmpDirKar.tarjeta, CpTmpDirKar.ano, CpTmpDirKar.mes"
  ElseIf nTipRpt = 8 Then
         sql = "SELECT DISTINCT " & _
         "CpTmRecord.fecha, CpTmRecord.CCosto, CpTmRecord.DscCCosto, CpTmRecord.CodUpr, CpTmRecord.DscUpr, CpTmRecord.tarjeta, CpGeneral.ApelMAT, CpGeneral.ApelPAT, CpGeneral.NOMBRES, CpTmRecord.ping, CpTmRecord.msal, CpTmRecord.tardanza, CpTmRecord.ming, CpTmRecord.psal " & _
         "FROM  CpTmRecord INNER JOIN CpGeneral ON CpTmRecord.tarjeta=CpGeneral.Tarjeta " & _
         "ORDER BY CpTmRecord.fecha, CpGeneral.ApelPAT, CpGeneral.ApelMAT, CpGeneral.NOMBRES"
  ElseIf nTipRpt = 9 Then
         sql = "SELECT DISTINCT " & _
         "CpTmRecord.fecha, CpTmRecord.turno, CpTmRecord.CCosto, CpTmRecord.DscCCosto, CpTmRecord.CodUpr, CpTmRecord.DscUpr, CpTmRecord.tarjeta, CpGeneral.ApelMAT, CpGeneral.ApelPAT, CpGeneral.NOMBRES, CpTmRecord.dscactividad, CpTmRecord.PIng,CpTmRecord.Ming " & _
         "FROM  CpTmRecord INNER JOIN CpGeneral ON CpTmRecord.tarjeta=CpGeneral.Tarjeta " & _
         "ORDER BY CpTmRecord.fecha, CpGeneral.ApelPAT, CpGeneral.ApelMAT, CpGeneral.NOMBRES"
  ElseIf nTipRpt = 10 Then
         sql = "SELECT DISTINCT " & _
         "CpTmpguardia.Dia, CpTmpguardia.turno, CpTmGuardia.dscTurno, CpTmpguardia.Tarjeta,CpTmpguardia.ApelPAT, CpTmpGuardia.ApelMAT, CpTmpGuardia.NOMBRES, CpTmpGuardia.NivGuardia, CpTmpGuardia.Tarifa " & _
         " FROM  CpTmpGuardia " & _
         " ORDER BY CpTmpGuardia.dia, CpTmpGuardia.Turno"
  ElseIf nTipRpt = 12 Then
         sql = "SELECT DISTINCT " & _
         "CpTmResDscto.cCosto, CpTmResDscto.codUpr, CpTmResDscto.apelMAT, CpTmResDscto.apelPAT, CpTmResDscto.NOMBRES, CpTmResDscto.CadHep, CpTmresDscto.NivGuardia, CpTmResDscto.DiasFalt, CpTmResDscto.Descuento, " & _
         "CpTmResDscto.DscCosto, CpTmResDscto.DscUpr, CpTmResDscto.dsccadhep,CpTmResDscto.dscCarPla, CpTmResDscto.RemHor, CpTmResDscto.Monto " & _
         "FROM   CpTmResDscto " & _
         "ORDER BY CpTmResDscto.cCosto, CpTmResDscto.codUpr, CpTmResDscto.apelMAT, CpTmResDscto.apelPAT"
  ElseIf nTipRpt = 12.1 Then
         sql = "SELECT DISTINCT " & _
         "CpTmResMP.cCosto, CpTmResMP.CodUpr, CpTmResMP.apelpat, CpTmResMP.apelmat, CpTmResMP.nombres, CpTmResMP.NivGuardia, CpTmResMP.RemHor, CpTmResMP.DiasFalt,CpTmResMP.HrsMp, CpTmResMP.Monto, " & _
         "CpTmResMP.DscCosto, CpTmResMP.DscUpr, CpTmResMP.cadhep, CpTmResMP.dsccadhep " & _
         "FROM  CpTmResMP " & _
         "ORDER BY CpTmResMP.cCosto, CpTmResMP.CodUpr, CpTmResMp.apelMAT, CpTmResMp.apelPAT"
  ElseIf nTipRpt = 13 Or nTipRpt = 13.5 Then
         sql = "SELECT DISTINCT " & _
         "cpTmPagoGuar.cCosto, cpTmPagoGuar.CodUpr, cpTmPagoGuar.apelpat, cpTmPagoGuar.apelmat, cpTmPagoGuar.nombres, cpTmPagoGuar.D_O, cpTmPagoGuar.N_O, cpTmPagoGuar.D_F, cpTmPagoGuar.N_F, cpTmPagoGuar.Monto, " & _
         "cpTmPagoGuar.dsccosto, cpTmPagoGuar.dscUpr, cpTmPagoGuar.nivGuardia, cpTmPagoGuar.cadhep, cpTmPagoGuar.dsccadhep , cpTmPagoGuar.dsccargo " & _
         "FROM   cpTmPagoGuar " & _
         "ORDER BY cpTmPagoGuar.cCosto, cpTmPagoGuar.CodUpr, cpTmPagoGuar.apelpat, cpTmPagoGuar.apelmat "
  ElseIf nTipRpt = 13.01 Or nTipRpt = 13.51 Then
         sql = "SELECT DISTINCT " & _
         "cpTmPagoGuar.cCosto, cpTmPagoGuar.CodUpr, cpTmPagoGuar.apelpat, cpTmPagoGuar.apelmat, cpTmPagoGuar.nombres, cpTmPagoGuar.D_O, cpTmPagoGuar.N_O, cpTmPagoGuar.D_F, cpTmPagoGuar.N_F, cpTmPagoGuar.Monto, " & _
         "cpTmPagoGuar.dsccosto, cpTmPagoGuar.dscUpr, cpTmPagoGuar.nivGuardia, cpTmPagoGuar.cadhep, cpTmPagoGuar.dsccadhep , cpTmPagoGuar.dsccargo " & _
         "FROM   cpTmPagoGuar ORDER BY cpTmPagoGuar.apelpat, cpTmPagoGuar.apelmat, cpTmPagoGuar.nombres"
  ElseIf nTipRpt = 14 Then
         sql = "SELECT DISTINCT " & _
         "cpTmResJGuar.cCosto, cpTmResJGuar.cCodUpr, cpTmResJGuar.apelpat, cpTmResJGuar.apelmat, cpTmResJGuar.nombres, cpTmResJGuar.nivGuardia, cpTmResJGuar.D_O, cpTmResJGuar.N_O, " & _
         "cpTmResJGuar.D_F, cpTmResJGuar.N_F, cpTmResJGuar.Monto, cpTmResJGuar.DscCosto, cpTmResJGuar.DscUpr, cpTmResJGuar.cadhep, cpTmResJGuar.dsccadhep, cpTmResJGuar.dsccarpla " & _
         "FROM   cpTmResJGuar " & _
         "ORDER BY cpTmResJGuar.cCosto, cpTmResJGuar.CodUpr"
  ElseIf nTipRpt = 15 Then
         sql = "SELECT DISTINCT " & _
         "cpTmDifPgoGuar.cCosto, cpTmDifPgoGuar.CodUpr, cpTmDifPgoGuar.nivGuardia, cpTmDifPgoGuar.D_O, cpTmDifPgoGuar.N_O, cpTmDifPgoGuar.D_F, cpTmDifPgoGuar.N_F, cpTmDifPgoGuar.Monto, cpTmDifPgoGuar.Monto0, " & _
         "cpTmDifPgoGuar.apelpat, cpTmDifPgoGuar.apelmat, cpTmDifPgoGuar.nombres, cpTmDifPgoGuar.DscCosto, cpTmDifPgoGuar.DscUpr, cpTmDifPgoGuar.cadhep, cpTmDifPgoGuar.dsccadhep , cpTmDifPgoGuar.dsccarPla " & _
         "FROM   cpTmDifPgoGuar " & _
         "ORDER BY cpTmDifPgoGuar.cCosto, cpTmDifPgoGuar.CodUpr, cpTmDifPgoGuar.apelpat, cpTmDifPgoGuar.apelmat "
  ElseIf nTipRpt = 16 Then
         sql = "SELECT DISTINCT " & _
         "cpTmDifPgoGuar.cCosto, cpTmDifPgoGuar.CodUpr, cpTmDifPgoGuar.nivGuardia, cpTmDifPgoGuar.D_O, cpTmDifPgoGuar.N_O, cpTmDifPgoGuar.D_F, cpTmDifPgoGuar.N_F, cpTmDifPgoGuar.Monto, cpTmDifPgoGuar.Monto0, " & _
         "cpTmDifPgoGuar.apelpat, cpTmDifPgoGuar.apelmat, cpTmDifPgoGuar.nombres, cpTmDifPgoGuar.DscCosto, cpTmDifPgoGuar.DscUpr, cpTmDifPgoGuar.cadhep, cpTmDifPgoGuar.dsccadhep , cpTmDifPgoGuar.dsccarPla " & _
         "FROM   cpTmDifPgoGuar " & _
         "ORDER BY cpTmDifPgoGuar.cCosto, cpTmDifPgoGuar.CodUpr, cpTmDifPgoGuar.apelpat, cpTmDifPgoGuar.apelmat "
   ElseIf nTipRpt = 17 Then
         sql = "SELECT DISTINCT " & _
         "CpTmDirXML.fecha, CpTmDirXML.dia, CpTmDirXML.servicio, CpTmDirXML.apellidos, CpTmDirXML.nombres, CpTmDirXML.horario " & _
         "FROM  CpTmDirXML " & _
         "ORDER BY  CpTmDirXML.servicio,CpTmDirXML.fecha, CpTmDirXML.apellidos"
  End If
    
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
  crxRpt.SQLQueryString = sql   'Input your SELECT query here
  
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
       Case "RegPlh"
         crxParamDef.AddCurrentValue wRegPlh
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wRegPlh
       Case "RPlh"
         crxParamDef.AddCurrentValue wrplh
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wRPlh
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
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wusunombre
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
       Case "msgpie"
         crxParamDef.AddCurrentValue wmsgpie
       Case "Onomastico"
         crxParamDef.AddCurrentValue Wonomastico
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wonomastico
       Case "NroReso0"
         crxParamDef.AddCurrentValue wNroReso0
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wNroreso0
       Case "FechaR0"
         crxParamDef.AddCurrentValue WfechaR0
         '-crxRpt.ParameterFields.GetItemByName(crxParamDef.ParameterFieldName).AddCurrentValue wFechaR0
       End Select
  Next
'-View the report
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



