VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptEPartidaResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Partidas Resumen
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim moConexion As New Connection
Dim lnIdServiciosHospitalarios As Long, lnIdReembolsosServicios As Long
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja

Sub EjecutaFormulario()
    Dim oFormulario As New EconPartidaResumen
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub



'02/05/2016
Sub CreaDatosReporte(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, _
                     mda_FechaInicio As Date, mda_FechaFin As Date, ml_IdPartidaFiltro As Long, _
                     lnHwnd As Long, lbEnValor As Boolean, lnIdCajero As Long, lbSoloCreditos As Boolean)
        Dim vmrs_Tmp As Double                     'debb-12/07/2016
        Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim mrs_Tmp As New Recordset
        Dim mrs_Tmp1 As New Recordset
        Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
        Dim oRsTmp9 As New Recordset
        Dim oConexion As New Connection
        Dim mo_sighProxies As New SIGHProxies.Procesos
        Dim lcPie As String, ldFechaProceso As Date, lbPrimeraVez As Boolean, ldFechaProceso1 As Date
        Dim ldFechaUltimaBoleta As Date, ldUltimaFechaProcesada As Date
        Dim lcSerie As String, lcCajaCodigo As String, lcCaja As String, lcBoletaInicio As String, lcBoletaFin As String
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        '
        With mrs_Tmp1
            .Fields.Append "CajaCodigo", adVarChar, 10, adFldIsNullable
            .Fields.Append "Caja", adVarChar, 100, adFldIsNullable
            .Fields.Append "Serie", adVarChar, 5, adFldIsNullable
            .Fields.Append "Desde", adVarChar, 10, adFldIsNullable
            .Fields.Append "Hasta", adVarChar, 10, adFldIsNullable
            .LockType = adLockOptimistic
            .Open
        End With
        
        Set oRsTmp9 = mo_ReglasCaja.CajaComprobantesPagoPorFechasConNombreCajas(mda_FechaInicio, mda_FechaFin)
        '
        If lnIdCajero > 0 Then
           oRsTmp9.Filter = "idCajero=" & lnIdCajero
        End If
        '
        If oRsTmp9.RecordCount > 0 Then
           oRsTmp9.MoveFirst
           Do While Not oRsTmp9.EOF
              lcSerie = oRsTmp9!nroSerie
              lcCajaCodigo = oRsTmp9!cajaCodigo
              lcCaja = oRsTmp9!caja
              lcBoletaInicio = oRsTmp9!nrodocumento
              Do While Not oRsTmp9.EOF And lcCajaCodigo = oRsTmp9!cajaCodigo And lcSerie = oRsTmp9!nroSerie
                    lcBoletaFin = oRsTmp9!nrodocumento
                    oRsTmp9.MoveNext
                    If oRsTmp9.EOF Then
                       Exit Do
                    End If
              Loop
              mrs_Tmp1.AddNew
              mrs_Tmp1.Fields!cajaCodigo = lcCajaCodigo
              mrs_Tmp1.Fields!caja = Left(lcCaja, 100)
              mrs_Tmp1.Fields!serie = lcSerie
              mrs_Tmp1.Fields!desde = lcBoletaInicio
              mrs_Tmp1.Fields!hasta = lcBoletaFin
              mrs_Tmp1.Update
           Loop
        End If
        oRsTmp9.Close
        '
        ldFechaProceso1 = CDate("31/12/" & Trim(Str(Year(mda_FechaFin))))
        ldUltimaFechaProcesada = CDate("01/01/1990")
        Set oRsTmp9 = mo_ReglasCaja.FactPartidasPresupuestalesXMesSelecionaUltimoProceso(ldFechaProceso1, oConexion)
        If oRsTmp9.RecordCount > 0 Then
           ldUltimaFechaProcesada = CDate(Format(oRsTmp9!fecha, sighentidades.DevuelveFechaSoloFormato_DMY))
        End If
        '
        
        
        mo_sighProxies.ReportePartidaREsumen mrs_Tmp, mda_FechaInicio, mda_FechaFin, ml_IdPartidaFiltro, _
                                            sghSoloParaReporte, oConexion, True, lnIdCajero, lbSoloCreditos
        oConexion.Close
        '*****************************************************************
        'Decisiòn para la impresion de partidas mayores a cero
        'A.Yañez 26/09/2014
        '*****************************************************************
        If mrs_Tmp.RecordCount > 0 Then
            If lbEnValor = True Then
                mrs_Tmp.MoveFirst
                Do While Not mrs_Tmp.EOF
                   If mrs_Tmp.Fields!ImpCancelado = 0 Then
                      mrs_Tmp.Delete
                   End If
                   mrs_Tmp.MoveNext
                Loop
                vmrs_Tmp = 1
            Else
                mrs_Tmp.MoveFirst
                Do While Not mrs_Tmp.EOF
                   vmrs_Tmp = vmrs_Tmp + mrs_Tmp.Fields!ImpCancelado
                   mrs_Tmp.MoveNext
                Loop
            End If
        End If
        '***************************************************************
        If mrs_Tmp.RecordCount = 0 Then
             MsgBox "No existe Datos", vbInformation, "Reporte"
        Else
          If vmrs_Tmp = 0 And ml_IdPartidaFiltro = 0 Then
             MsgBox "No existe Datos", vbInformation, "Reporte"
          Else
             If ml_IdPartidaFiltro > 0 Then
                mrs_Tmp.Sort = "Descripcion"
            End If
             
             lcPie = ""
             
             If lbEnExcel = True Then
                 mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
             Else
                 Set RpEPartidaResumen.DataSource = mrs_Tmp
                 RpEPartidaResumen.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
                 RpEPartidaResumen.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
                 RpEPartidaResumen.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
                 RpEPartidaResumen.Sections("cabecera").Controls("lblFecha").Caption = "Fecha: " & lcBuscaParametro.RetornaFechaServidorSQL
                 RpEPartidaResumen.Sections("cabecera").Controls("lblHora").Caption = "Hora: " & lcBuscaParametro.RetornaHoraServidorSQL
                 
                 RpEPartidaResumen.Sections("cabecera").Controls("lblPc").Caption = "PC: " & sighentidades.RetornaNombrePC
                 RpEPartidaResumen.Sections("cabecera").Controls("lblUsuario").Caption = "Usuario: " & lcBuscaParametro.RetornaLoginUsuario(sighentidades.Usuario)
                 RpEPartidaResumen.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
                 RpEPartidaResumen.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
                 RpEPartidaResumen.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaHoraServidorSQL
                 Set RpEPartidaResumen.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
                 RpEPartidaResumen.Sections("pie").Controls("lblPie").Caption = lcPie
                 RpEPartidaResumen.RightMargin = 100
                 RpEPartidaResumen.TopMargin = 100
                 RpEPartidaResumen.LeftMargin = 100
                 RpEPartidaResumen.BottomMargin = 100
                 RpEPartidaResumen.Orientation = rptOrientDefault
                 RpEPartidaResumen.PrintReport True    '  RpEPartidaResumen.Show 1
                 'debb-27/05/2015
                 Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
                 mo_ReglasComunes.grabaTablaAuditoria ("RpEPartidaResumen: " & Mid(lcSubTitulo, IIf(InStr(lcSubTitulo, "FILTROS: ") > 0, 10, 1)))
                 Set mo_ReglasComunes = Nothing
                 '
             End If
             If ml_IdPartidaFiltro = 0 Then
                 mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp1, lcTitulo, lcSubTitulo, lcPie, lnHwnd
             End If
          End If
       End If
'******************************************************************************************************************************************************************
        Set mo_ReglasCaja = Nothing
        Set lcBuscaParametro = Nothing
        Set mrs_Tmp = Nothing
        Set mo_ReglasReportes = Nothing
        Set oRsTmp9 = Nothing
        Set oConexion = Nothing
End Sub





Sub CreaDatosReporteConNotaDeCredito(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, _
                     mda_FechaInicio As Date, mda_FechaFin As Date, ml_IdPartidaFiltro As Long, _
                     lnHwnd As Long, lbEnValor As Boolean, lnIdCajero As Long)
        Dim vmrs_Tmp As Double                     'debb-12/07/2016
        Dim mo_ReglasCaja As New SIGHNegocios.ReglasCaja
        Dim lcBuscaParametro As New SIGHDatos.Parametros
        Dim mrs_Tmp As New Recordset
        Dim mrs_Tmp1 As New Recordset
        Dim oRsNotaCredito As New Recordset
        Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
        Dim oRsTmp9 As New Recordset
        Dim oRsTmp10 As New Recordset
        Dim oRsCtas As New Recordset
        Dim oConexion As New Connection
        Dim mo_sighProxies As New SIGHProxies.Procesos
        Dim lcPie As String, ldFechaProceso As Date, lbPrimeraVez As Boolean, ldFechaProceso1 As Date
        Dim ldFechaUltimaBoleta As Date, ldUltimaFechaProcesada As Date, lcSql As String, lnTotal11 As Double
        Dim lcSerie As String, lcCajaCodigo As String, lcCaja As String, lcBoletaInicio As String, lcBoletaFin As String
        Dim lnFacturasTotal As Long, lnFacturasAnuladas As Long, lnBoletasTotal As Long, lnBoletasAnuladas As Long
        Dim lnEsNuevo As Boolean
actualizaParametros "paso el dim"
         
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        '
        With mrs_Tmp1
            .Fields.Append "CajaCodigo", adVarChar, 10, adFldIsNullable
            .Fields.Append "Caja", adVarChar, 100, adFldIsNullable
            .Fields.Append "Serie", adVarChar, 5, adFldIsNullable
            .Fields.Append "Desde", adVarChar, 10, adFldIsNullable
            .Fields.Append "Hasta", adVarChar, 10, adFldIsNullable
            .LockType = adLockOptimistic
            .Open
        End With
actualizaParametros "orstmp9"

        Set oRsTmp9 = mo_ReglasCaja.CajaComprobantesPagoPorFechasConNombreCajas(mda_FechaInicio, mda_FechaFin)
        '
        If lnIdCajero > 0 Then
           oRsTmp9.Filter = "idCajero=" & lnIdCajero
        End If
        '
        If oRsTmp9.RecordCount > 0 Then
           oRsTmp9.MoveFirst
           Do While Not oRsTmp9.EOF
              lcSerie = oRsTmp9!nroSerie
              lcCajaCodigo = oRsTmp9!cajaCodigo
              lcCaja = oRsTmp9!caja
              lcBoletaInicio = oRsTmp9!nrodocumento
              Do While Not oRsTmp9.EOF And lcCajaCodigo = oRsTmp9!cajaCodigo And lcSerie = oRsTmp9!nroSerie
                    lcBoletaFin = oRsTmp9!nrodocumento
                    oRsTmp9.MoveNext
                    If oRsTmp9.EOF Then
                       Exit Do
                    End If
              Loop
              mrs_Tmp1.AddNew
              mrs_Tmp1.Fields!cajaCodigo = lcCajaCodigo
              mrs_Tmp1.Fields!caja = Left(lcCaja, 100)
              mrs_Tmp1.Fields!serie = lcSerie
              mrs_Tmp1.Fields!desde = lcBoletaInicio
              mrs_Tmp1.Fields!hasta = lcBoletaFin
              mrs_Tmp1.Update
           Loop
        End If
        oRsTmp9.Close
        
        
actualizaParametros "cerro orstmp9"
        '
        ldFechaProceso1 = CDate("31/12/" & Trim(Str(Year(mda_FechaFin))))
        ldUltimaFechaProcesada = CDate("01/01/1990")
'        Set oRsTmp9 = mo_ReglasCaja.FactPartidasPresupuestalesXMesSelecionaUltimoProceso(ldFechaProceso1, oConexion)
'        If oRsTmp9.RecordCount > 0 Then
'           ldUltimaFechaProcesada = CDate(Format(oRsTmp9!fecha, sighEntidades.DevuelveFechaSoloFormato_DMY))
'        End If
        '
        
        
        mo_sighProxies.ReportePartidaREsumenConNotaCredito mrs_Tmp, mda_FechaInicio, mda_FechaFin, ml_IdPartidaFiltro, _
                                            sghSoloParaReporte, oConexion, True, lnIdCajero, _
                                            lnFacturasTotal, lnFacturasAnuladas, lnBoletasTotal, _
                                            lnBoletasAnuladas, False
        
        '*********************************notas de credito (inicio)*******************************
        mo_sighProxies.ReportePartidaREsumenConNotaCredito oRsNotaCredito, mda_FechaInicio, mda_FechaFin, ml_IdPartidaFiltro, _
                                                sghSoloParaReporte, oConexion, True, lnIdCajero, _
                                                lnFacturasTotal, lnFacturasAnuladas, lnBoletasTotal, _
                                                lnBoletasAnuladas, True
        If ml_IdPartidaFiltro = 0 Then
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               oRsNotaCredito.MoveFirst
               oRsNotaCredito.Find "idPartida=" & mrs_Tmp!idPartida
               If Not oRsNotaCredito.EOF Then
                    'mrs_Tmp!ImpNormal = mrs_Tmp!ImpNormal - oRsNotaCredito!ImpCancelado
                    mrs_Tmp!ImpCancelado = mrs_Tmp!ImpCancelado - oRsNotaCredito!ImpCancelado
                    mrs_Tmp!ImpNormal = oRsNotaCredito!ImpCancelado
                    mrs_Tmp!ImpNotaCredito = oRsNotaCredito!ImpCancelado
                    mrs_Tmp.Update
               End If
               mrs_Tmp.MoveNext
            Loop
            
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               mrs_Tmp!ImpAnulado = Round(mrs_Tmp!ImpAnulado, 2)
               mrs_Tmp!ImpExonerado = Round(mrs_Tmp!ImpExonerado, 2)
               mrs_Tmp!ImpNormal = Round(mrs_Tmp!ImpNormal, 2)
               mrs_Tmp!ImpCancelado = Round(mrs_Tmp!ImpCancelado, 2)
               mrs_Tmp!ImpNotaCredito = Round(mrs_Tmp!ImpNotaCredito, 2)
               mrs_Tmp.Update
               mrs_Tmp.MoveNext
            Loop
            
            
            
            
            
        Else
            If oRsNotaCredito.RecordCount > 0 Then
               oRsNotaCredito.MoveFirst
               Do While Not oRsNotaCredito.EOF
                  lnEsNuevo = True
                  If mrs_Tmp.RecordCount > 0 Then
                     mrs_Tmp.MoveFirst
                     mrs_Tmp.Find "identificador=" & oRsNotaCredito!identificador
                     If Not mrs_Tmp.EOF Then
                        lnEsNuevo = False
                     End If
                  End If
                  If lnEsNuevo = True Then
                     mrs_Tmp.AddNew
                     mrs_Tmp!identificador = oRsNotaCredito!identificador
                     mrs_Tmp!Codigo = oRsNotaCredito!Codigo
                     mrs_Tmp!descripcion = oRsNotaCredito!descripcion
                     mrs_Tmp!ImpExonerado = -oRsNotaCredito!ImpExonerado
                     mrs_Tmp!ImpNormal = -oRsNotaCredito!ImpCancelado
                     mrs_Tmp!ImpCancelado = -oRsNotaCredito!ImpCancelado
                     mrs_Tmp.Update
                  Else
                     mrs_Tmp!ImpExonerado = mrs_Tmp!ImpExonerado - oRsNotaCredito!ImpExonerado
                     mrs_Tmp!ImpNormal = mrs_Tmp!ImpNormal - oRsNotaCredito!ImpCancelado
                     mrs_Tmp!ImpCancelado = mrs_Tmp!ImpCancelado - oRsNotaCredito!ImpCancelado
                     mrs_Tmp.Update
                  End If
                  oRsNotaCredito.MoveNext
               Loop
            End If
        End If
        '*********************************notas de credito (fin)*******************************
        If mrs_Tmp.RecordCount > 0 Then
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               If mrs_Tmp!ImpCancelado = 0 And mrs_Tmp!ImpAnulado = 0 Then
                     'mrs_Tmp.Delete
               Else
                     mrs_Tmp!ImpSubTotal = mrs_Tmp!ImpCancelado + (mrs_Tmp!ImpAnulado + mrs_Tmp!ImpExonerado)
               End If
               mrs_Tmp.Update
               mrs_Tmp.MoveNext
            Loop
        End If
        
        
actualizaParametros "paso proceso largo"
        'oConexion.Close
        '*****************************************************************
        'Decisiòn para la impresion de partidas mayores a cero
        'A.Yañez 26/09/2014
        '*****************************************************************
        If mrs_Tmp.RecordCount > 0 Then
            If lbEnValor = True Then
                mrs_Tmp.MoveFirst
                Do While Not mrs_Tmp.EOF
                   If mrs_Tmp.Fields!ImpCancelado = 0 And mrs_Tmp.Fields!ImpNotaCredito = 0 Then
                      mrs_Tmp.Delete
                   End If
                   mrs_Tmp.MoveNext
                Loop
                vmrs_Tmp = 1
            Else
                mrs_Tmp.MoveFirst
                Do While Not mrs_Tmp.EOF
                   vmrs_Tmp = vmrs_Tmp + mrs_Tmp.Fields!ImpCancelado
                   mrs_Tmp.MoveNext
                Loop
            End If
        End If
        '***************************************************************
actualizaParametros "antes de existir datos        "
        If mrs_Tmp.RecordCount = 0 Then
             MsgBox "No existe Datos", vbInformation, "Reporte"
        Else
             If lbEnExcel = True Then
                 mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
             Else
actualizaParametros "antes datareport             "
                 Set RpEPartidaResumen.DataSource = mrs_Tmp
                 RpEPartidaResumen.Sections("cabecera").Controls("lblEESS").Caption = lcBuscaParametro.SeleccionaFilaParametro(205)
                 RpEPartidaResumen.Sections("cabecera").Controls("lblEESSdireccion").Caption = lcBuscaParametro.SeleccionaFilaParametro(206)
                 RpEPartidaResumen.Sections("cabecera").Controls("lblEESStelefono").Caption = "TELEFONO: " & lcBuscaParametro.SeleccionaFilaParametro(207)
                 RpEPartidaResumen.Sections("cabecera").Controls("lblFecha").Caption = "Fecha: " & lcBuscaParametro.RetornaFechaServidorSQL
                 RpEPartidaResumen.Sections("cabecera").Controls("lblHora").Caption = "Hora: " & lcBuscaParametro.RetornaHoraServidorSQL
                 
                 RpEPartidaResumen.Sections("cabecera").Controls("lblPc").Caption = "PC: " & sighentidades.RetornaNombrePC
                 RpEPartidaResumen.Sections("cabecera").Controls("lblUsuario").Caption = "Usuario: " & lcBuscaParametro.RetornaLoginUsuario(sighentidades.Usuario)
                 RpEPartidaResumen.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
                 RpEPartidaResumen.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
                 RpEPartidaResumen.Sections("cabecera").Controls("lblFecha").Caption = lcBuscaParametro.RetornaFechaHoraServidorSQL
                 RpEPartidaResumen.Sections("cabecera").Controls("Etiqueta5").Caption = "Not.Crédito"
actualizaParametros "paso cabecera                 "
                 
                 
                 Set RpEPartidaResumen.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\logotipo.jpg")
                 RpEPartidaResumen.Sections("pie").Controls("lblPie").Caption = lcPie
                 RpEPartidaResumen.RightMargin = 100
                 RpEPartidaResumen.TopMargin = 100
                 RpEPartidaResumen.LeftMargin = 100
                 RpEPartidaResumen.BottomMargin = 100
                 RpEPartidaResumen.Orientation = rptOrientDefault
                 RpEPartidaResumen.Show 1   'RpEPartidaResumen.PrintReport True
actualizaParametros "paso show"

             End If
        End If
        oConexion.Close
        Set mo_ReglasCaja = Nothing
        Set lcBuscaParametro = Nothing
        Set mrs_Tmp = Nothing
        Set mo_ReglasReportes = Nothing
        Set oRsTmp9 = Nothing
        Set oConexion = Nothing
        Set oRsCtas = Nothing
        Set oRsNotaCredito = Nothing
End Sub
'debb-12/07/2019
Sub actualizaParametros(lcMensaje As String)
'    Dim oRsTmp99988 As New Recordset
'    oRsTmp99988.Open "update parametros set valorTexto='" & lcMensaje & "' where idparametro=7", sighEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
'    Set oRsTmp99988 = Nothing
End Sub




