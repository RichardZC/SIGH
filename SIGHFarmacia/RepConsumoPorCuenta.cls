VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepConsumoPorCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para Reporte Consumo por Cuenta
'        Programado por: Barrantes D
'        Fecha: Febrero 2009
'
'------------------------------------------------------------------------------------
'***************daniel barrantes**************
'***************Filtro de Datos, Configuración y Emisión del Reporte
'***************saldos por almacen
'***************Consideraciones:
Option Explicit
Dim mo_ReporteUtil As New sighentidades.ReporteUtil
Dim ml_OrdenFiltro As String
Dim mda_FechaInicio As Date
Dim mda_FechaFin As Date
Dim ml_TextoDelFiltro  As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_IdCuentaAtencion As Long
Dim ml_idUsuario  As Long
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property

Property Let FechaInicio(daValue As Date)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As Date)
    mda_FechaFin = daValue
End Property
Property Let CuentaAtencion(lValue As Long)
    ml_IdCuentaAtencion = lValue
End Property
Property Let OrdenFiltro(lValue As String)
    ml_OrdenFiltro = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New rConsumoPorCuenta
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub

Sub ReporteXformaFarmaceutica(lnIdCuentaAtencion As Long, lcTitulo As String, lnHwnd As Long)
Dim rsReporte As New Recordset
Dim NroHistoria As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim iFila As Integer
Dim mo_ReglasSISgalenhos As New SIGHSis.ReglasSISgalenhos
Dim oRsTmp1 As New Recordset
Dim lcFF As String, lnCantidadTotal As Long
Dim lbEsOpenOffice As Boolean
Dim lcSql As String
'lnHwnd = Me.hwnd

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorError

    If lbEsOpenOffice = True Then
        Dim ServiceManager As Object
        Dim Desktop As Object
        Dim Document As Object
        Dim Feuille As Object
        Dim Plage As Object
        Dim args()
        Dim Chemin As String
        Dim Fichier As String
        Dim lcArchivoExcel As String
        Dim PrintArea(0)
        Dim Style As Object
        Dim Border As Object
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    
    If lbEsOpenOffice = True Then
        'Abre el archivo ExcelOpenOffice
        lcArchivoExcel = App.Path + "\Plantillas\HojaLibre.ods"
'        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'        Chemin = "file:///" & App.Path & "\Plantillas\"
'        Chemin = Replace(Chemin, "\", "/")
'        Fichier = Chemin & "/OpenOffice.ods"
        Fichier = Format(Time, "hhmmss") & ".ods"
        FileCopy lcArchivoExcel, App.Path + "\Plantillas\" & Fichier
        lcArchivoExcel = Fichier
        Chemin = "file:///" & App.Path & "\Plantillas\"
        Chemin = Replace(Chemin, "\", "/")
        Fichier = Chemin & "/" & lcArchivoExcel
        '
        Set ServiceManager = CreateObject("com.sun.star.ServiceManager")
        Set Desktop = ServiceManager.createInstance("com.sun.star.frame.Desktop")
        Set Document = Desktop.loadComponentFromURL(Fichier, "_blank", 0, args)
        Set Feuille = Document.getSheets().getByIndex(0)
        'Encabezado de Pagina
        mo_CabeceraReportes.CabeceraReportes Document, True
        ' Pone la ventana en primer plano, pasándole el Hwnd
        ret = SetForegroundWindow(lnHwnd)
    Else
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
        oWorkBookPlantilla.Worksheets("Hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
    End If
    Set oRsTmp1 = mo_ReglasSISgalenhos.SISDespachosFarmaciaIncluyeDevoluciones(lnIdCuentaAtencion)
    oRsTmp1.Sort = "FormaFarmaceutica"
    If oRsTmp1.RecordCount = 0 Then
        MsgBox "No hay Datos", vbInformation, "Reporte"
    Else
        iFila = 1
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, iFila - 1).setFormula("Listado x Forma Farmacéutica " & lcTitulo)
            Else
                oWorkSheet.Cells(iFila, 1).Value = "Listado x Forma Farmacéutica " & lcTitulo
            End If
        iFila = iFila + 2
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(0, iFila - 1).setFormula("Forma Farmacéutica")
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula("Cantidad")
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula("Documento")
                Call Feuille.getcellbyposition(3, iFila - 1).setFormula("Fecha")
                Call Feuille.getcellbyposition(4, iFila - 1).setFormula("Código")
                Call Feuille.getcellbyposition(5, iFila - 1).setFormula("Nombre")
            Else
                oWorkSheet.Cells(iFila, 1).Value = "Forma Farmacéutica"
                oWorkSheet.Cells(iFila, 2).Value = "Cantidad"
                oWorkSheet.Cells(iFila, 3).Value = "Documento"
                oWorkSheet.Cells(iFila, 4).Value = "Fecha"
                oWorkSheet.Cells(iFila, 5).Value = "Código"
                oWorkSheet.Cells(iFila, 6).Value = "Nombre"
            End If
        iFila = iFila + 2
        oRsTmp1.Sort = "FormaFarmaceutica"
        
        If oRsTmp1.RecordCount > 0 Then
            oRsTmp1.MoveFirst
            Do While Not oRsTmp1.EOF
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(0, iFila - 1).setFormula(oRsTmp1!FormaFarmaceutica)
                Else
                    oWorkSheet.Cells(iFila, 1).Value = oRsTmp1!FormaFarmaceutica
                End If
               lcFF = oRsTmp1.Fields!FormaFarmaceutica
               lnCantidadTotal = 0
               Do While Not oRsTmp1.EOF And lcFF = oRsTmp1.Fields!FormaFarmaceutica
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(oRsTmp1!Cantidad)
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(oRsTmp1!DocumentoNumero)
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula(Format(oRsTmp1!fechaCreacion, "dd/mm/yyyy hh:mm"))
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula(oRsTmp1!codigo)
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(oRsTmp1!Nombre)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = oRsTmp1!Cantidad
                    oWorkSheet.Cells(iFila, 3).Value = oRsTmp1!DocumentoNumero
                    oWorkSheet.Cells(iFila, 4).Value = Format(oRsTmp1!fechaCreacion, "dd/mm/yyyy hh:mm")
                    oWorkSheet.Cells(iFila, 5).Value = oRsTmp1!codigo
                    oWorkSheet.Cells(iFila, 6).Value = oRsTmp1!Nombre
                End If
                  iFila = iFila + 1
                  lnCantidadTotal = lnCantidadTotal + oRsTmp1.Fields!Cantidad
                  oRsTmp1.MoveNext
                    If oRsTmp1.EOF Then
                       Exit Do
                    End If
               Loop
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula("___________________________________")
                Else
                    oWorkSheet.Cells(iFila, 2).Value = "___________________________________"
                End If
               iFila = iFila + 1
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, iFila - 1).setFormula(lnCantidadTotal)
                Else
                    oWorkSheet.Cells(iFila, 2).Value = lnCantidadTotal
                End If
               iFila = iFila + 2
            Loop
        
        End If
        If lbEsOpenOffice = True Then
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oExcel.Visible = True
            oWorkSheet.PrintPreview
        End If
        If lbEsOpenOffice = True Then
           Set Plage = Nothing
           Set Feuille = Nothing
           Set Document = Nothing
           Set Desktop = Nothing
           Set ServiceManager = Nothing
           Set Style = Nothing
           Set Border = Nothing
           'encabezado de pagina
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
        Else
            oWorkBook.Close SaveChanges:=False
            Set mo_ReglasSISgalenhos = Nothing
            Set oRsTmp1 = Nothing
            Set oWorkSheet = Nothing
            Set oExcel = Nothing
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
        End If
    End If
Exit Sub
ManejadorError:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub

End Sub

Function ProcesaConsumoXcuenta(ml_IdCuenta As Long, ByRef lnTotal As Double, lbDesdeVentas As Boolean, _
                               lbConsideraItemsDePaquetes As Boolean) As Recordset
        Dim oConexion As New Connection
        Dim rsReporte As New Recordset
        Dim mrs_Tmp As New Recordset
        Dim mrs_Tmp1 As New Recordset
        Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
        Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
        Dim lnTotalRegistros As Long, lnPrecio As Double, lcTexto1 As String
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set rsReporte = mo_ReglasFarmacia.FarmMovimientoVentasDetalleSeleccionarPorCuenta(ml_IdCuenta, oConexion)
        rsReporte.Filter = "idFuenteFinanciamiento<>1"
        lnTotalRegistros = rsReporte.RecordCount
        lnPrecio = 0
        With mrs_Tmp
              If lbDesdeVentas = True Then
                 .Fields.Append "Tipo", adVarChar, 30, adFldIsNullable
              End If
              .Fields.Append "FechaCreacion", adDate, 10, adFldIsNullable
              .Fields.Append "HoraCreacion", adVarChar, 5, adFldIsNullable
              .Fields.Append "MovNumero", adVarChar, 15, adFldIsNullable
              .Fields.Append "codigo", adVarChar, 20, adFldIsNullable
              .Fields.Append "Nombre", adVarChar, 150, adFldIsNullable
              .Fields.Append "cantidad", adInteger, 4, adFldIsNullable
              .Fields.Append "Precio", adDouble
              .Fields.Append "Total", adDouble
              .Fields.Append "Estado", adVarChar, 20, adFldIsNullable
              .Fields.Append "dAlmacen", adVarChar, 100, adFldIsNullable
              .Fields.Append "dFinanciamiento", adVarChar, 50, adFldIsNullable
              .Fields.Append "Usuario", adVarChar, 30, adFldIsNullable
              .Fields.Append "esPaquete", adBoolean
              If lbDesdeVentas = False Then
                 .Fields.Append "Tipo", adVarChar, 30, adFldIsNullable
              End If
              .LockType = adLockOptimistic
              .Open
        End With
        mrs_Tmp.Sort = "fechaCreacion desc,HoraCreacion desc,nombre"
        If lnTotalRegistros = 0 Then
            If lbDesdeVentas = False Then
                MsgBox "No hay consumos para esa Cuenta", vbInformation, "Consumo por Cuenta"
            End If
        Else
            rsReporte.MoveFirst
            
            Do While Not rsReporte.EOF
               lcTexto1 = ""
               Set mrs_Tmp1 = mo_ReglasFacturacion.FacturacionBienesPagosSeleccionarPorMovNumeroProducto(rsReporte.Fields!movNumero, "S", rsReporte.Fields!idProducto, oConexion)
               If mrs_Tmp1.RecordCount > 0 Then
                  If mrs_Tmp1.Fields!idEstadoFacturacion = 4 And mrs_Tmp1.Fields!idComprobantepago > 0 Then
                     lcTexto1 = "Pago"
                  End If
               End If
               mrs_Tmp.AddNew
               mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
               mrs_Tmp.Fields!HoraCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveHoraSoloFormato_HM)
               mrs_Tmp.Fields!movNumero = rsReporte.Fields!DocumentoNumero
               mrs_Tmp.Fields!codigo = rsReporte.Fields!codigo
               mrs_Tmp.Fields!Nombre = rsReporte.Fields!Nombre
               mrs_Tmp.Fields!Cantidad = rsReporte.Fields!Cantidad
               mrs_Tmp.Fields!Precio = rsReporte.Fields!Precio
               mrs_Tmp.Fields!total = rsReporte.Fields!total
               mrs_Tmp.Fields!Estado = rsReporte.Fields!Estado
               mrs_Tmp.Fields!dAlmacen = rsReporte.Fields!dAlmacen
               mrs_Tmp.Fields!dFinanciamiento = rsReporte.Fields!dFinanciamiento
               mrs_Tmp.Fields!Usuario = rsReporte.Fields!Usuario
               mrs_Tmp.Fields!tipo = IIf(lcTexto1 = "", "Venta", "Pagó")
               mrs_Tmp.Update
               If lcTexto1 = "" Then
                  lnPrecio = lnPrecio + rsReporte.Fields!total
               End If
               rsReporte.MoveNext
            Loop
            'el CODIGO DEL ITEM es un Paquete de Farmacia,se desagrega en CODIGOS DIGEMID   'debb-08/11/2016
'            If mrs_Tmp.RecordCount > 0 And lbConsideraItemsDePaquetes = True Then
'
'                rsItems.MoveFirst
'                Do While Not rsItems.EOF
'                   lcCodigo = rsItems!Codigo
'                   If rsItems!esPaquete = True Then
'                      lnCantidadBolsas = rsItems!cantidad
'                      Set oRsPqte = mo_ReglasFarmacia.CatalogoDIGEMIDdevuelveITEMS(lcCodigo)
'                      If oRsPqte.RecordCount > 0 Then
'                         oRsPqte.MoveFirst
'                         Do While Not oRsPqte.EOF
'                            lnPrecio = oRsPqte!Precio
'                            rsItems.MoveFirst
'                            rsItems.Find "codigo='" & oRsPqte!Codigo & "'"
'                            If rsItems.EOF Then
'                               lnCantidad = oRsPqte!cantidad * lnCantidadBolsas
'                               lnTotal = Round(lnCantidad * lnPrecio, 2)
'                               rsItems.AddNew
'                               rsItems.Fields!PuntoCarga = "Farmacia"
'                               rsItems.Fields!Codigo = oRsPqte!Codigo
'                               rsItems.Fields!Item = Left(oRsPqte!Nombre, 200)
'                               rsItems.Fields!Precio = lnPrecio
'                            Else
'                               lnCantidad = rsItems!cantidad + (oRsPqte!cantidad * lnCantidadBolsas)
'                               lnTotal = rsItems!Total + Round((oRsPqte!cantidad * lnCantidadBolsas) * lnPrecio, 2)
'                            End If
'                            rsItems.Fields!cantidad = lnCantidad
'                            rsItems.Fields!Total = lnTotal
'                            rsItems.Update
'                            oRsPqte.MoveNext
'                         Loop
'                      End If
'                   End If
'                   rsItems.MoveNext
'                Loop
'                rsItems.MoveFirst
'                Do While Not rsItems.EOF
'                   lcCodigo = rsItems!Codigo
'                   If rsItems!esPaquete = True Then
'                      rsItems.Delete
'                      rsItems.Update
'                   End If
'                   rsItems.MoveNext
'                Loop
'
'
'            End If
            'Devoluciones
            Set rsReporte = mo_ReglasFarmacia.FarmMovimientoNotaIngresoSeleccionarPorCuenta(ml_IdCuenta)
            rsReporte.Filter = "idTipoConcepto=21"    'solo DEVOLUCIONES del PACIENTE
            lnTotalRegistros = rsReporte.RecordCount
            
            If lnTotalRegistros > 0 Then
               Do While Not rsReporte.EOF
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!fechaCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveFechaSoloFormato_DMY)
                    mrs_Tmp.Fields!HoraCreacion = Format(rsReporte.Fields!fechaCreacion, sighentidades.DevuelveHoraSoloFormato_HM)
                    mrs_Tmp.Fields!movNumero = rsReporte.Fields!movNumero
                    mrs_Tmp.Fields!codigo = rsReporte.Fields!codigo
                    mrs_Tmp.Fields!Nombre = rsReporte.Fields!Nombre
                    mrs_Tmp.Fields!Cantidad = rsReporte.Fields!Cantidad
                    mrs_Tmp.Fields!Precio = rsReporte.Fields!Precio
                    mrs_Tmp.Fields!total = rsReporte.Fields!total
                    mrs_Tmp.Fields!Estado = rsReporte.Fields!Estado
                    mrs_Tmp.Fields!dAlmacen = rsReporte.Fields!dAlmacen
                    mrs_Tmp.Fields!dFinanciamiento = rsReporte.Fields!dFinanciamiento
                    mrs_Tmp.Fields!Usuario = rsReporte.Fields!Usuario
                    mrs_Tmp.Fields!tipo = "Devol."
                    mrs_Tmp.Update
                    If rsReporte.Fields!idEstadoMovimiento = 1 Then
                       lnPrecio = lnPrecio - rsReporte.Fields!total
                    End If
                    rsReporte.MoveNext
                Loop
            End If
            If mrs_Tmp.RecordCount > 0 Then
               
               mrs_Tmp.MoveFirst
            End If
        End If
        lnTotal = lnPrecio
        Set ProcesaConsumoXcuenta = mrs_Tmp
        Set oConexion = Nothing
        Set rsReporte = Nothing
        Set mrs_Tmp = Nothing
        Set mrs_Tmp1 = Nothing
        Set mo_ReglasFarmacia = Nothing
        Set mo_ReglasFacturacion = Nothing
End Function






