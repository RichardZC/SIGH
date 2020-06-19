VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clCatalogoServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para catálogo de Servicios
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_reglasComunes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes


Sub ListaServiciosEnGeneral(lnHwnd As Long)
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lbEsOpenOffice As Boolean
    lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False) 'lbEsOpenOffice = lcBuscaParametro.EsOpenOffice
'    If MsgBox("¿Usarà OPEN OFFICE?", vbYesNo, "") = vbYes Then
'       'ListaServiciosEnGeneralOpenOffice
'       'Exit Sub
'       lbEsOpenOffice = True
'    End If
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
        'encabezado
        Dim PageStyles As Object
        Dim Sheet As Object
        Dim StyleFamilies As Object
        Dim DefPage As Object
        Dim Htext As Object
        Dim Hcontent As Object
        '
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    Dim iFila As Long
    Dim lnTotal As Long
    Dim rsReporte As New Recordset
    Dim rsReporte1 As New Recordset
    Dim oRsTarifas As New Recordset
    Dim mo_ReporteUtil As New ReporteUtil
    Dim lnCant As Long, lnCol As Integer
    Dim lcNombre As String, lcSql As String
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHEntidades.CadenaConexion
    '
    With oRsTarifas
          .Fields.Append "id", adInteger, 4, adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 200, adFldIsNullable
          .Fields.Append "Columna", adInteger
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set rsReporte = mo_reglasComunes.TiposFinanciamientoSegunFiltro("seIngresPrecios=1 and idTipoFinanciamiento>0")
    If rsReporte.RecordCount > 0 Then
       lnCol = 4
       rsReporte.MoveFirst
       Do While Not rsReporte.EOF
          oRsTarifas.AddNew
          oRsTarifas.Fields!ID = rsReporte.Fields!idTipoFinanciamiento
          oRsTarifas.Fields!descripcion = rsReporte.Fields!descripcion
          oRsTarifas.Fields!columna = lnCol
          oRsTarifas.Update
          lnCol = lnCol + 1
          rsReporte.MoveNext
       Loop
    End If
    rsReporte.Close
    
    '
    Set rsReporte = mo_reglasComunes.CatalogoServiciosSeleccionarTodos(oConexion)
    
    If rsReporte.RecordCount > 0 Then
          
        If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HerrListaServiciosMedicamentos.ods"
            '
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
            mo_CabeceraReportes.CabeceraReportes Document, True
          ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHwnd)
        Else
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HerrListaServiciosMedicamentos.xls")
            oWorkBookPlantilla.Worksheets("listaServiciosMedicamentos").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        '
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, 1).setFormula("RELACION DE SERVICIOS (CPT)")
        Else
            oWorkSheet.Cells(2, 2).Value = "RELACION DE SERVICIOS (CPT)"

        End If
        '
        oRsTarifas.MoveFirst
        Do While Not oRsTarifas.EOF
           If lbEsOpenOffice = True Then
              Call Feuille.getcellbyposition(oRsTarifas!columna - 1, 4).setFormula(oRsTarifas!descripcion)
           Else
              oWorkSheet.Cells(5, oRsTarifas!columna).Value = oRsTarifas!descripcion
           End If
           oRsTarifas.MoveNext
        Loop
        '
'        iFila = 6
'        lnTotal = 0
'        rsReporte.MoveFirst
'        Do While Not rsReporte.EOF
'            lcNombre = rsReporte.Fields!NombreProducto
'            lnCant = 1
'            Do While Not rsReporte.EOF And lcNombre = rsReporte.Fields!NombreProducto
'                If lnCant <= 1 Then
'                    If lbEsOpenOffice = True Then
'                        lcSql = rsReporte.Fields!Codigo
'                        Call Feuille.getcellbyposition(1, iFila).setFormula(lcSql)
'                        lcSql = Trim(rsReporte.Fields!NombreProducto)
'                        Call Feuille.getcellbyposition(2, iFila).setFormula(lcSql)
'                    Else
'                        oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields!Codigo
'                        oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields!NombreProducto
'                    End If
                    '
                    
'                    Set rsReporte1 = mo_reglasComunes.FactCatalogoServiciosHospPorIdProducto(rsReporte.Fields!idProducto, oConexion)
'                    lnCol = 6
'                    If rsReporte1.RecordCount > 0 Then
'                       rsReporte1.MoveFirst
'                       Do While Not rsReporte1.EOF
'                          If lbEsOpenOffice = True Then
'                             lcSql = Trim(rsReporte1.Fields!descripcion) & " : " & Trim(Str(rsReporte1.Fields!PrecioUnitario))
'                             Call Feuille.getcellbyposition(lnCol - 1, iFila - 1).setFormula(lcSql)
'                          Else
'                             oWorkSheet.Cells(iFila, lnCol).Value = Trim(rsReporte1.Fields!descripcion) & " : " & Trim(Str(IIf(IsNull(rsReporte1.Fields!PrecioUnitario), "", rsReporte1.Fields!PrecioUnitario)))
'                          End If
'                          lnCol = lnCol + 1
'                          rsReporte1.MoveNext
'                       Loop
'                    End If
''                    rsReporte1.Close
'                    '
'
'                    iFila = iFila + 1
'                    lnTotal = lnTotal + 1
'                End If
'                lnCant = lnCant + 1
'                rsReporte.MoveNext
'                If rsReporte.EOF Then
'                   Exit Do
'                End If
'            Loop
'        Loop

        iFila = 6
        lnTotal = 0
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
            If lbEsOpenOffice = True Then
                lcSql = rsReporte.Fields!Codigo
                Call Feuille.getcellbyposition(1, iFila).setFormula(lcSql)
                lcSql = Trim(rsReporte.Fields!NombreProducto)
                Call Feuille.getcellbyposition(2, iFila).setFormula(lcSql)
            Else
                oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields!Codigo
                oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields!NombreProducto
            End If
            lcNombre = rsReporte.Fields!NombreProducto
            Do While Not rsReporte.EOF And lcNombre = rsReporte.Fields!NombreProducto
                oRsTarifas.MoveFirst
                oRsTarifas.Find "id=" & rsReporte!idTipoFinanciamiento
                If Not oRsTarifas.EOF Then
                   If lbEsOpenOffice = True Then
                       lcSql = Trim(Str(rsReporte.Fields!PrecioUnitario))
                       Call Feuille.getcellbyposition(oRsTarifas!columna - 1, iFila - 1).setFormula(lcSql)
                   Else
                       oWorkSheet.Cells(iFila, oRsTarifas!columna).Value = rsReporte!PrecioUnitario
                   End If
                End If
                rsReporte.MoveNext
                If rsReporte.EOF Then
                   Exit Do
                End If
            Loop
            iFila = iFila + 1
            lnTotal = lnTotal + 1
        Loop
            
        iFila = iFila + 1
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, iFila).setFormula("Nro: ")
            lcSql = Format(lnTotal, "####,###")
            Call Feuille.getcellbyposition(2, iFila).setFormula(lcSql)
        Else
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 7
            oWorkSheet.Cells(iFila, 2).Value = "Nro: "
            oWorkSheet.Cells(iFila, 3).Value = Format(lnTotal, "####,###")
        End If
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 6
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
            'Liberar Memoria
            Set Plage = Nothing
            Set Feuille = Nothing
            Set Document = Nothing
            Set Desktop = Nothing
            Set ServiceManager = Nothing
            'encabezado de pagina
            Set PageStyles = Nothing
            Set Sheet = Nothing
            Set StyleFamilies = Nothing
            Set DefPage = Nothing
            Set Htext = Nothing
            Set Hcontent = Nothing
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
            If oWorkSheet.PageSetup.PrintArea <> "" Then
                oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
            End If
            oExcel.Visible = True
            oWorkSheet.PrintPreview
            'oWorkSheet.PrintOut
            'liberar memoria
            Set oExcel = Nothing
            Set oWorkBookPlantilla = Nothing
            Set oWorkBook = Nothing
            Set oWorkSheet = Nothing
        End If
        
    End If
    'rsReporte.Close
    Set oConexion = Nothing
End Sub

Sub ListaServiciosPorPuntosDeCarga(lnHwnd As Long)
    Dim iFila As Long
    Dim lnTotal As Long
    Dim rsReporte As New Recordset
    Dim rsReporte1 As New Recordset
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_ReporteUtil As New ReporteUtil
    Dim oRsTarifas As New Recordset
    Dim lnIdPuntoCarga As Long
    Dim lcNombre As String
    Dim lcSql As String
    Dim lnCol As Integer
    With oRsTarifas
          .Fields.Append "id", adInteger, 4, adFldIsNullable
          .Fields.Append "Descripcion", adVarChar, 200, adFldIsNullable
          .Fields.Append "Columna", adInteger
          .CursorType = adOpenKeyset
          .LockType = adLockOptimistic
          .Open
    End With
    Set rsReporte = mo_reglasComunes.TiposFinanciamientoSegunFiltro("seIngresPrecios=1 and idTipoFinanciamiento>0")
    If rsReporte.RecordCount > 0 Then
       lnCol = 4
       rsReporte.MoveFirst
       Do While Not rsReporte.EOF
          oRsTarifas.AddNew
          oRsTarifas.Fields!ID = rsReporte.Fields!idTipoFinanciamiento
          oRsTarifas.Fields!descripcion = rsReporte.Fields!descripcion
          oRsTarifas.Fields!columna = lnCol
          oRsTarifas.Update
          lnCol = lnCol + 1
          rsReporte.MoveNext
       Loop
    End If
    rsReporte.Close
    
    '
    
    
    Set rsReporte = mo_reglasComunes.FactCatalogoServiciosPtosOrdenasXcpt
    If rsReporte.RecordCount > 0 Then
    Dim lbEsOpenOffice As Boolean
    
    lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
    
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
        '
        Dim ret As Long
    Else
        Dim oExcel As Excel.Application
        Dim oWorkBookPlantilla As Workbook
        Dim oWorkBook As Workbook
        Dim oWorkSheet As Worksheet
    End If
    If lbEsOpenOffice = True Then
            'Abre el archivo ExcelOpenOffice
            lcArchivoExcel = App.Path + "\Plantillas\HerrListaServiciosMedicamentos.ods"
    '        FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
    '        Chemin = "file:///" & App.Path & "\Plantillas\"
    '        Chemin = Replace(Chemin, "\", "/")
    '        Fichier = Chemin & "/OpenOffice.ods"
            '
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
            mo_CabeceraReportes.CabeceraReportes Document, True
            ' Pone la ventana en primer plano, pasándole el Hwnd
            ret = SetForegroundWindow(lnHwnd)
        Else
            'Crea nueva hoja
            Set oExcel = GalenhosExcelApplication()  'New Excel.Application
            Set oWorkBook = oExcel.Workbooks.Add
            'Abre, copia y cierra la plantilla
            Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HerrListaServiciosMedicamentos.xls")
            oWorkBookPlantilla.Worksheets("listaServiciosMedicamentos").Copy Before:=oWorkBook.Sheets(1)
            oWorkBookPlantilla.Close
            'Activa la primera hoja
            Set oWorkSheet = oWorkBook.Sheets(1)
            mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
        End If
        If lbEsOpenOffice = True Then
            Call Feuille.getcellbyposition(1, 1).setFormula("RELACION DE SERVICIOS (CPT) POR PUNTO DE CARGA")
'            Call Feuille.getcellbyposition(5, 4).setFormula("CENTRO COSTOS")
'            Call Feuille.getcellbyposition(6, 4).setFormula("PARTIDA")
'            Call Feuille.getcellbyposition(7, 4).setFormula("PRECIOS")
        Else
            oWorkSheet.Cells(2, 2).Value = "RELACION DE SERVICIOS (CPT) POR PUNTO DE CARGA"
'            oWorkSheet.Cells(5, 6).Value = "CENTRO COSTOS"
'            oWorkSheet.Cells(5, 7).Value = "PARTIDA"
'            oWorkSheet.Cells(5, 8).Value = "PRECIOS"
        End If
        '
        oRsTarifas.MoveFirst
        Do While Not oRsTarifas.EOF
           If lbEsOpenOffice = True Then
              Call Feuille.getcellbyposition(oRsTarifas!columna - 1, 4).setFormula(oRsTarifas!descripcion)
           Else
              oWorkSheet.Cells(5, oRsTarifas!columna).Value = oRsTarifas!descripcion
           End If
           oRsTarifas.MoveNext
        Loop
        '
        iFila = 6
        lnTotal = 0
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
            lcNombre = rsReporte.Fields!descripcion
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(lcNombre)
            Else
                oWorkSheet.Cells(iFila, 2).Value = lcNombre
            End If
            If lbEsOpenOffice = True Then
                Set Plage = Feuille.getCellRangeByName("B" & CStr(iFila) & ":F" & CStr(iFila))
                mo_ReporteUtil.ExcelOpenOfficeCuadricularRango Plage, 50
            Else
                mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 7
            End If
            iFila = iFila + 1
            lnIdPuntoCarga = rsReporte.Fields!IdPuntoCarga
            Do While Not rsReporte.EOF And lnIdPuntoCarga = rsReporte.Fields!IdPuntoCarga
            If lbEsOpenOffice = True Then
                Call Feuille.getcellbyposition(1, iFila - 1).setFormula(rsReporte.Fields!Codigo)
                Call Feuille.getcellbyposition(2, iFila - 1).setFormula(IIf(rsReporte.Fields!EsCPT = 0, "(Insumo) ", "") & rsReporte.Fields!Nombre)
                'Call Feuille.getcellbyposition(5, iFila - 1).setFormula(IIf(IsNull(rsReporte.Fields!dCentroCosto), "", rsReporte.Fields!dCentroCosto))
                'Call Feuille.getcellbyposition(6, iFila - 1).setFormula(rsReporte.Fields!dPartida)
            Else
                oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields!Codigo
                oWorkSheet.Cells(iFila, 3).Value = IIf(rsReporte.Fields!EsCPT = 0, "(Insumo) ", "") & rsReporte.Fields!Nombre
                'oWorkSheet.Cells(iFila, 6).Value = rsReporte.Fields!dCentroCosto
                'oWorkSheet.Cells(iFila, 7).Value = rsReporte.Fields!dPartida
            End If
                Set rsReporte1 = mo_reglasComunes.FactCatalogoServiciosHospXidProducto(rsReporte.Fields!idProducto)
                lnCol = 8
                If rsReporte1.RecordCount > 0 Then
                   rsReporte1.MoveFirst
                   Do While Not rsReporte1.EOF
'                        If lbEsOpenOffice = True Then
'                            Call Feuille.getcellbyposition(lnCol - 1, iFila - 1).setFormula(Trim(rsReporte1.Fields!descripcion) & " : " & Trim(Str(rsReporte1.Fields!PrecioUnitario)))
'                        Else
'                            oWorkSheet.Cells(iFila, lnCol).Value = Trim(rsReporte1.Fields!descripcion) & " : " & Trim(Str(rsReporte1.Fields!PrecioUnitario))
'                        End If
'                        lnCol = lnCol + 1
                        oRsTarifas.MoveFirst
                        oRsTarifas.Find "id=" & rsReporte1!idTipoFinanciamiento
                        If Not oRsTarifas.EOF Then
                           If lbEsOpenOffice = True Then
                               lcSql = Trim(Str(rsReporte1.Fields!PrecioUnitario))
                               Call Feuille.getcellbyposition(oRsTarifas!columna - 1, iFila - 1).setFormula(lcSql)
                           Else
                               oWorkSheet.Cells(iFila, oRsTarifas!columna).Value = rsReporte1!PrecioUnitario
                           End If
                        End If
                        rsReporte1.MoveNext
                   Loop
                End If
                rsReporte1.Close
                '
                iFila = iFila + 1
                rsReporte.MoveNext
                If rsReporte.EOF Then
                   Exit Do
                End If
            Loop
            iFila = iFila + 1
        Loop
'        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 7
'        oWorkSheet.Cells(iFila, 2).Value = "Nro: "
'        oWorkSheet.Cells(iFila, 3).Value = Format(lnTotal, "####,###")
        If lbEsOpenOffice = True Then
            Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
            PrintArea(0).Sheet = 0
            PrintArea(0).startcolumn = 1
            PrintArea(0).StartRow = 0
            PrintArea(0).EndColumn = 6
            PrintArea(0).EndRow = iFila
            Call Feuille.SetPrintAreas(PrintArea())
            Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            Call Document.getCurrentController.getFrame.getComponentWindow.setVisible(True)
            '
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
        Else
            oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
                End If
             'SCCQ 19/06/2020 Cambio24 Inicio
                 'El siguiente codigo es para verificar si la ruta existe
                 Dim parametro_reporte As String
                 parametro_reporte = lcBuscaParametro.SeleccionaFilaParametro(269)
                 Dim strPath As String 'Variable que contiene la ruta del archivo de la carpeta donde se generan los reportes
                 Dim posicion As Integer
                 posicion = InStrRev(parametro_reporte, "\")
                 strPath = Mid(parametro_reporte, 1, posicion - 1) '"c:\Reportes" 'Ruta para genear reporte
                 If Dir(strPath, vbDirectory) = "" Then 'Si el directorio no existe
                    MkDir strPath 'Se crea la carpeta
                 End If
            'SCCQ 19/06/2020 Cambio24 Fin
            oWorkSheet.SaveAs lcBuscaParametro.SeleccionaFilaParametro(269)
            'oWorkBook.Close
            MsgBox "Se generó el archivo " & lcBuscaParametro.SeleccionaFilaParametro(269)
            oExcel.Visible = True
            oWorkSheet.PrintPreview
        End If
        'oWorkSheet.PrintOut
    Else
        MsgBox "No existen datos", vbInformation, "Mensaje"
    End If
    'rsReporte.Close
    If lbEsOpenOffice = True Then
        'Liberar Memoria
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
        'Liberar memoria
        Set oExcel = Nothing
        Set oWorkBookPlantilla = Nothing
        Set oWorkBook = Nothing
        Set oWorkSheet = Nothing
    End If
End Sub

Sub EjecutaFormulario()
    Dim oFormulario As New RpCatServicios
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub


