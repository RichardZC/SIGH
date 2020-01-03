VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptCatalogoServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub ListaServiciosEnGeneral()
    Dim oExcel As Excel.Application
    Dim oWorkBookPlantilla As Workbook
    Dim oWorkBook As Workbook
    Dim oWorkSheet As Worksheet
    Dim iFila As Integer
    Dim lnTotal As Long
    Dim rsReporte As New Recordset
    Dim rsReporte1 As New Recordset
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_ReporteUtil As New ReporteUtil
    Dim lnCant As Long, lnCol As Integer
    Dim lcNombre As String, lcSql As String
    
    rsReporte.Open "select * from FactCatalogoServicios order by Nombre", SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    If rsReporte.RecordCount > 0 Then
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HerrListaServiciosMedicamentos.xls")
        oWorkBookPlantilla.Worksheets("listaServiciosMedicamentos").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
        oWorkSheet.Cells(2, 3).Value = "RELACION DE SERVICIOS "
        iFila = 6
        lnTotal = 0
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
            lcNombre = rsReporte.Fields!Nombre
            lnCant = 1
            Do While Not rsReporte.EOF And lcNombre = rsReporte.Fields!Nombre
                If lnCant <= 1 Then
                    oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields!Codigo
                    oWorkSheet.Cells(iFila, 3).Value = rsReporte.Fields!Nombre
                    
                    '
                    lcSql = "SELECT     dbo.TiposFinanciamiento.Descripcion, dbo.FactCatalogoServiciosHosp.PrecioUnitario, " & _
                            "           dbo.FactCatalogoServiciosHosp.IdProducto" & _
                            " FROM      dbo.FactCatalogoServiciosHosp LEFT OUTER JOIN" & _
                            "           dbo.TiposFinanciamiento ON dbo.FactCatalogoServiciosHosp.IdTipoFinanciamiento = dbo.TiposFinanciamiento.IdTipoFinanciamiento" & _
                            " WHERE     dbo.FactCatalogoServiciosHosp.IdProducto=" & rsReporte.Fields!idProducto & _
                            " ORDER BY dbo.TiposFinanciamiento.Descripcion"
                    rsReporte1.Open lcSql, SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
                    lnCol = 6
                    If rsReporte1.RecordCount > 0 Then
                       rsReporte1.MoveFirst
                       Do While Not rsReporte1.EOF
                          oWorkSheet.Cells(iFila, lnCol).Value = Trim(rsReporte1.Fields!Descripcion) & " : " & Trim(Str(rsReporte1.Fields!PrecioUnitario))
                          lnCol = lnCol + 1
                          rsReporte1.MoveNext
                       Loop
                    End If
                    rsReporte1.Close
                    '
                    
                    iFila = iFila + 1
                    lnTotal = lnTotal + 1
                End If
                lnCant = lnCant + 1
                rsReporte.MoveNext
                If rsReporte.EOF Then
                   Exit Do
                End If
            Loop
        Loop
        iFila = iFila + 1
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 7
        oWorkSheet.Cells(iFila, 2).Value = "Nro: "
        oWorkSheet.Cells(iFila, 3).Value = Format(lnTotal, "####,###")
        
        oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
        If oWorkSheet.PageSetup.PrintArea <> "" Then
            oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
        End If
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
    End If
    'rsReporte.Close
End Sub




Sub ListaServiciosPorPuntosDeCarga()
    Dim oExcel As Excel.Application
    Dim oWorkBookPlantilla As Workbook
    Dim oWorkBook As Workbook
    Dim oWorkSheet As Worksheet
    Dim iFila As Integer
    Dim lnTotal As Long
    Dim rsReporte As New Recordset
    Dim rsReporte1 As New Recordset
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim mo_ReporteUtil As New ReporteUtil
    Dim lnIdPuntoCarga As Long
    Dim lcNombre As String
    Dim lcSql As String
    Dim lnCol As Integer
    lcSql = "SELECT      dbo.FactPuntosCarga.Descripcion, dbo.FactCatalogoServicios.Codigo, dbo.FactCatalogoServicios.Nombre, " & _
                "                      dbo.FactCatalogoServicios.EsCPT, dbo.FactCatalogoServiciosPtos.idPuntoCarga,dbo.FactCatalogoServicios.idProducto" & _
                " FROM         dbo.FactCatalogoServiciosPtos LEFT OUTER JOIN" & _
                "                      dbo.FactCatalogoServicios ON dbo.FactCatalogoServiciosPtos.idProducto = dbo.FactCatalogoServicios.IdProducto LEFT OUTER JOIN" & _
                "                      dbo.FactPuntosCarga ON dbo.FactCatalogoServiciosPtos.idPuntoCarga = dbo.FactPuntosCarga.IdPuntoCarga" & _
                " ORDER BY dbo.FactCatalogoServiciosPtos.idPuntoCarga, dbo.FactCatalogoServicios.Nombre"
    rsReporte.Open lcSql, SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
    If rsReporte.RecordCount > 0 Then
        'Crea nueva hoja
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HerrListaServiciosMedicamentos.xls")
        oWorkBookPlantilla.Worksheets("listaServiciosMedicamentos").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
        oWorkSheet.Cells(2, 3).Value = "RELACION DE SERVICIOS POR PUNTO DE CARGA"
        iFila = 6
        lnTotal = 0
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
            lcNombre = rsReporte.Fields!Descripcion
            oWorkSheet.Cells(iFila, 2).Value = lcNombre
            mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 7
            iFila = iFila + 1
            lnIdPuntoCarga = rsReporte.Fields!idPuntoCarga
            Do While Not rsReporte.EOF And lnIdPuntoCarga = rsReporte.Fields!idPuntoCarga
                oWorkSheet.Cells(iFila, 2).Value = rsReporte.Fields!Codigo
                oWorkSheet.Cells(iFila, 3).Value = IIf(rsReporte.Fields!EsCPT = 0, "(Insumo) ", "") & rsReporte.Fields!Nombre
                lcSql = "SELECT     dbo.TiposFinanciamiento.Descripcion, dbo.FactCatalogoServiciosHosp.PrecioUnitario, " & _
                        "           dbo.FactCatalogoServiciosHosp.IdProducto" & _
                        " FROM      dbo.FactCatalogoServiciosHosp LEFT OUTER JOIN" & _
                        "           dbo.TiposFinanciamiento ON dbo.FactCatalogoServiciosHosp.IdTipoFinanciamiento = dbo.TiposFinanciamiento.IdTipoFinanciamiento" & _
                        " WHERE     dbo.FactCatalogoServiciosHosp.IdProducto=" & rsReporte.Fields!idProducto & _
                        " ORDER BY dbo.TiposFinanciamiento.Descripcion"
                rsReporte1.Open lcSql, SIGHEntidades.CadenaConexionShape, adOpenKeyset, adLockOptimistic
                lnCol = 6
                If rsReporte1.RecordCount > 0 Then
                   rsReporte1.MoveFirst
                   Do While Not rsReporte1.EOF
                      oWorkSheet.Cells(iFila, lnCol).Value = Trim(rsReporte1.Fields!Descripcion) & " : " & Trim(Str(rsReporte1.Fields!PrecioUnitario))
                      lnCol = lnCol + 1
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

        oWorkSheet.PageSetup.PrintTitleRows = "$1:$5"
        If oWorkSheet.PageSetup.PrintArea <> "" Then
           oWorkSheet.PageSetup.PrintArea = SIGHEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila)
        End If
        If lcBuscaParametro.SeleccionaFilaParametro(284) = "S" Then
            oWorkSheet.SaveAs lcBuscaParametro.SeleccionaFilaParametro(269)
            oWorkBook.Close
            MsgBox "Se generó el archivo " & lcBuscaParametro.SeleccionaFilaParametro(269)
        Else
            oExcel.Visible = True
            oWorkSheet.PrintPreview
        End If
        'oWorkSheet.PrintOut
    Else
        MsgBox "No existen datos", vbCritical, "Mensaje"
    End If
    'rsReporte.Close
End Sub

