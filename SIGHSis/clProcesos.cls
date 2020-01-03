VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Procesos con el FUA
'        Programado por: Barrantes D
'        Fecha: Enero 2013
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_lcNombrePc As String
Dim ml_idUsuario As Long
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim ms_MensajeError As String
Dim lcSql As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
Dim lcParaMsgbox As String
 
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property


Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Function DevuelveDESDEdeTablaEdad000(lnTipoEdad As Long, lcRangoTabla As String) As Integer
    Dim lcBuscarTxt As String, lnBuscarLen As Integer, lnPosic As Integer
    Dim lcRango As String, lnFor As Integer
    lcBuscarTxt = "(" & Trim(Str(lnTipoEdad)) & "."
    lnBuscarLen = Len(lcBuscarTxt)
    lnPosic = InStr(lcRangoTabla, lcBuscarTxt)
    lcRango = ""
    For lnFor = lnPosic + lnBuscarLen To Len(lcRangoTabla)
        If Mid(lcRangoTabla, lnFor, 1) = "-" Then
           Exit For
        End If
        lcRango = lcRango + Mid(lcRangoTabla, lnFor, 1)
    Next
    DevuelveDESDEdeTablaEdad000 = Val(lcRango)
End Function

Function DevuelveHASTAdeTablaEdad000(lnTipoEdad As Long, lcRangoTabla As String) As Integer
    Dim lcBuscarTxt As String, lnBuscarLen As Integer, lnPosic As Integer
    Dim lcRango As String, lnFor As Integer
    lcBuscarTxt = "-" & Trim(Str(lnTipoEdad)) & "."
    lnBuscarLen = Len(lcBuscarTxt)
    lnPosic = InStr(lcRangoTabla, lcBuscarTxt)
    lcRango = ""
    For lnFor = lnPosic + lnBuscarLen To Len(lcRangoTabla)
        If Mid(lcRangoTabla, lnFor, 1) = ")" Then
           Exit For
        End If
        lcRango = lcRango + Mid(lcRangoTabla, lnFor, 1)
    Next
    DevuelveHASTAdeTablaEdad000 = Val(lcRango)
End Function











Private Function isValidField(obj_Field As ADODB.Field) As Boolean
    With obj_Field

        On Error GoTo error_handler
        Select Case obj_Field.Type
            Case adBinary, adIDispatch, adIUnknown, adUserDefined
                isValidField = False
            ' -- Campo válido
            Case Else
                isValidField = True
        End Select
    End With
Exit Function
error_handler:
End Function

Sub GeneraTXT_una_linea(oRsTmp1 As Recordset, lcArchivoTXT As String, lnBarraTotal As Long, lnBarraNumero As Integer)
        On Error GoTo ErrGenTXT
        Dim lniFreeFile As Integer, lnUltimoCampo As Integer, lnCont As Integer, lcLinea As String
        Dim lcCampoValor As String, lnContador As Long
        Const lcSeparadorCampo As String = "|": Dim lcSeparadorSaltoLinea As String
        lcSeparadorSaltoLinea = Chr(10)
        lniFreeFile = FreeFile
        Open lcArchivoTXT For Output As #lniFreeFile
        lnUltimoCampo = oRsTmp1.Fields.Count
        '
        lnContador = 1
        '
        oRsTmp1.MoveFirst
        lcLinea = ""
        Do While Not oRsTmp1.EOF
           lnContador = lnContador + 1
           '
           
           For lnCont = 0 To lnUltimoCampo - 1
               lcCampoValor = ""
               If Not IsNull(oRsTmp1.Fields(lnCont).Value) Then
                    Select Case oRsTmp1.Fields(lnCont).Type
                    Case adDate
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    Case adChar, adVarChar
                         lcCampoValor = Trim(oRsTmp1.Fields(lnCont))
                    Case adDouble, adInteger
                         lcCampoValor = Trim(Str((oRsTmp1.Fields(lnCont))))
                    Case Else
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    End Select
               End If
               If lnUltimoCampo = lnCont Then
                  lcLinea = lcLinea & lcCampoValor & lcSeparadorSaltoLinea
               Else
                  lcLinea = lcLinea & lcCampoValor & lcSeparadorCampo
               End If
           Next
           
           oRsTmp1.MoveNext
        Loop
        Print #lniFreeFile, lcLinea
        Close #lniFreeFile
        Exit Sub
ErrGenTXT:
        MsgBox Err.Description
        Resume
End Sub



Sub CrearReporteObservaciones_excel(mrs_Tmp As Recordset)
Dim oExcel As Excel.Application
Dim oWorkBookPlantilla As Workbook
Dim oWorkBook As Workbook
Dim oWorkSheet As Worksheet
Dim mo_ReporteUtil As New SIGHEntidades.ReporteUtil
Dim iFila As Long
Dim lnTotal As Double
        On Error GoTo ManejadorErrorExcel
        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
        'Crea nueva hoja
        Set oWorkBook = oExcel.Workbooks.Add
        'Abre, copia y cierra la plantilla
        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
        oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
        oWorkBookPlantilla.Close
        'Activa la primera hoja
        Set oWorkSheet = oWorkBook.Sheets(1)
        'oWorkSheet.PageSetup.LeftHeader = lcBuscaParametro.SeleccionaFilaParametro(205)
        iFila = 2
        oWorkSheet.Cells(iFila, 2).Value = "Lista de Cuentas Observadas (No fueron grabados sus FUAS)"
        iFila = 4
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\Imagen de reportes.jpg"
        oWorkSheet.Cells(iFila, 2).Value = "Nro Cuenta"
        oWorkSheet.Cells(iFila, 3).Value = "Apellido Paterno"
        oWorkSheet.Cells(iFila, 4).Value = "Apellido Materno"
        oWorkSheet.Cells(iFila, 5).Value = "Primer Nombre"
        oWorkSheet.Cells(iFila, 6).Value = "N° Historia Clinica"
        oWorkSheet.Cells(iFila, 7).Value = "Tipo Servicio"
        oWorkSheet.Cells(iFila, 8).Value = "Servicio"
        oWorkSheet.Cells(iFila, 9).Value = "Fecha Ingreso"
        oWorkSheet.Cells(iFila, 10).Value = "Fecha Alta"
        oWorkSheet.Cells(iFila, 11).Value = "Observaciones"
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
        iFila = 6: lnTotal = 0
        mrs_Tmp.MoveFirst
        Do While Not mrs_Tmp.EOF
           oWorkSheet.Cells(iFila, 2).Value = mrs_Tmp.Fields("idCuentaAtencion").Value
           oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("ApellidoPaterno").Value
           oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("apellidoMaterno").Value
           oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("PrimerNombre").Value
           oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp.Fields("nroHistoriaClinica").Value
           oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields("tipoServicio").Value
           oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp.Fields("Servicio").Value
           oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp.Fields("FechaIngreso").Value
           oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp.Fields("FechaEgreso").Value
           oWorkSheet.Cells(iFila, 11).Value = mrs_Tmp.Fields("observaciones").Value
           iFila = iFila + 1
           mrs_Tmp.MoveNext
        Loop
        mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
        oWorkSheet.Cells(iFila, 2).Value = "Nº Cuentas Observadas: " + Trim(Str(mrs_Tmp.RecordCount))
        oExcel.Visible = True
        oWorkSheet.PrintPreview
        'oWorkSheet.PrintOut
        Set oWorkSheet = Nothing
        Set oExcel = Nothing
        Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
        Resume
    End Select
    Exit Sub
End Sub



Sub CreandoRutaEnDisco(lcRutaTemporal As String)
    Dim lcSql As String
    On Error Resume Next
    lcSql = "c:"
    SIGHEntidades.ejecutarComando lcSql
    lcSql = "cd \"
    SIGHEntidades.ejecutarComando lcSql
    lcSql = "md " & lcRutaTemporal
    SIGHEntidades.ejecutarComando lcSql
    lcSql = "cd " & lcRutaTemporal
    SIGHEntidades.ejecutarComando lcSql
    lcSql = "del *.*"

End Sub
Function LeerArchivoTexto(nombreFichero As String) As String
    Dim numlib As Integer, isOpen As Boolean
    On Error GoTo Manejador_Error
    ' Obtengo el siguiente número libre de archivo
    numlib = FreeFile()
    Open nombreFichero For Input As #numlib
    ' Se ha abierto el fichero sin problemas
    isOpen = True
    ' Leo todo el contenido en una única operación
    LeerArchivoTexto = Input(LOF(numlib), numlib)
    ' Cierro el archivo
Manejador_Error:
    If isOpen Then Close #numlib
    If Err Then Err.Raise Err.Number, , Err.Description
End Function




