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
Option Explicit
Dim mo_ProgressRpt1 As XP_ProgressBar
Dim mo_ProgressRpt2 As XP_ProgressBar
Dim mo_ProgressRpt3 As XP_ProgressBar
Dim mo_ProgressRpt4 As XP_ProgressBar
Dim mo_ProgressRpt5 As XP_ProgressBar
Dim mo_ProgressRpt6 As XP_ProgressBar
Dim mo_ProgressRpt7 As XP_ProgressBar
Dim mo_lcNombrePc As String
Dim ml_idUsuario As Long
Dim mo_lnIdTablaLISTBARITEMS  As Long
Dim ms_MensajeError As String
Dim lcSql As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
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

Property Set ProgressRpt1(oValue As XP_ProgressBar)
    Set mo_ProgressRpt1 = oValue
End Property

Property Set progressRpt2(oValue As XP_ProgressBar)
    Set mo_ProgressRpt2 = oValue
End Property
Property Set progressRpt3(oValue As XP_ProgressBar)
    Set mo_ProgressRpt3 = oValue
End Property
Property Set progressRpt4(oValue As XP_ProgressBar)
    Set mo_ProgressRpt4 = oValue
End Property
Property Set progressRpt5(oValue As XP_ProgressBar)
    Set mo_ProgressRpt5 = oValue
End Property
Property Set progressRpt6(oValue As XP_ProgressBar)
    Set mo_ProgressRpt6 = oValue
End Property
Property Set progressRpt7(oValue As XP_ProgressBar)
    Set mo_ProgressRpt7 = oValue
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

Function AtencionesSeleccionarPorFechasDeIngresoYtipoServicio(lcMes As String, lcAnio As String, oConexion As Connection) As Recordset
   Dim oRsTmp As New Recordset, lcSql1 As String, lcSql2 As String
   lcSql1 = "SELECT     dbo.Atenciones.*, dbo.Pacientes.ApellidoPaterno AS ApellidoPaterno, dbo.Pacientes.ApellidoMaterno AS ApellidoMaterno, " & _
            "                      dbo.Pacientes.PrimerNombre AS PrimerNombre, dbo.Pacientes.SegundoNombre AS SegundoNombre, dbo.Pacientes.TercerNombre AS TercerNombre," & _
            "                      dbo.Pacientes.NroHistoriaClinica AS NroHistoriaClinica, dbo.Servicios.Nombre AS dServicio, dbo.Servicios.Codigo AS CodServ," & _
            "                      dbo.Pacientes.IdTipoSexo, dbo.Servicios.codigoServicioSEM, dbo.TiposCondicionAlta.codigoCondicionAltaSEM," & _
            "                      dbo.FuentesFinanciamiento.CodigoFuenteFinanciamientoSEM, dbo.Medicos.Colegiatura, dbo.Servicios.ubicacionSEM, dbo.TiposEdad.tipoEdadSEM," & _
            "                      dbo.Servicios.codigoServicioHIS, dbo.Empleados.IdEmpleado, dbo.TiposEmpleado.TipoEmpleadoHIS, dbo.Empleados.CodigoPlanilla," & _
            "                      dbo.Distritos.IdDistrito, dbo.Medicos.LoteHIS, dbo.TiposEdad.Codigo AS EdadCodigo," & _
            "                      dbo.Empleados.ApellidoPaterno + ' ' + dbo.Empleados.ApellidoMaterno + ' ' + dbo.Empleados.Nombres AS dEmpleado, dbo.Medicos.idColegioHIS," & _
            "                      dbo.Empleados.DNI, dbo.Pacientes.NroDocumento, dbo.Pacientes.IdEtnia, dbo.FuentesFinanciamiento.CodigoHIS AS HISfuenteFinanciamiento," & _
            "                      dbo.Pacientes.IdPaisProcedencia, dbo.TiposDocIdentidad.CodigoHIS AS HIStipoDocumento,dbo.Pacientes.FichaFamiliar,dbo.empleados.idTipoDocumento" & _
            " FROM         dbo.Pacientes LEFT OUTER JOIN" & _
            "                      dbo.TiposDocIdentidad ON dbo.Pacientes.IdDocIdentidad = dbo.TiposDocIdentidad.IdDocIdentidad LEFT OUTER JOIN" & _
            "                      dbo.Distritos ON dbo.Pacientes.IdDistritoDomicilio = dbo.Distritos.IdDistrito RIGHT OUTER JOIN" & _
            "                      dbo.Empleados INNER JOIN" & _
            "                      dbo.Medicos ON dbo.Empleados.IdEmpleado = dbo.Medicos.IdEmpleado LEFT OUTER JOIN" & _
            "                      dbo.TiposEmpleado ON dbo.Empleados.IdTipoEmpleado = dbo.TiposEmpleado.IdTipoEmpleado RIGHT OUTER JOIN" & _
            "                      dbo.Servicios RIGHT OUTER JOIN" & _
            "                      dbo.Atenciones ON dbo.Servicios.IdServicio = dbo.Atenciones.IdServicioIngreso ON" & _
            "                      dbo.Medicos.IdMedico = dbo.Atenciones.IdMedicoIngreso LEFT OUTER JOIN" & _
            "                      dbo.TiposEdad ON dbo.Atenciones.IdTipoEdad = dbo.TiposEdad.IdTipoEdad LEFT OUTER JOIN" & _
            "                      dbo.FuentesFinanciamiento ON dbo.Atenciones.idFuenteFinanciamiento = dbo.FuentesFinanciamiento.IdFuenteFinanciamiento LEFT OUTER JOIN" & _
            "                      dbo.TiposCondicionAlta ON dbo.Atenciones.IdCondicionAlta = dbo.TiposCondicionAlta.IdCondicionAlta ON" & _
            "                      dbo.Pacientes.idPaciente = dbo.Atenciones.idPaciente"
            
      lcSql2 = "    where  month(dbo.Atenciones.FechaIngreso)=" & lcMes & " and year(dbo.Atenciones.FechaIngreso)= " & lcAnio & "  and   (dbo.Atenciones.IdTipoServicio = 1)" & _
            "                        and dbo.Atenciones.idEstadoAtencion<>0" & _
            "             order by dbo.Empleados.IdEmpleado,dbo.Atenciones.FechaIngreso,dbo.Atenciones.HoraIngreso,dbo.Servicios.codigoServicioHIS"
    oRsTmp.Open lcSql1 & lcSql2, oConexion, adOpenKeyset, adLockOptimistic
    Set AtencionesSeleccionarPorFechasDeIngresoYtipoServicio = oRsTmp
End Function


Function DevuelveCodigoEnTablaHIS_TabDiag(lcCodigoDx As String, oConexionFox As Connection) As String
    Dim oRsTmp As New Recordset
    lcSql = "select * from tabDiag where Diag='" & Trim(lcCodigoDx) & "'"
    oRsTmp.Open lcSql, oConexionFox, adOpenKeyset, adLockOptimistic
    DevuelveCodigoEnTablaHIS_TabDiag = "0"
    If oRsTmp.RecordCount > 0 Then
       DevuelveCodigoEnTablaHIS_TabDiag = Trim(Str(oRsTmp.Fields!Codigo))
    End If
    oRsTmp.Close
    Set oRsTmp = Nothing
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
        Select Case lnBarraNumero
        Case 1
            mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnBarraTotal: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
        Case 2
            mo_ProgressRpt2.Min = 0: mo_ProgressRpt2.Max = lnBarraTotal: mo_ProgressRpt2.ShowText = True: mo_ProgressRpt2.Color = vbGreen
        Case 3
            mo_ProgressRpt3.Min = 0: mo_ProgressRpt3.Max = lnBarraTotal: mo_ProgressRpt3.ShowText = True: mo_ProgressRpt3.Color = vbGreen
        Case 4
            mo_ProgressRpt4.Min = 0: mo_ProgressRpt4.Max = lnBarraTotal: mo_ProgressRpt4.ShowText = True: mo_ProgressRpt4.Color = vbGreen
        Case 5
            mo_ProgressRpt5.Min = 0: mo_ProgressRpt5.Max = lnBarraTotal: mo_ProgressRpt5.ShowText = True: mo_ProgressRpt5.Color = vbGreen
        Case 6
            mo_ProgressRpt6.Min = 0: mo_ProgressRpt6.Max = lnBarraTotal: mo_ProgressRpt6.ShowText = True: mo_ProgressRpt6.Color = vbGreen
        End Select
        '
        oRsTmp1.MoveFirst
        lcLinea = ""
        Do While Not oRsTmp1.EOF
           Select Case lnBarraNumero
           Case 1
                mo_ProgressRpt1.Value = lnContador
           Case 2
                mo_ProgressRpt2.Value = lnContador
           Case 3
                mo_ProgressRpt3.Value = lnContador
           Case 4
                mo_ProgressRpt4.Value = lnContador
           Case 5
                mo_ProgressRpt5.Value = lnContador
           Case 6
                mo_ProgressRpt6.Value = lnContador
           End Select
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

Sub GeneraTXT(oRsTmp1 As Recordset, lcArchivoTXT As String, lnBarraTotal As Long, lnBarraNumero As Integer)
        On Error GoTo ErrGenTXT
        Dim lniFreeFile As Integer, lnUltimoCampo As Integer, lnCont As Integer, lcLinea As String
        Dim lcCampoValor As String, lnContador As Long, lcDNI As String
        Const lcSeparadorCampo As String = "|": Dim lcSeparadorSaltoLinea As String
        Dim lcDisa As String, lcFormato As String, lcNumero As String
        Dim oConexionExterna As New Connection
        Dim oRsAfiliadosSIS As New Recordset
        If Right(lcArchivoTXT, 13) = "Pacientes.txt" Then
            oConexionExterna.CommandTimeout = 300
            oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
            oConexionExterna.CursorLocation = adUseClient
        End If
        
        lcSeparadorSaltoLinea = Chr(10)
        lniFreeFile = FreeFile
        Open lcArchivoTXT For Output As #lniFreeFile
        lnUltimoCampo = oRsTmp1.Fields.Count
        '
        lnContador = 1
        Select Case lnBarraNumero
        Case 1
            mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnBarraTotal: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
        Case 2
            mo_ProgressRpt2.Min = 0: mo_ProgressRpt2.Max = lnBarraTotal: mo_ProgressRpt2.ShowText = True: mo_ProgressRpt2.Color = vbGreen
        Case 3
            mo_ProgressRpt3.Min = 0: mo_ProgressRpt3.Max = lnBarraTotal: mo_ProgressRpt3.ShowText = True: mo_ProgressRpt3.Color = vbGreen
        Case 4
            mo_ProgressRpt4.Min = 0: mo_ProgressRpt4.Max = lnBarraTotal: mo_ProgressRpt4.ShowText = True: mo_ProgressRpt4.Color = vbGreen
        Case 5
            mo_ProgressRpt5.Min = 0: mo_ProgressRpt5.Max = lnBarraTotal: mo_ProgressRpt5.ShowText = True: mo_ProgressRpt5.Color = vbGreen
        Case 6
            mo_ProgressRpt6.Min = 0: mo_ProgressRpt6.Max = lnBarraTotal: mo_ProgressRpt6.ShowText = True: mo_ProgressRpt6.Color = vbGreen
        End Select
        '
        oRsTmp1.MoveFirst
        Do While Not oRsTmp1.EOF
           Select Case lnBarraNumero
           Case 1
                mo_ProgressRpt1.Value = lnContador
           Case 2
                mo_ProgressRpt2.Value = lnContador
           Case 3
                mo_ProgressRpt3.Value = lnContador
           Case 4
                mo_ProgressRpt4.Value = lnContador
           Case 5
                mo_ProgressRpt5.Value = lnContador
           Case 6
                mo_ProgressRpt6.Value = lnContador
           End Select
           DoEvents: lnContador = lnContador + 1
           '
           lcLinea = ""
           lcDNI = ""
           For lnCont = 0 To lnUltimoCampo - 1
               If UCase(oRsTmp1.Fields(lnCont).Name) = "CABNROENVIOALSIS" Then
                  If IsNull(oRsTmp1.Fields(lnCont).Value) Then
                     oRsTmp1.Fields(lnCont).Value = 1
                  Else
                     oRsTmp1.Fields(lnCont).Value = Trim(Str(Val(oRsTmp1.Fields(lnCont).Value) + 1))
                  End If
                  oRsTmp1.Update
               End If
               lcCampoValor = ""
               
               If Not IsNull(oRsTmp1.Fields(lnCont).Value) Then
                    Select Case oRsTmp1.Fields(lnCont).Type
                    Case adDate
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    Case adChar, adVarChar
                         lcCampoValor = Trim(oRsTmp1.Fields(lnCont))
                         If UCase(oRsTmp1.Fields(lnCont).Name) = "DXCODIGO" Or UCase(oRsTmp1.Fields(lnCont).Name) = "CODIGO" And oRsTmp1.Fields(lnCont).DefinedSize = 5 Then
                             lcCampoValor = SIGHEntidades.DevuelveCodigoDxSinPUNTO(Left(oRsTmp1.Fields(lnCont).Value, 5))
                         End If
                         If lcCampoValor = lcSeparadorCampo Then
                            lcCampoValor = "_"
                         End If
                    Case adDouble, adInteger
                         lcCampoValor = Trim(Str((oRsTmp1.Fields(lnCont))))
                    Case Else
                         lcCampoValor = lcCampoValor & oRsTmp1.Fields(lnCont)
                    End Select
               End If
               If (lnUltimoCampo - 1) = lnCont Then
                  lcLinea = lcLinea & lcCampoValor & lcSeparadorCampo
               Else
                  lcLinea = lcLinea & lcCampoValor & lcSeparadorCampo
               End If
               If Right(lcArchivoTXT, 13) = "Pacientes.txt" And UCase(oRsTmp1.Fields(lnCont).Name) = "NRODOCUMENTO" Then
                  lcDNI = lcCampoValor
               End If
           Next
           '
           If Right(lcArchivoTXT, 13) = "Pacientes.txt" Then
                lcDisa = "": lcFormato = "": lcNumero = ""
                If Len(lcDNI) = 8 Then
                     lcSql = "select AfiliacionDisa,AfiliacionTipoFormato,AfiliacionNroFormato" & _
                             "  from SisFiliaciones where DocumentoNumero='" & lcDNI & "'"
                     If oRsAfiliadosSIS.State = 1 Then oRsAfiliadosSIS.Close
                     oRsAfiliadosSIS.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
                     If oRsAfiliadosSIS.RecordCount > 0 Then
                        lcDisa = Trim(oRsAfiliadosSIS.Fields!AfiliacionDisa)
                        lcFormato = Trim(oRsAfiliadosSIS.Fields!AfiliacionTipoFormato)
                        lcNumero = Trim(oRsAfiliadosSIS.Fields!AfiliacionNroFormato)
                     End If
                     oRsAfiliadosSIS.Close
                End If
                lcLinea = lcLinea & lcDisa & lcSeparadorCampo & lcFormato & lcSeparadorCampo & lcNumero & lcSeparadorCampo
           End If
           Print #lniFreeFile, lcLinea
           oRsTmp1.MoveNext
        Loop
        Close #lniFreeFile
        Set oConexionExterna = Nothing
        Set oRsAfiliadosSIS = Nothing
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
Dim iFila As Integer
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
        oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
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

Function ChequeaCuentasSinFUAgrabado(ByVal mrs_Tmp As Recordset, oConexionExterna As Connection, lcFechaInicio As String, lcFechaFinal As String) As Recordset
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim lnTotalReg As Long, lnIdCuentaInicial As Long, lbSeImprime As Boolean
        On Error GoTo ErrorProceso1
        lcSql = "SELECT      dbo.Atenciones.IdCuentaAtencion, dbo.Pacientes.ApellidoPaterno, dbo.Pacientes.ApellidoMaterno, dbo.Pacientes.PrimerNombre, " & _
                "                      dbo.Pacientes.NroHistoriaClinica, dbo.TiposServicio.Descripcion AS TipoServicio, dbo.Servicios.Nombre AS Servicio, dbo.Atenciones.FechaIngreso," & _
                "                      dbo.Atenciones.FechaEgreso , dbo.Atenciones.idEstadoAtencion, dbo.Atenciones.IdFormaPago,dbo.Atenciones.IdTipoServicio" & _
                " FROM         dbo.Pacientes RIGHT OUTER JOIN" & _
                "                      dbo.Servicios RIGHT OUTER JOIN" & _
                "                      dbo.Atenciones LEFT OUTER JOIN" & _
                "                      dbo.TiposServicio ON dbo.Atenciones.IdTipoServicio = dbo.TiposServicio.IdTipoServicio ON" & _
                "                      dbo.Servicios.IdServicio = dbo.Atenciones.IdServicioIngreso ON dbo.Pacientes.IdPaciente = dbo.Atenciones.IdPaciente" & _
                " Where (dbo.Atenciones.IdFormaPago = 2) And (dbo.Atenciones.idEstadoAtencion <> 0)" & _
                "       and dbo.Atenciones.fechaIngreso>=CONVERT(DATETIME,'" & lcFechaInicio & "',103) and dbo.Atenciones.fechaIngreso<=CONVERT(DATETIME,'" & lcFechaFinal & "',103)  " & _
                " ORDER BY dbo.Atenciones.IdCuentaAtencion"
        oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
        lnTotalReg = oRsTmp1.RecordCount
        'Chequea si hay Cuentas Observadas (No se grabaron en tablas SIS FUA)
        If lnTotalReg > 0 Then
           mo_ProgressRpt7.Min = 0: mo_ProgressRpt7.Max = lnTotalReg: mo_ProgressRpt7.ShowText = True: mo_ProgressRpt7.Color = vbGreen
           lnIdCuentaInicial = 1
           With mrs_Tmp
                  .Fields.Append "idCuentaAtencion", adInteger
                  .Fields.Append "ApellidoPaterno", adVarChar, 50, adFldIsNullable
                  .Fields.Append "apellidoMaterno", adVarChar, 50, adFldIsNullable
                  .Fields.Append "PrimerNombre", adVarChar, 50, adFldIsNullable
                  .Fields.Append "nroHistoriaClinica", adInteger
                  .Fields.Append "tipoServicio", adVarChar, 50, adFldIsNullable
                  .Fields.Append "Servicio", adVarChar, 100, adFldIsNullable
                  .Fields.Append "FechaIngreso", adDate
                  .Fields.Append "FechaEgreso", adDate, 8, adFldIsNullable
                  .Fields.Append "Observaciones", adVarChar, 100, adFldIsNullable
                  .LockType = adLockOptimistic
                  .Open
           End With
           oRsTmp1.MoveFirst
           Do While Not oRsTmp1.EOF
              mo_ProgressRpt7.Value = lnIdCuentaInicial: lnIdCuentaInicial = lnIdCuentaInicial + 1
              lcSql = "select * from sisFuaAtencion where idCuentaAtencion=" & oRsTmp1.Fields!idCuentaAtencion
              oRsTmp2.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
              If oRsTmp2.RecordCount = 0 Then
                 lbSeImprime = False
                 If oRsTmp1.Fields!IdTipoServicio = 1 Then
                    lbSeImprime = True
                    lcSql = "Cuenta de Consultorios externos, no se ha impreso FICHA FUA"
                 Else
                    If Not IsNull(oRsTmp1.Fields!FechaEgreso) Then
                       lbSeImprime = True
                       lcSql = "Cuenta de Hospitalización/emergencia, tiene ALTA MEDICA pero no se ha impreso FICHA FUA"
                    End If
                 End If
                 If lbSeImprime = True Then
                    mrs_Tmp.AddNew
                    mrs_Tmp.Fields!idCuentaAtencion = oRsTmp1.Fields!idCuentaAtencion
                    mrs_Tmp.Fields!ApellidoPaterno = oRsTmp1.Fields!ApellidoPaterno
                    mrs_Tmp.Fields!ApellidoMaterno = oRsTmp1.Fields!ApellidoMaterno
                    mrs_Tmp.Fields!PrimerNombre = oRsTmp1.Fields!PrimerNombre
                    mrs_Tmp.Fields!NroHistoriaClinica = oRsTmp1.Fields!NroHistoriaClinica
                    mrs_Tmp.Fields!TipoServicio = oRsTmp1.Fields!TipoServicio
                    mrs_Tmp.Fields!Servicio = oRsTmp1.Fields!Servicio
                    mrs_Tmp.Fields!FechaIngreso = oRsTmp1.Fields!FechaIngreso
                    mrs_Tmp.Fields!FechaEgreso = oRsTmp1.Fields!FechaEgreso
                    mrs_Tmp.Fields!Observaciones = lcSql
                    mrs_Tmp.Update
                 End If
              End If
              oRsTmp2.Close
              oRsTmp1.MoveNext
           Loop
        End If
        Set ChequeaCuentasSinFUAgrabado = oRsTmp1
        Exit Function
ErrorProceso1:
        MsgBox Err.Description
        Resume
End Function


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


Function ExportaCitasWeb(ldFechaInicial As Date, ldFechaFinal As Date, oRsCuposWebSeleccionado As Recordset, _
         lnIdUsuario As Long, lbEnviarServicio As Boolean, lbEnviarPacientes As Boolean, _
         lbEnviarFteFinanciamiento As Boolean, lbEnviarMedicos As Boolean, lbEnviarTurnos As Boolean) As Boolean
       On Error GoTo ErrCitasWebExp
       ExportaCitasWeb = False
       Dim oConexion As New Connection, oConexionExterna As New Connection
       Dim oRsCitasWebCupos As New Recordset, oRsCitaBloqueada As New Recordset
       Dim oRsCitaBloqueada1 As New Recordset, oRsTurnos As New Recordset
       Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
       Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
       Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
       Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
       Dim lcRutaExportar As String, lnTotalRegistros As Long, lnIdCitaBloqueada As Long, lnIdTurno As Long
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open SIGHEntidades.CadenaConexion
       oConexion.BeginTrans
       oConexionExterna.CommandTimeout = 300
       oConexionExterna.CursorLocation = adUseClient
       oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
       oConexionExterna.BeginTrans
       'Elimina CUPOS RESERVADOS del mes
       If mo_ReglasDeProgMedica.CitasWebCuposEliminarPorFechas(ldFechaInicial, ldFechaFinal, oConexionExterna) = False Then
          lcParaMsgbox = mo_ReglasDeProgMedica.MensajeError: GoTo ErrCitasWebExp
       End If
       'Elimina CITAS BLOQUEADAS del mes
'       If mo_ReglasDeProgMedica.CitasBloqueadasEliminarPorFechas(ldFechaInicial, ldFechaFinal, oConexion) = False Then
'          lcParaMsgbox = mo_ReglasDeProgMedica.MensajeError: GoTo ErrCitasWebExp
'       End If
       oRsCuposWebSeleccionado.Filter = ""
       lnTotalRegistros = oRsCuposWebSeleccionado.RecordCount
       If oRsCuposWebSeleccionado.RecordCount > 0 Then
            Set oRsTurnos = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdTipoServicio(1)
            mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnTotalRegistros: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
            lcRutaExportar = lcBuscaParametro.SeleccionaFilaParametro(313)
            oRsCitaBloqueada.Open "select * from CitasBloqueadas", oConexion, adOpenKeyset, adLockOptimistic
            oRsCitasWebCupos.Open "select * from CitasWebCupos", oConexionExterna, adOpenKeyset, adLockOptimistic
            oRsCuposWebSeleccionado.MoveFirst
            Do While Not oRsCuposWebSeleccionado.EOF
If oRsCuposWebSeleccionado.Fields!IdServicio = 93 And Day(oRsCuposWebSeleccionado.Fields!fecha) = 31 Then
ms_MensajeError = ""
End If
                mo_ProgressRpt1.Value = mo_ProgressRpt1.Value + 1
                'Bloquea Cita
                lnIdCitaBloqueada = IIf(IsNull(oRsCuposWebSeleccionado.Fields!idCitaBloqueada), 0, oRsCuposWebSeleccionado.Fields!idCitaBloqueada)
                If oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb <> sghCitaWebEstados.CupoLlenadoEnCitaGalenHos Then
                    If lnIdCitaBloqueada = 0 Then
                        oRsCitaBloqueada.AddNew
                        oRsCitaBloqueada.Fields!idUsuario = lnIdUsuario
                        oRsCitaBloqueada.Fields!fecha = oRsCuposWebSeleccionado.Fields!fecha
                        oRsCitaBloqueada.Fields!HoraInicio = oRsCuposWebSeleccionado.Fields!HoraInicio
                        oRsCitaBloqueada.Fields!HoraFin = oRsCuposWebSeleccionado.Fields!HoraFinal
                        oRsCitaBloqueada.Fields!FechaBloqueo = Date
                        oRsCitaBloqueada.Fields!HoraBloqueo = "Web"
                        oRsCitaBloqueada.Fields!idMedico = oRsCuposWebSeleccionado.Fields!idMedico
                        oRsCitaBloqueada.Update
                        lnIdCitaBloqueada = oRsCitaBloqueada.Fields!idCitaBloqueada
                    End If
                Else
                    If lnIdCitaBloqueada > 0 Then
                        lcSql = "delete from CitasBloqueadas where idCitaBloqueada=" & lnIdCitaBloqueada
                        oRsCitaBloqueada1.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
                    End If
                End If
                '
                If IsNull(oRsCuposWebSeleccionado.Fields!IdTurno) Then
                    lnIdTurno = 4
                    If oRsTurnos.RecordCount > 0 Then
                       oRsTurnos.MoveFirst
                       Do While Not oRsTurnos.EOF
                          If oRsCuposWebSeleccionado.Fields!HoraInicio >= oRsTurnos.Fields!HoraInicio And oRsCuposWebSeleccionado.Fields!HoraInicio < oRsTurnos.Fields!HoraFin Then
                             lnIdTurno = oRsTurnos.Fields!IdTurno
                             Exit Do
                          End If
                          oRsTurnos.MoveNext
                       Loop
                    End If
                Else
                    lnIdTurno = oRsCuposWebSeleccionado.Fields!IdTurno
                End If
                'Carga cupo
                oRsCitasWebCupos.AddNew
                oRsCitasWebCupos.Fields!HoraInicio = oRsCuposWebSeleccionado.Fields!HoraInicio
                oRsCitasWebCupos.Fields!HoraFinal = oRsCuposWebSeleccionado.Fields!HoraFinal
                oRsCitasWebCupos.Fields!IdServicio = oRsCuposWebSeleccionado.Fields!IdServicio
                oRsCitasWebCupos.Fields!fecha = oRsCuposWebSeleccionado.Fields!fecha
                oRsCitasWebCupos.Fields!idMedico = oRsCuposWebSeleccionado.Fields!idMedico
                oRsCitasWebCupos.Fields!idEstadoCitaWeb = oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb
                oRsCitasWebCupos.Fields!IdTurno = lnIdTurno
                If oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb <> sghCitaWebEstados.CupoLlenadoEnCitaGalenHos Then
                    oRsCitasWebCupos.Fields!idCitaBloqueada = lnIdCitaBloqueada
                    oRsCitasWebCupos.Fields!Ubigeo = IIf(IsNull(oRsCuposWebSeleccionado.Fields!Ubigeo), 0, oRsCuposWebSeleccionado.Fields!Ubigeo)
                    oRsCitasWebCupos.Fields!FechaNacimiento = Null
                Else
                    oRsCitasWebCupos.Fields!DNI = IIf(IsNull(oRsCuposWebSeleccionado.Fields!DNI), "", oRsCuposWebSeleccionado.Fields!DNI)
                    oRsCitasWebCupos.Fields!ApellidoPaterno = IIf(IsNull(oRsCuposWebSeleccionado.Fields!ApellidoPaterno), "", oRsCuposWebSeleccionado.Fields!ApellidoPaterno)
                    oRsCitasWebCupos.Fields!ApellidoMaterno = IIf(IsNull(oRsCuposWebSeleccionado.Fields!ApellidoMaterno), "", oRsCuposWebSeleccionado.Fields!ApellidoMaterno)
                    oRsCitasWebCupos.Fields!PrimerNombre = IIf(IsNull(oRsCuposWebSeleccionado.Fields!PrimerNombre), "", oRsCuposWebSeleccionado.Fields!PrimerNombre)
                    oRsCitasWebCupos.Fields!SegundoNombre = IIf(IsNull(oRsCuposWebSeleccionado.Fields!SegundoNombre), "", oRsCuposWebSeleccionado.Fields!SegundoNombre)
                    oRsCitasWebCupos.Fields!idTipoSexo = IIf(IsNull(oRsCuposWebSeleccionado.Fields!idTipoSexo), 0, oRsCuposWebSeleccionado.Fields!idTipoSexo)
                    If oRsCuposWebSeleccionado.Fields!FechaNacimiento <> 0 Then
                       oRsCitasWebCupos.Fields!FechaNacimiento = oRsCuposWebSeleccionado.Fields!FechaNacimiento
                    Else
                       oRsCitasWebCupos.Fields!FechaNacimiento = Null
                    End If
                    oRsCitasWebCupos.Fields!Ubigeo = IIf(IsNull(oRsCuposWebSeleccionado.Fields!Ubigeo), 0, oRsCuposWebSeleccionado.Fields!Ubigeo)
                End If
                oRsCitasWebCupos.Update
                oRsCuposWebSeleccionado.MoveNext
            Loop
            'Exporta Cupos Web y Citas
            GeneraTXT oRsCitasWebCupos, lcRutaExportar & "CuposWeb.txt", oRsCitasWebCupos.RecordCount, 2
            'Exporta Servicios
            If lbEnviarServicio = True Then
                oRsCitasWebCupos.Close
                Set oRsCitasWebCupos = mo_ReglasArchivoClinico.ServiciosSeleccionarSoloCE
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Servicios.txt", oRsCitasWebCupos.RecordCount, 3
            End If
            'Exporta Medicos
            If lbEnviarMedicos = True Then
                oRsCitasWebCupos.Close
                Set oRsCitasWebCupos = mo_ReglasDeProgMedica.MedicosSeleccionarTodosParaWeb
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Medicos.txt", oRsCitasWebCupos.RecordCount, 4
            End If
            'Exporta Pacientes
            If lbEnviarPacientes = True Then
                oRsCitasWebCupos.Close
                Set oRsCitasWebCupos = mo_ReglasAdmision.PacientesSeleccionarTodos
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Pacientes.txt", oRsCitasWebCupos.RecordCount, 5
            End If
            'Exporta Fuente Financiamiento
            If lbEnviarFteFinanciamiento = True Then
                oRsCitasWebCupos.Close
                Set oRsCitasWebCupos = mo_ReglasComunes.FuentesFinanciamientoSeleccionarTodos
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "ftefinan.txt", oRsCitasWebCupos.RecordCount, 6
            End If
            If lbEnviarTurnos = True Then
                oRsCitasWebCupos.Close
                Set oRsCitasWebCupos = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdTipoServicio(1)
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Turnos.txt", oRsCitasWebCupos.RecordCount, 1
            End If
            MsgBox "Se exportó correctamente los archivos textos, en la Ruta: " & lcRutaExportar
       End If
       
       ExportaCitasWeb = True
ErrCitasWebExp:
       If ExportaCitasWeb = True Then
          oConexion.CommitTrans
          oConexionExterna.CommitTrans
       Else
          oConexion.RollbackTrans
          oConexionExterna.RollbackTrans
       End If
       oConexion.Close
       Set oConexion = Nothing
       oConexionExterna.Close
       Set oConexionExterna = Nothing
       Set oRsCitasWebCupos = Nothing
       Set oRsCitaBloqueada = Nothing
       Set mo_ReglasDeProgMedica = Nothing
       Set mo_ReglasArchivoClinico = Nothing
       Set mo_ReglasAdmision = Nothing
       Set oRsCitaBloqueada1 = Nothing
       Set oRsTurnos = Nothing
End Function


Function ExportaCitasWebCargaTabla(ldFechaInicial As Date, ldFechaFinal As Date, ByRef oRsCitasWebCupos As Recordset) As Boolean
       ExportaCitasWebCargaTabla = False
       On Error GoTo excwct
       Dim oRsCuposWebSeleccionado As New Recordset
       Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
       Set oRsCuposWebSeleccionado = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorFechas(ldFechaInicial, ldFechaFinal)
       If oRsCuposWebSeleccionado.RecordCount > 0 Then
            oRsCuposWebSeleccionado.MoveFirst
            Do While Not oRsCuposWebSeleccionado.EOF
                oRsCitasWebCupos.AddNew
                oRsCitasWebCupos.Fields!HoraInicio = oRsCuposWebSeleccionado.Fields!HoraInicio
                oRsCitasWebCupos.Fields!HoraFinal = oRsCuposWebSeleccionado.Fields!HoraFinal
                oRsCitasWebCupos.Fields!IdServicio = oRsCuposWebSeleccionado.Fields!IdServicio
                oRsCitasWebCupos.Fields!fecha = oRsCuposWebSeleccionado.Fields!fecha
                oRsCitasWebCupos.Fields!idMedico = oRsCuposWebSeleccionado.Fields!idMedico
                oRsCitasWebCupos.Fields!idEstadoCitaWeb = oRsCuposWebSeleccionado.Fields!idEstadoCitaWeb
                oRsCitasWebCupos.Fields!idCitaBloqueada = oRsCuposWebSeleccionado.Fields!idCitaBloqueada
                oRsCitasWebCupos.Fields!DNI = IIf(IsNull(oRsCuposWebSeleccionado.Fields!DNI), "", oRsCuposWebSeleccionado.Fields!DNI)
                oRsCitasWebCupos.Fields!ApellidoPaterno = IIf(IsNull(oRsCuposWebSeleccionado.Fields!ApellidoPaterno), "", oRsCuposWebSeleccionado.Fields!ApellidoPaterno)
                oRsCitasWebCupos.Fields!ApellidoMaterno = IIf(IsNull(oRsCuposWebSeleccionado.Fields!ApellidoMaterno), "", oRsCuposWebSeleccionado.Fields!ApellidoMaterno)
                oRsCitasWebCupos.Fields!PrimerNombre = IIf(IsNull(oRsCuposWebSeleccionado.Fields!PrimerNombre), "", oRsCuposWebSeleccionado.Fields!PrimerNombre)
                oRsCitasWebCupos.Fields!SegundoNombre = IIf(IsNull(oRsCuposWebSeleccionado.Fields!SegundoNombre), "", oRsCuposWebSeleccionado.Fields!SegundoNombre)
                oRsCitasWebCupos.Fields!idTipoSexo = IIf(IsNull(oRsCuposWebSeleccionado.Fields!idTipoSexo), 0, oRsCuposWebSeleccionado.Fields!idTipoSexo)
                oRsCitasWebCupos.Fields!FechaNacimiento = IIf(IsNull(oRsCuposWebSeleccionado.Fields!FechaNacimiento), 0, oRsCuposWebSeleccionado.Fields!FechaNacimiento)
                oRsCitasWebCupos.Update
                oRsCuposWebSeleccionado.MoveNext
            Loop
       End If
       oRsCuposWebSeleccionado.Close
       Set oRsCuposWebSeleccionado = Nothing
       Set mo_ReglasDeProgMedica = Nothing
       ExportaCitasWebCargaTabla = True
       Exit Function
excwct:
       lcParaMsgbox = Err.Description
End Function

Function ImportaCitasWeb(ByRef oRsCitasWebCupos As Recordset, lcArchivoTexto As String) As Boolean
     On Error GoTo ErrImpCWEB
     ImportaCitasWeb = False
     Dim oExcel As Excel.Application
     Dim oWorkBookPlantilla As Workbook
     Dim oWorkBook As Workbook
     Dim oWorkSheet As Worksheet
     Dim mo_ReporteUtil As New SIGHEntidades.ReporteUtil
     Dim iFila As Integer, lnIdWeb As Long
     Dim lineas() As String, lnRegistros As Long, Registro As String, Fichero As Integer, Campos() As String, Columna As Single
     Dim ldFecha As Date, lnIdServicio As Long, lnIdMedico As Long, lcHoraInicio As String
     Dim lcHoraFinal As String, lnIdEstadoCitaWeb As Long, lnIdCitaBloqueada As Long, lcDNI As String
     Dim lcApellidoPaterno As String, lcApellidoMaterno As String, lcPrimerNombre As String
     Dim ldFechaConfirmacion As Date, lcHoraConfirmacion As String, lnNroRegistro As Long
     Dim lcSegundoNombre As String, lnIdTipoSexo As Long, ldFechaNacimiento As Date, lnUbigeo As Long
     Dim lnDiasMaximoConfirmacion As Integer, lnDiasQfaltaPcita As Integer, lnIdFuenteFinanciamiento As Long
     Dim ldFechaHoy As Date, lcHoraHoy As String, lnHistoriaCreada As Long
     Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
     Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
     Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
     Dim oConexion As New Connection, oRsTmp1 As New Recordset
     '
     Set oExcel = GalenhosExcelApplication()
     Set oWorkBook = oExcel.Workbooks.Add
     Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HojaLibre.xls")
     oWorkBookPlantilla.Worksheets("hoja_Libre").Copy Before:=oWorkBook.Sheets(1)
     oWorkBookPlantilla.Close
     Set oWorkSheet = oWorkBook.Sheets(1)
     iFila = 2
     oWorkSheet.Cells(iFila, 2).Value = "Lista de HISTORIAS creadas en forma automática (Citas Web)"
     iFila = 4
     oWorkSheet.PageSetup.LeftHeaderPicture.FileName = App.Path + "\imagenes\logotipo.jpg"
     oWorkSheet.Cells(iFila, 2).Value = "N° Historia Clinica"
     oWorkSheet.Cells(iFila, 3).Value = "DNI"
     oWorkSheet.Cells(iFila, 4).Value = "Apellido Paterno"
     oWorkSheet.Cells(iFila, 5).Value = "Apellido Materno"
     oWorkSheet.Cells(iFila, 6).Value = "Primer Nombre"
     oWorkSheet.Cells(iFila, 7).Value = "Servicio"
     oWorkSheet.Cells(iFila, 8).Value = "Fecha Cita"
     oWorkSheet.Cells(iFila, 9).Value = "Hora Cita"
     mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
     iFila = 6
     '
     ldFechaHoy = CDate(lcBuscaParametro.RetornaFechaServidorSQL): lcHoraHoy = lcBuscaParametro.RetornaHoraServidorSQL
     lineas() = Split(LeerArchivoTexto(lcArchivoTexto), vbCrLf)
     lnRegistros = UBound(lineas)
     mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnRegistros: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
     oConexion.CommandTimeout = 300
     oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
     oConexion.CursorLocation = adUseClient
     Fichero = FreeFile
     Open lcArchivoTexto For Input As #Fichero
     lnNroRegistro = 1
     While Not EOF(Fichero)
        mo_ProgressRpt1.Value = lnNroRegistro: lnNroRegistro = lnNroRegistro + 1
        Line Input #Fichero, Registro
        Campos = Split(Registro, "|")
        lnIdWeb = 0
        ldFecha = 0: lnIdServicio = 0: lnIdMedico = 0: lcHoraInicio = "": lcHoraFinal = "": lnIdEstadoCitaWeb = 0
        lnIdCitaBloqueada = 0: lcDNI = "": lcApellidoPaterno = "": lcApellidoMaterno = "": lcPrimerNombre = ""
        lcSegundoNombre = "": lnIdTipoSexo = 0: ldFechaNacimiento = 0: lnUbigeo = 0
        For Columna = 0 To UBound(Campos)
            Select Case Columna
            Case 0
                 lnIdWeb = Val(Campos(Columna))
            Case 1
                 ldFecha = CDate(Campos(Columna))
            Case 2
                 lnIdServicio = Val(Campos(Columna))
            Case 3
                 lnIdMedico = Val(Campos(Columna))
            Case 4
                 lcHoraInicio = Campos(Columna)
            Case 5
                 lcHoraFinal = Campos(Columna)
            Case 6
                 lnIdEstadoCitaWeb = Val(Campos(Columna))
            Case 7
                 lnIdCitaBloqueada = Val(Campos(Columna))
            Case 8
                 lcDNI = Campos(Columna)
            Case 9
                 lcApellidoPaterno = Campos(Columna)
            Case 10
                 lcApellidoMaterno = Campos(Columna)
            Case 11
                 lcPrimerNombre = Campos(Columna)
            Case 12
                 lcSegundoNombre = Campos(Columna)
            Case 13
                 lnIdTipoSexo = Val(Campos(Columna))
            Case 14
                 If Campos(Columna) <> "" Then
                    ldFechaNacimiento = CDate(Campos(Columna))
                 End If
            Case 15
                 lnUbigeo = Val(Campos(Columna))
            Case 16
                 If Campos(Columna) <> "" Then
                    ldFechaConfirmacion = CDate(Campos(Columna))
                 End If
            Case 17
                 lcHoraConfirmacion = Campos(Columna)
            Case 18
                 lnIdFuenteFinanciamiento = Val(Campos(Columna))
            End Select
        Next
        If lnIdCitaBloqueada > 0 And ldFechaHoy <= ldFecha Then
           lnDiasMaximoConfirmacion = Val(lcBuscaParametro.SeleccionaFilaParametro(321))
           If lnIdEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb Then
                lnDiasQfaltaPcita = DateDiff("d", CDate(Format(ldFechaConfirmacion, "dd/mm/yyyy") & " " & lcHoraConfirmacion), CDate(Format(ldFecha, "dd/mm/yyyy") & " " & lcHoraInicio))
                Set oRsTmp1 = mo_ReglasDeProgMedica.CitasWebCuposSeleccionarPorIdCitaBloqueada(lnIdCitaBloqueada)
                If oRsTmp1.RecordCount > 0 Then
                    If lnDiasQfaltaPcita < lnDiasMaximoConfirmacion Then
                       'elimina cupo
                       oRsTmp1.Delete
                       oRsTmp1.Update
                       If mo_ReglasDeProgMedica.CitasBloqueadasEliminarPorId(lnIdCitaBloqueada) = False Then
                          MsgBox mo_ReglasDeProgMedica.MensajeError: GoTo ErrImpCWEB
                       End If
                    Else
                       'asigna al cupo reservado la cita web, incluyendo el paciente web (si es necesario crea HISTORIA)
                       oRsTmp1.Fields!idWeb = lnIdWeb
                       oRsTmp1.Fields!idEstadoCitaWeb = sghCitaWebEstados.CupoConfirmadoEnCitaWeb
                       oRsTmp1.Fields!idCitaBloqueada = Null
                       oRsTmp1.Fields!DNI = lcDNI
                       oRsTmp1.Fields!ApellidoPaterno = lcApellidoPaterno
                       oRsTmp1.Fields!ApellidoMaterno = lcApellidoMaterno
                       oRsTmp1.Fields!PrimerNombre = lcPrimerNombre
                       oRsTmp1.Fields!SegundoNombre = lcSegundoNombre
                       oRsTmp1.Fields!idTipoSexo = lnIdTipoSexo
                       oRsTmp1.Fields!FechaNacimiento = ldFechaNacimiento
                       oRsTmp1.Fields!Ubigeo = lnUbigeo
                       oRsTmp1.Fields!FechaConfirmacion = ldFechaConfirmacion
                       oRsTmp1.Fields!HoraConfirmacion = lcHoraConfirmacion
                       oRsTmp1.Update
                       If mo_ReglasDeProgMedica.CitasBloqueadasEliminarPorId(lnIdCitaBloqueada) = False Then
                          MsgBox mo_ReglasDeProgMedica.MensajeError: GoTo ErrImpCWEB
                       End If
                       If CrearCitaAutomatica(lcDNI, ldFecha, lcHoraInicio, lcHoraFinal, lnIdServicio, lnIdMedico, lcApellidoMaterno, _
                                      lcApellidoPaterno, ldFechaNacimiento, lnUbigeo, lnIdTipoSexo, lcPrimerNombre, _
                                      lcSegundoNombre, lnIdFuenteFinanciamiento, ldFechaHoy, lcHoraHoy, lnHistoriaCreada) = False Then
                          MsgBox ms_MensajeError: GoTo ErrImpCWEB
                       End If
                       If lnHistoriaCreada > 0 Then
                          oWorkSheet.Cells(iFila, 2).Value = "'" & lnHistoriaCreada
                          oWorkSheet.Cells(iFila, 3).Value = lcDNI
                          oWorkSheet.Cells(iFila, 4).Value = lcApellidoPaterno
                          oWorkSheet.Cells(iFila, 5).Value = lcApellidoMaterno
                          oWorkSheet.Cells(iFila, 6).Value = lcPrimerNombre
                          oWorkSheet.Cells(iFila, 7).Value = mo_ReglasServiciosHosp.ServiciosDevuelveNombre(lnIdServicio)
                          oWorkSheet.Cells(iFila, 8).Value = ldFecha
                          oWorkSheet.Cells(iFila, 9).Value = lcHoraInicio
                          iFila = iFila + 1
                       End If
                    End If
                End If
                oRsTmp1.Close
           End If
        End If
        
    Wend
    Close #Fichero
    '
    mo_ReporteUtil.ExcelCuadricularRango oExcel, oWorkSheet, iFila, 2, iFila, 11
    oExcel.Visible = True
    oWorkSheet.PrintPreview
    '
    ImportaCitasWeb = True
ErrImpCWEB:
    Set mo_ReglasDeProgMedica = Nothing
    Set mo_ReglasAdmision = Nothing
    Set mo_ReglasServiciosHosp = Nothing
    Set oConexion = Nothing
    Set oRsTmp1 = Nothing
    'Resume
End Function

Function CrearCitaAutomatica(lcDNI As String, ldFechaCita As Date, lcHoraInicio As String, lcHoraFinal As String, lnIdServicio As Long, _
                   lnIdMedico As Long, lcApellidoMaterno, lcApellidoPaterno As String, ldFechaNacimiento As Date, _
                   lnUbigeo As Long, lnIdTipoSexo As Long, lcPrimerNombre As String, lcSegundoNombre As String, _
                   lnIdFuenteFinanciamiento As Long, ldFechaHoy As Date, lcHoraHoy As String, _
                   ByRef lnHistoriaCreada As Long) As Boolean
     On Error GoTo ErrCrearCitaWeb
     ms_MensajeError = ""
     CrearCitaAutomatica = False
     Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision
     Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
     Dim mo_ReglasServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
     Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
     Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
     Dim mo_CuentasAtencion As New DOCuentaAtencion, mo_Atenciones As New DOAtencion, mo_Pacientes As New doPaciente
     Dim mo_DOFacturacionPaquetes As New DOFacturacionPaquetes, oDoSunasaPacientesHistoricos As New DoSunasaPacientesHistoricos
     Dim mo_Cita As New DOCita, mo_Historia As New DOHistoriaClinica, mo_DoAtencionDatosAdicionales As New DoAtencionDatosAdicionales
     Dim mo_FacturacionServicios As New Collection
     Dim oPacientes As New Pacientes, oHistoria As New HistoriasClinicas
     Dim oRsTmp21 As New Recordset
     Dim oConexion As New Connection
     Dim lnTipoNumeracionAnterior As Long, lnIdHistoriaClinicaAnterior As Long, lbYaSeTransfirioHCdeUnServicioAotro As Boolean
     Dim mo_Diagnosticos As New Collection, mo_Procedimientos As New Collection, mo_Examenes As New Collection
     Dim lnIdTipoFinanciamiento As Long, lnIdProductoCpt As Long, lnIdEspecialidadMedico As Long, lnIdProgramacion As Long
     Dim oEdad As Edad, lbNecesitaCrearHistoria As Boolean
     Dim lnIdCondicionAlServicio As Long, lnIdCondicionAlEstablecimiento As Long
     oConexion.CommandTimeout = 300
     oConexion.CursorLocation = adUseClient
     oConexion.Open SIGHEntidades.CadenaConexion
     '
     
     '
     Set oRsTmp21 = mo_ReglasFacturacion.TiposFinanciamientosTarifaSeleccionarPorPlan(lnIdFuenteFinanciamiento)
     lnIdTipoFinanciamiento = 0
     If oRsTmp21.RecordCount > 0 Then
        lnIdTipoFinanciamiento = oRsTmp21.Fields!idTipoFinanciamiento
     End If
     oRsTmp21.Close
     '
     lnIdProductoCpt = 0
     Set oRsTmp21 = mo_ReglasServiciosHosp.EspecialidadCESeleccionarPorIdServicio(lnIdServicio)
     If oRsTmp21.RecordCount > 0 Then
        lnIdProductoCpt = oRsTmp21.Fields!IdProductoConsulta
     End If
     oRsTmp21.Close
     '
     lnIdEspecialidadMedico = 0: lnIdProgramacion = 0
     Set oRsTmp21 = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorMedicoFechaServicio(lnIdMedico, Format(ldFechaCita, "dd/mm/yyyy"), lnIdServicio)
     If oRsTmp21.RecordCount > 0 Then
        lnIdEspecialidadMedico = oRsTmp21.Fields!IdEspecialidad: lnIdProgramacion = oRsTmp21.Fields!IdProgramacion
     End If
     oRsTmp21.Close
     '
     Set mo_Pacientes = mo_AdminAdmision.PacientesSeleccionarPorDNI(lcDNI)
     If mo_Pacientes.idPaciente = 0 Then
        lbNecesitaCrearHistoria = True
        mo_Pacientes.ApellidoMaterno = UCase(lcApellidoMaterno)
        mo_Pacientes.ApellidoPaterno = UCase(lcApellidoPaterno)
        mo_Pacientes.FechaNacimiento = ldFechaNacimiento
        mo_Pacientes.IdDistritoDomicilio = lnUbigeo
        mo_Pacientes.IdDocIdentidad = 1
        mo_Pacientes.idPaciente = 0
        mo_Pacientes.idTipoSexo = lnIdTipoSexo
        mo_Pacientes.IdUsuarioAuditoria = ml_idUsuario
        mo_Pacientes.NroDocumento = lcDNI
        mo_Pacientes.NroHistoriaClinica = 0
        mo_Pacientes.PrimerNombre = UCase(lcPrimerNombre)
        mo_Pacientes.SegundoNombre = lcSegundoNombre
        mo_Pacientes.IdTipoNumeracion = 1
        mo_Pacientes.Autogenerado = mo_AdminAdmision.PacienteCrearNroAutogenerado(mo_Pacientes)
     End If
     oEdad = CalcularEdad(mo_Pacientes.FechaNacimiento, CDate(Format(ldFechaCita, "dd/mm/yyyy") & " " & lcHoraInicio))
     mo_ReglasComunes.TiposCondicionPacienteCondicionAlEstablecimientoYservicio lnIdCondicionAlEstablecimiento, _
                              lnIdCondicionAlServicio, mo_Pacientes.idPaciente, _
                              Format(ldFechaCita, SIGHEntidades.DevuelveFechaSoloFormato_DMY), 0, lnIdServicio
     If mo_Pacientes.idPaciente = 0 Then
        mo_Historia.FechaCreacion = ldFechaHoy
        mo_Historia.IdEstadoHistoria = 1
        mo_Historia.idPaciente = 0
        mo_Historia.IdTipoNumeracion = 1
        mo_Historia.IdUsuarioAuditoria = ml_idUsuario
        mo_Historia.NroHistoriaClinica = 0
        mo_Historia.IdTipoHistoria = 1
     Else
        mo_Historia.NroHistoriaClinica = mo_Pacientes.NroHistoriaClinica
        Set oHistoria.Conexion = oConexion
        If oHistoria.SeleccionarPorId(mo_Historia) = False Then
           ms_MensajeError = oHistoria.MensajeError: GoTo ErrCrearCitaWeb
        End If
     End If
     '
     mo_CuentasAtencion.FechaApertura = ldFechaHoy
     mo_CuentasAtencion.FechaCreacion = ldFechaHoy
     mo_CuentasAtencion.HoraApertura = lcHoraHoy
     mo_CuentasAtencion.idEstado = sghEstadoCuenta.sghAbierto
     mo_CuentasAtencion.idPaciente = mo_Pacientes.idPaciente
     mo_CuentasAtencion.IdUsuarioAuditoria = ml_idUsuario
     '
     mo_Atenciones.Edad = oEdad.Edad
     mo_Atenciones.EsPacienteExterno = False
     mo_Atenciones.FechaIngreso = ldFechaCita
     mo_Atenciones.HoraIngreso = lcHoraInicio
     mo_Atenciones.idAtencion = 0
     mo_Atenciones.idCuentaAtencion = 0
     mo_Atenciones.IdEspecialidadMedico = lnIdEspecialidadMedico
     mo_Atenciones.IdEstadoAtencion = sghEstadoTabla.sghRegistrado
     mo_Atenciones.IdFormaPago = lnIdTipoFinanciamiento
     mo_Atenciones.IdFuenteFinanciamiento = lnIdFuenteFinanciamiento
     mo_Atenciones.idMedicoIngreso = lnIdMedico
     'mo_Atenciones.IdMedicoEgreso
     mo_Atenciones.IdOrigenAtencion = 10   'domicilio
     mo_Atenciones.idPaciente = mo_Pacientes.idPaciente
     mo_Atenciones.IdServicioIngreso = lnIdServicio
     mo_Atenciones.IdTipoEdad = oEdad.TipoEdad
     mo_Atenciones.IdTipoServicio = 1
     mo_Atenciones.IdUsuarioAuditoria = ml_idUsuario
     mo_Atenciones.IdTipoCondicionALEstab = lnIdCondicionAlEstablecimiento
     mo_Atenciones.IdTipoCondicionAlServicio = lnIdCondicionAlServicio
     '
     mo_Cita.fecha = ldFechaCita
     mo_Cita.FechaSolicitud = ldFechaCita
     mo_Cita.HoraFin = lcHoraFinal
     mo_Cita.HoraInicio = lcHoraInicio
     mo_Cita.HoraSolicitud = lcHoraInicio
     mo_Cita.idAtencion = 0
     mo_Cita.IdCita = 0
     mo_Cita.IdEspecialidad = lnIdEspecialidadMedico
     mo_Cita.IdEstadoCita = 1
     mo_Cita.idMedico = lnIdMedico
     mo_Cita.idPaciente = mo_Pacientes.idPaciente
     mo_Cita.idProducto = lnIdProductoCpt
     mo_Cita.IdProgramacion = lnIdProgramacion
     mo_Cita.IdServicio = lnIdServicio
     mo_Cita.IdUsuarioAuditoria = ml_idUsuario
     '
     'mo_Diagnosticos               'no hay datos
     'mo_Procedimientos             'no hay datos
     'mo_Examenes                   'no hay datos
     'mo_DOFacturacionPaquetes      'no hay datos
     'oDoSunasaPacientesHistoricos  'no hay datos
     '
     Set mo_FacturacionServicios = New Collection
     mo_FacturacionServicios.Add CargarTipoDeConsulta(lnIdTipoFinanciamiento, lnIdFuenteFinanciamiento, lnIdProductoCpt)
     '
     lnTipoNumeracionAnterior = 1: lnIdHistoriaClinicaAnterior = 0: lbYaSeTransfirioHCdeUnServicioAotro = False
     
     If mo_AdminAdmision.AdmisionCEAgregar(mo_CuentasAtencion, mo_Atenciones, mo_Pacientes, mo_Cita, mo_Historia, _
                       lnTipoNumeracionAnterior, mo_Diagnosticos, mo_Procedimientos, mo_Examenes, _
                       lnIdHistoriaClinicaAnterior, mo_FacturacionServicios, mo_DOFacturacionPaquetes, _
                       lbYaSeTransfirioHCdeUnServicioAotro, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, _
                       "Agrega Cita Web", oDoSunasaPacientesHistoricos, mo_DoAtencionDatosAdicionales) = False Then
        ms_MensajeError = mo_AdminAdmision.MensajeError: GoTo ErrCrearCitaWeb
     End If
     If lbNecesitaCrearHistoria = True Then
        lnHistoriaCreada = mo_Pacientes.NroHistoriaClinica
     Else
        lnHistoriaCreada = 0
     End If
     CrearCitaAutomatica = True
ErrCrearCitaWeb:
     Set mo_AdminAdmision = Nothing
     Set mo_ReglasFacturacion = Nothing
     Set mo_ReglasServiciosHosp = Nothing
     Set mo_ReglasDeProgMedica = Nothing
     Set mo_ReglasComunes = Nothing
     Set mo_CuentasAtencion = Nothing
     Set mo_Atenciones = Nothing
     Set mo_Pacientes = Nothing
     Set mo_DOFacturacionPaquetes = Nothing
     Set oDoSunasaPacientesHistoricos = Nothing
     Set mo_Cita = Nothing
     Set mo_Historia = Nothing
     Set mo_DoAtencionDatosAdicionales = Nothing
     Set mo_FacturacionServicios = Nothing
     Set oPacientes = Nothing
     Set oHistoria = Nothing
     Set oRsTmp21 = Nothing
     Set oConexion = Nothing
End Function



Function CargarTipoDeConsulta(lnIdTipoFinanciamiento As Long, lnIdFuenteFinanciamiento As Long, _
                              lnIdProductoCpt As Long) As DOFacturacionServicios
Dim oConsulta As New DOFacturacionServicios
Dim oRsBuscaSeguro As New ADODB.Recordset
Dim mo_AdminFacturacion As New SIGHNegocios.ReglasFacturacion
Dim PrSeguro As Double
        With oConsulta
            .idAtencion = 0
            .IdFacturacionServicio = 0
            .idTipoFinanciamiento = lnIdTipoFinanciamiento
            .IdFuenteFinanciamiento = lnIdFuenteFinanciamiento
            .Cantidad = 1
            .idProducto = lnIdProductoCpt
            .IdUsuarioAuditoria = ml_idUsuario
            .IdEstadoFacturacion = sghEstadoFacturacion.sghRegistrado
            .FechaAutorizaPendiente = 0
            .FechaAutorizaSeguro = 0
            .IdCentroCosto = 0
            .IdUsuarioAutorizaPendiente = 0
            .IdUsuarioAutorizaSeguro = 0
            .PrecioUnitario = 0
            .TotalPorPagar = 0
            .idPuntoCarga = 6
            '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-inicio
'            If mi_Opcion = sghModificar Then
'               .IdOrden = lnIdFactServicios
'            End If
            PrSeguro = 0
            If lnIdTipoFinanciamiento = 9 Then
               'Si es EXONERACIONES tomará el PRECIO de un Paciente Normal
               Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(lnIdProductoCpt, 1)
            Else
               Set oRsBuscaSeguro = mo_AdminFacturacion.CatalogoServiciosHospSeleccionarXidProductoIdTipoFinanciamiento(lnIdProductoCpt, lnIdTipoFinanciamiento)
            End If
            If oRsBuscaSeguro.RecordCount > 0 Then
               PrSeguro = oRsBuscaSeguro.Fields!PrecioUnitario
            End If
            '.idTipoFinanciamiento = lnIdTipoFinanciamiento
            .CantidadSIS = 0
            .PrecioSIS = 0
            .ImporteSIS = 0
            .CantidadSOAT = 0
            .PrecioSOAT = 0
            .ImporteSOAT = 0
            .importeEXO = 0
            .cantidadConv = 0
            .precConv = 0
            .ImporteConv = 0
            Select Case lnIdTipoFinanciamiento
            Case 1  'Contado
            Case 2  'SIS
                 If PrSeguro > 0 Then
                    .CantidadSIS = 1
                    .PrecioSIS = PrSeguro
                    .ImporteSIS = PrSeguro
                    .IdEstadoFacturacion = 10
                    .FechaAutorizaSeguro = Now
                 Else
                 End If
            Case 3  'Soat
                 If PrSeguro > 0 Then
                    .CantidadSOAT = 1
                    .PrecioSOAT = PrSeguro
                    .ImporteSOAT = PrSeguro
                    .IdEstadoFacturacion = 10
                    .FechaAutorizaSeguro = Now
                 Else
                    .idTipoFinanciamiento = 1
                 End If
            Case 4  'Convenios
                 If PrSeguro > 0 Then
                    .cantidadConv = 1
                    .precConv = PrSeguro
                    .ImporteConv = PrSeguro
                    .IdEstadoFacturacion = 10
                    .PrecioUnitario = PrSeguro
                    .FechaAutorizaSeguro = Now
                 Else
                    .idTipoFinanciamiento = 1
                 End If
            Case 9  'Exonerados
                .importeEXO = PrSeguro
                .IdEstadoFacturacion = 10
                .idTipoFinanciamiento = 1
                .FechaAutorizaEXO = Now
            End Select
            '********pone Seguros en forma automatica, sin necesidad de ir a SEGUROS-fin
        End With
        Set CargarTipoDeConsulta = oConsulta
        oRsBuscaSeguro.Close
        Set oRsBuscaSeguro = Nothing
        Set mo_AdminFacturacion = Nothing
End Function

Function ExportaCitasWebSoloTablas(lbEnviarServicio As Boolean, lbEnviarPacientes As Boolean, _
                                   lbEnviarFteFinanciamiento As Boolean, lbEnviarMedicos As Boolean, _
                                   lbEnviarTurnos As Boolean) As Boolean
            On Error GoTo ErrCitasWebExp1
            ExportaCitasWebSoloTablas = False
            Dim oConexion As New Connection, oConexionExterna As New Connection
            Dim oRsCitasWebCupos As New Recordset, oRsCitaBloqueada As New Recordset
            Dim oRsCitaBloqueada1 As New Recordset
            Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
            Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
            Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
            Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
            Dim lcRutaExportar As String, lnTotalRegistros As Long, lnIdCitaBloqueada As Long
            lcRutaExportar = lcBuscaParametro.SeleccionaFilaParametro(313)
            'Exporta Servicios
            If lbEnviarServicio = True Then
                Set oRsCitasWebCupos = mo_ReglasArchivoClinico.ServiciosSeleccionarSoloCE
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Servicios.txt", oRsCitasWebCupos.RecordCount, 3
            End If
            'Exporta Medicos
            If lbEnviarMedicos = True Then
                Set oRsCitasWebCupos = mo_ReglasDeProgMedica.MedicosSeleccionarTodosParaWeb
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Medicos.txt", oRsCitasWebCupos.RecordCount, 4
            End If
            'Exporta Pacientes
            If lbEnviarPacientes = True Then
                Set oRsCitasWebCupos = mo_ReglasAdmision.PacientesSeleccionarTodos
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Pacientes.txt", oRsCitasWebCupos.RecordCount, 5
            End If
            'Exporta Fuente Financiamiento
            If lbEnviarFteFinanciamiento = True Then
                Set oRsCitasWebCupos = mo_ReglasComunes.FuentesFinanciamientoSeleccionarTodos
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "ftefinan.txt", oRsCitasWebCupos.RecordCount, 6
            End If
            If lbEnviarTurnos = True Then
                Set oRsCitasWebCupos = mo_ReglasDeProgMedica.TurnosSeleccionarPorIdTipoServicio(1)
                GeneraTXT oRsCitasWebCupos, lcRutaExportar & "Turnos.txt", oRsCitasWebCupos.RecordCount, 1
            End If
            MsgBox "Se exportó correctamente los archivos textos, en la Ruta: " & lcRutaExportar
            ExportaCitasWebSoloTablas = True
ErrCitasWebExp1:
       If ExportaCitasWebSoloTablas = True Then
       Else
       End If
       Set oRsCitasWebCupos = Nothing
       Set mo_ReglasDeProgMedica = Nothing
       Set mo_ReglasArchivoClinico = Nothing
       Set mo_ReglasAdmision = Nothing
       Set oRsCitaBloqueada1 = Nothing
End Function

Sub ImportaAcreditadosSIScursor(lcArchivoAcreditadosZIP As String)
     On Error GoTo ErrImporSIS
     Dim lcArchivoTmp As String, lcArchivo As String
     Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
     Dim Fichero As Integer, Registro As String, Campos() As String, fila As Single, Columna As Single
     Dim lnidSiasis As Long, lcCodigo As String, lcAfiliacionDisa As String, lcAfiliacionTipoFormato As String
     Dim lcAfiliacionNroFormato As String, lcAfiliacionNroIntegrante As String, lcDocumentoTipo As String
     Dim lcCodigoEstablAdscripcion As String, ldAfiliacionFecha As Date, lcPaterno As String, lcMaterno As String
     Dim lcPnombre As String, lcOnombres As String, lcGenero As String, ldFnacimiento As Date
     Dim lcIdDistritoDomicilio As String, lcEstado As String, ldFbaja As Date, lcDocumentoNumero As String
     Dim lcMotivoBaja As String, lcFecha As String, lcHoraInicioProceso As String, lcHoraFinalProceso As String
     Dim oRsTmp1 As New Recordset, oConexion As New Connection
     Dim lineas() As String, lnRegistros As Long, lnContador As Long
     Dim oCrypKey As New CrypKey.Util
     Dim oCommand As New ADODB.Command
     lcHoraInicioProceso = lcBuscaParametro.RetornaHoraServidorSQL1
     ms_MensajeError = ""
     lcArchivoTmp = "c:\debbSis"
     CreandoRutaEnDisco lcArchivoTmp
     SIGHEntidades.DescomprimeArchivoZIP oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(318)), lcArchivoAcreditadosZIP, lcArchivoTmp, False
     lcArchivo = Dir(lcArchivoTmp & "\*.*")
     SIGHEntidades.ejecutarComando "ren " & lcArchivo & " filiacionesSIS.txt"
     '
     oConexion.CommandTimeout = 300
     oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
     oConexion.CursorLocation = adUseClient
     With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = oConexion
       .CommandText = "SisFiliacionesImportaAcreditados"
       .Execute
     End With
     '
     lcSql = "update parametros set ValorTexto='Afiliados SIS: " & lcArchivo & "' where idParametro=312"
     If oRsTmp1.State = 1 Then oRsTmp1.Close
     oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
     '
     Set oRsTmp1 = Nothing
     oConexion.Close
     Set oConexion = Nothing
     Set mo_ReglasSISgalenhos = Nothing
     lcHoraFinalProceso = lcBuscaParametro.RetornaHoraServidorSQL1
     MsgBox "Proceso ya terminó en: " & DateDiff("s", CDate(lcHoraInicioProceso), CDate(lcHoraFinalProceso)), vbInformation, "SIS"
     
     Exit Sub
ErrImporSIS:
    ms_MensajeError = Err.Description
    MsgBox Err.Description
    
    'Resume

End Sub

Sub ImportaAcreditadosSIS(lcArchivoAcreditadosZIP As String)
     On Error GoTo ErrImporSIS
     Dim lcArchivoTmp As String, lcArchivo As String
     Dim mo_ReglasSISgalenhos As New ReglasSISgalenhos
     Dim Fichero As Integer, Registro As String, Campos() As String, fila As Single, Columna As Single
     Dim lnidSiasis As Long, lcCodigo As String, lcAfiliacionDisa As String, lcAfiliacionTipoFormato As String
     Dim lcAfiliacionNroFormato As String, lcAfiliacionNroIntegrante As String, lcDocumentoTipo As String
     Dim lcCodigoEstablAdscripcion As String, ldAfiliacionFecha As Date, lcPaterno As String, lcMaterno As String
     Dim lcPnombre As String, lcOnombres As String, lcGenero As String, ldFnacimiento As Date
     Dim lcIdDistritoDomicilio As String, lcEstado As String, ldFbaja As Date, lcDocumentoNumero As String
     Dim lcMotivoBaja As String, lcFecha As String, lcHoraInicioProceso As String, lcHoraFinalProceso As String
     Dim oRsTmp1 As New Recordset, oConexion As New Connection
     Dim lineas() As String, lnRegistros As Long, lnContador As Long
     Dim oCrypKey As New CrypKey.Util
     lcHoraInicioProceso = lcBuscaParametro.RetornaHoraServidorSQL1
     ms_MensajeError = ""
     lcArchivoTmp = "c:\debbSis"
     CreandoRutaEnDisco lcArchivoTmp
     SIGHEntidades.DescomprimeArchivoZIP oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(318)), lcArchivoAcreditadosZIP, lcArchivoTmp, False
     lcArchivo = Dir(lcArchivoTmp & "\*.*")
     '
     lineas() = Split(LeerArchivoTexto(lcArchivoTmp & "\" & lcArchivo), vbCrLf)
     lnRegistros = UBound(lineas)
     mo_ProgressRpt1.Min = 0: mo_ProgressRpt1.Max = lnRegistros: mo_ProgressRpt1.ShowText = True: mo_ProgressRpt1.Color = vbGreen
     oConexion.CommandTimeout = 300
     oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
     oConexion.CursorLocation = adUseClient
     Fichero = FreeFile
     Open lcArchivoTmp & "\" & lcArchivo For Input As #Fichero
     lnContador = 0
     While Not EOF(Fichero)
        mo_ProgressRpt1.Value = lnContador: lnContador = lnContador + 1: DoEvents
        Line Input #Fichero, Registro
        Campos = Split(Registro, ",")
        lnidSiasis = 0
        lcCodigo = ""
        lcAfiliacionDisa = ""
        lcAfiliacionTipoFormato = ""
        lcAfiliacionNroFormato = ""
        lcAfiliacionNroIntegrante = ""
        lcDocumentoTipo = ""
        lcCodigoEstablAdscripcion = ""
        ldAfiliacionFecha = 0
        lcPaterno = ""
        lcMaterno = ""
        lcPnombre = ""
        lcOnombres = ""
        lcGenero = ""
        ldFnacimiento = 0
        lcIdDistritoDomicilio = ""
        lcEstado = ""
        ldFbaja = 0
        lcDocumentoNumero = ""
        lcMotivoBaja = ""
        For Columna = 0 To UBound(Campos)
            Select Case Columna
            Case 0
                 lnidSiasis = Val(Campos(Columna))
            Case 1
                lcCodigo = Campos(Columna)
            Case 2
                lcAfiliacionDisa = Campos(Columna)
            Case 3
                lcAfiliacionTipoFormato = Campos(Columna)
            Case 4
                lcAfiliacionNroFormato = Campos(Columna)
            Case 5
                lcAfiliacionNroIntegrante = Campos(Columna)
            Case 6
                lcDocumentoTipo = Campos(Columna)
            Case 7
                lcCodigoEstablAdscripcion = Campos(Columna)
            Case 8
                lcFecha = Campos(Columna)
                lcFecha = Right(lcFecha, 2) & "/" & Mid(lcFecha, 5, 2) & "/" & Left(lcFecha, 4)
                If SIGHEntidades.EsFecha(lcFecha, "DD/MM/AAAA") Then
                   ldAfiliacionFecha = CDate(lcFecha)
                End If
            Case 9
                lcPaterno = Campos(Columna)
            Case 10
                lcMaterno = Campos(Columna)
            Case 11
                lcPnombre = Campos(Columna)
            Case 12
                lcOnombres = Campos(Columna)
            Case 13
                lcGenero = Campos(Columna)
            Case 14
                lcFecha = Campos(Columna)
                lcFecha = Right(lcFecha, 2) & "/" & Mid(lcFecha, 5, 2) & "/" & Left(lcFecha, 4)
                If SIGHEntidades.EsFecha(lcFecha, "DD/MM/AAAA") Then
                    ldFnacimiento = CDate(lcFecha)
                End If
            Case 15
                lcIdDistritoDomicilio = Campos(Columna)
            Case 16
                lcEstado = Campos(Columna)
            Case 17
                lcFecha = Campos(Columna)
                lcFecha = Right(lcFecha, 2) & "/" & Mid(lcFecha, 5, 2) & "/" & Left(lcFecha, 4)
                If SIGHEntidades.EsFecha(lcFecha, "DD/MM/AAAA") Then
                    ldFbaja = lcFecha
                 End If
            Case 18
                lcDocumentoNumero = Campos(Columna)
            Case 19
                lcMotivoBaja = Campos(Columna)
            End Select
        Next
        Set oRsTmp1 = mo_ReglasSISgalenhos.SisFiliacionesSeleccionarPorIdSiaSis(lnidSiasis, lcCodigo, oConexion)
        If oRsTmp1.RecordCount = 0 Then
           oRsTmp1.AddNew
           oRsTmp1.Fields!idSiasis = lnidSiasis
        End If
        oRsTmp1.Fields!Codigo = lcCodigo
        oRsTmp1.Fields!AfiliacionDisa = lcAfiliacionDisa
        oRsTmp1.Fields!AfiliacionTipoFormato = lcAfiliacionTipoFormato
        oRsTmp1.Fields!AfiliacionNroFormato = lcAfiliacionNroFormato
        oRsTmp1.Fields!AfiliacionNroIntegrante = lcAfiliacionNroIntegrante
        oRsTmp1.Fields!DocumentoTipo = lcDocumentoTipo
        oRsTmp1.Fields!CodigoEstablAdscripcion = lcCodigoEstablAdscripcion
        If ldAfiliacionFecha <> 0 Then
           oRsTmp1.Fields!AfiliacionFecha = ldAfiliacionFecha
        End If
        oRsTmp1.Fields!Paterno = lcPaterno
        oRsTmp1.Fields!Materno = lcMaterno
        oRsTmp1.Fields!Pnombre = lcPnombre
        oRsTmp1.Fields!Onombres = lcOnombres
        oRsTmp1.Fields!Genero = lcGenero
        If ldFnacimiento <> 0 Then
           oRsTmp1.Fields!fnacimiento = ldFnacimiento
        End If
        oRsTmp1.Fields!IdDistritoDomicilio = lcIdDistritoDomicilio
        oRsTmp1.Fields!estado = lcEstado
        If ldFbaja <> 0 Then
           oRsTmp1.Fields!Fbaja = ldFbaja
        End If
        oRsTmp1.Fields!DocumentoNumero = lcDocumentoNumero
        oRsTmp1.Fields!MotivoBaja = lcMotivoBaja
        oRsTmp1.Update
        
        
        fila = fila + 1
    Wend
    Close #Fichero
    '
    lcSql = "update parametros set ValorTexto='Afiliados SIS: " & lcArchivo & "' where idParametro=312"
    If oRsTmp1.State = 1 Then oRsTmp1.Close
    oRsTmp1.Open lcSql, SIGHEntidades.CadenaConexion, adOpenKeyset, adLockOptimistic
    '
    
    Set oRsTmp1 = Nothing
    oConexion.Close
    Set oConexion = Nothing
    Set mo_ReglasSISgalenhos = Nothing
     lcHoraFinalProceso = lcBuscaParametro.RetornaHoraServidorSQL1
     MsgBox "Proceso ya terminó en: " & DateDiff("s", CDate(lcHoraInicioProceso), CDate(lcHoraFinalProceso)), vbInformation, "SIS"
    Exit Sub
ErrImporSIS:
    ms_MensajeError = Err.Description
    MsgBox Err.Description
    
    'Resume

End Sub

Sub ExportaSIS(lcFechaInicio As String, lcFechaFinal As String)
        On Error GoTo ErrorProceso
        Dim lnIdCuentaInicial As Long, lnIdCuentaFinal As Long
        Dim oRsTmp1 As New Recordset
        Dim oRsTmp2 As New Recordset
        Dim mrs_Tmp As New Recordset
        Dim lnTotalReg As Long, lbProcesoOK As Boolean
        Dim oConexionExterna As New Connection
        Dim oConexion As New Connection
        Dim lbSeImprime As Boolean, lcRutaExportar As String, lcArchivoExpZip As String
        Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
        Dim oCrypKey As New CrypKey.Util
        oConexion.CommandTimeout = 300
        oConexion.Open SIGHEntidades.CadenaConexion
        oConexion.CursorLocation = adUseClient
        oConexionExterna.CommandTimeout = 300
        oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
        oConexionExterna.CursorLocation = adUseClient
        Set oRsTmp1 = ChequeaCuentasSinFUAgrabado(mrs_Tmp, oConexionExterna, lcFechaInicio, lcFechaFinal)
        lbSeImprime = True
        lbProcesoOK = False
        If mrs_Tmp.RecordCount = 0 Then
            MsgBox "No existen Datos en ese rango de Fechas", vbInformation, "Mensaje"
            lbSeImprime = False
        ElseIf mrs_Tmp.RecordCount > 0 Then
            CrearReporteObservaciones_excel mrs_Tmp
            If MsgBox("Existen Cuentas que no se han grabado en FUA, desea exportar FUA de todas maneras ?", vbQuestion + vbYesNo, "Mensaje") <> vbYes Then
               lbSeImprime = False
            End If
        End If
        If lbSeImprime = True Then
            lcRutaExportar = lcBuscaParametro.SeleccionaFilaParametro(313)
            '
            oRsTmp1.MoveFirst
            lnIdCuentaInicial = oRsTmp1.Fields!idCuentaAtencion
            oRsTmp1.MoveLast
            lnIdCuentaFinal = oRsTmp1.Fields!idCuentaAtencion
            oRsTmp1.Close
            'SisFuaAtencion
            lcSql = "select * from sisFuaAtencion where idCuentaAtencion>=" & lnIdCuentaInicial & " and idCuentaAtencion<=" & lnIdCuentaFinal
            oRsTmp1.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
            lnTotalReg = oRsTmp1.RecordCount
            If lnTotalReg > 0 Then
               GeneraTXT oRsTmp1, lcRutaExportar & "\sisAtencion.txt", lnTotalReg, 1
            End If
            oRsTmp1.Close
            'SisFuaAtencionDIA
            lcSql = "select * from sisFuaAtencionDIA where idCuentaAtencion>=" & lnIdCuentaInicial & " and idCuentaAtencion<=" & lnIdCuentaFinal
            oRsTmp1.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
            lnTotalReg = oRsTmp1.RecordCount
            If lnTotalReg > 0 Then
               GeneraTXT oRsTmp1, lcRutaExportar & "\sisAtencionDIA.txt", lnTotalReg, 2
            End If
            oRsTmp1.Close
            'SisFuaAtencionINS
            lcSql = "select * from sisFuaAtencionINS where idCuentaAtencion>=" & lnIdCuentaInicial & " and idCuentaAtencion<=" & lnIdCuentaFinal
            oRsTmp1.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
            lnTotalReg = oRsTmp1.RecordCount
            If lnTotalReg > 0 Then
               GeneraTXT oRsTmp1, lcRutaExportar & "\sisAtencionINS.txt", lnTotalReg, 3
            End If
            oRsTmp1.Close
            'SisFuaAtencionMED
            lcSql = "select * from sisFuaAtencionMED where idCuentaAtencion>=" & lnIdCuentaInicial & " and idCuentaAtencion<=" & lnIdCuentaFinal
            oRsTmp1.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
            lnTotalReg = oRsTmp1.RecordCount
            If lnTotalReg > 0 Then
               GeneraTXT oRsTmp1, lcRutaExportar & "\sisAtencionMED.txt", lnTotalReg, 4
            End If
            oRsTmp1.Close
            'SisFuaAtencionPRO
            lcSql = "select * from sisFuaAtencionPRO where idCuentaAtencion>=" & lnIdCuentaInicial & " and idCuentaAtencion<=" & lnIdCuentaFinal
            oRsTmp1.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
            lnTotalReg = oRsTmp1.RecordCount
            If lnTotalReg > 0 Then
               GeneraTXT oRsTmp1, lcRutaExportar & "\sisAtencionPRO.txt", lnTotalReg, 5
            End If
            oRsTmp1.Close
            'SisFuaAtencionSMI
            lcSql = "select * from sisFuaAtencionSMI where idCuentaAtencion>=" & lnIdCuentaInicial & " and idCuentaAtencion<=" & lnIdCuentaFinal
            oRsTmp1.Open lcSql, oConexionExterna, adOpenKeyset, adLockOptimistic
            lnTotalReg = oRsTmp1.RecordCount
            If lnTotalReg > 0 Then
               GeneraTXT oRsTmp1, lcRutaExportar & "\sisAtencionSMI.txt", lnTotalReg, 6
            End If
            oRsTmp1.Close
            '
            lbProcesoOK = True
            Call mo_ReglasSeguridad.AuditoriaAgregarV(ml_idUsuario, "A", 0, "FUA", oConexion, 500 + 180, mo_lcNombrePc, "Exportar al SIS: " & lcFechaInicio & " al " & lcFechaFinal)        '500+ ListBarReporte.idReporte
        End If
        Set oRsTmp1 = Nothing
        Set mrs_Tmp = Nothing
        Set oConexionExterna = Nothing
        If lbProcesoOK = True Then
            lcArchivoExpZip = "c:\sis.zip"
            SIGHEntidades.ComprimeArchivoZip lcArchivoExpZip, oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(319)), lcRutaExportar & "\sis*.txt"
            MsgBox "Se creó el archivo: " & lcArchivoExpZip
        End If
        Set oCrypKey = Nothing
    Exit Sub
ErrorProceso:
    MsgBox Err.Description
    Resume
End Sub

