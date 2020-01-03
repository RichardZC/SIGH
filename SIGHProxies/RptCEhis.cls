VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptCEhis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para emisión del Formato HIS
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mrs_Tmp As New ADODB.Recordset
Dim ml_IdServicioCE As Long
Dim ml_IdResponsable As Long
Dim mo_Conexion As ADODB.Connection
Dim mo_ReporteUtil As New ReporteUtil
Dim ml_OrdenFiltro As String
Dim mb_SonDatosDePruebas As Boolean
Dim ml_TextoDelFiltro As String
Dim ml_TextoDelFiltro1 As String
Dim mda_FechaInicio As String
Dim mda_FechaFin As String
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_AgregaCPT As Boolean
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes
Dim mo_ReglasDeProgMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision

Property Let AgregaCPT(daValue As Boolean)
    ml_AgregaCPT = daValue
End Property

Property Let FechaInicio(daValue As String)
    mda_FechaInicio = daValue
End Property
Property Let FechaFin(daValue As String)
    mda_FechaFin = daValue
End Property
Property Let TextoDelFiltro(lValue As String)
    ml_TextoDelFiltro = lValue
End Property
Property Let TextoDelFiltro1(lValue As String)
    ml_TextoDelFiltro1 = lValue
End Property
Property Let SonDatosDePruebas(lValue As Boolean)
    mb_SonDatosDePruebas = lValue
End Property

Property Let IdServicioCE(lValue As Long)
    ml_IdServicioCE = lValue
End Property
Property Let IdResponsable(lValue As Long)
    ml_IdResponsable = lValue
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Let OrdenFiltro(lValue As String)
    ml_OrdenFiltro = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New CEhis
    oFormulario.Show 1
End Sub

'Frank 15092014
Sub CrearReporte_excel(lnHwnd As Long, lbAgregaOrdenesMedicas As Boolean, lbHaceVistaPrevia As Boolean)
Dim rsReporte As New Recordset
Dim iFila As Long
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim oDOPaciente As New DOPaciente
Dim oDOAtencion As New DOAtencion

Dim lnLineas As Integer: Dim lnIdAtencion  As Long
Dim lcProcedencia As String, lcHcFicha As String, lbEsCS As Boolean
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String
Dim lnSeccionDx As Integer 'Frank 15092014
Dim rsDx As New Recordset
Dim rsCpt As New Recordset
Dim rsDx1 As New Recordset
Dim lcCodigoHIsResponsable As String 'Actualizado 14102014
Dim rsDatosMedico As Recordset 'Actualizado 14102014
Dim lnoRsImpHIStodos As Long
Dim oRsImpHIStodos As New Recordset
Dim oConexionExterna As New Connection
Dim lcAD040 As String, lcFilter As String
'
lcAD040 = lcBuscaParametro.SeleccionaFilaParametro(354)
'
oConexionExterna.CursorLocation = adUseClient
oConexionExterna.CommandTimeout = 150
oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorErrorExcel

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
        
    Dim lnHojas As Long
    Dim lbCargaUnaVezPorHoja As Boolean
    Dim lnTotalDxCpt As Long
        
    'Filtra los Datos
    Set mrs_Tmp = mo_AdminAdmision.BuscaAtencionesCEparaFormatoHIS(ml_IdServicioCE, ml_IdResponsable, mda_FechaInicio, mda_FechaFin, "")
    
    Set rsDatosMedico = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(ml_IdResponsable)
'    lcCodigoHIsResponsable ml_IdResponsable
    'Actualizado 14102014
    If rsDatosMedico.RecordCount > 0 Then
        rsDatosMedico.MoveFirst
        Do While Not rsDatosMedico.EOF
            lcCodigoHIsResponsable = "000000000"
            If Not IsNull(rsDatosMedico.Fields!dni) Then
                If Trim(rsDatosMedico.Fields!dni) <> "" Then
                    lcCodigoHIsResponsable = "1" & Trim(rsDatosMedico.Fields!dni)
                End If
            End If
            lcCodigoHIsResponsable = lcCodigoHIsResponsable & Right("00" & rsDatosMedico.Fields!TipoEmpleadoSIS, 2)
            rsDatosMedico.MoveNext
        Loop
    End If
    '''''''''''''''''''''
    
    If mrs_Tmp.RecordCount = 0 Then
        If lbHaceVistaPrevia = True Then
           MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        End If
    Else
            Set oRsImpHIStodos = mo_AdminAdmision.ServiciosAtenSimultaneaImpHISxUPS(mo_AdminAdmision.BuscaUPSactualDelPaciente(ml_IdServicioCE))
            lnoRsImpHIStodos = oRsImpHIStodos.RecordCount
            lbEsCS = IIf(lcBuscaParametro.SeleccionaFilaParametro(282) = "S", True, False)
            lbCargaUnaVezPorHoja = True
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
                If lbCargaUnaVezPorHoja = True Then
                    If lbEsOpenOffice = True Then
                        'Abre el archivo ExcelOpenOffice
                        lcArchivoExcel = App.Path + "\Plantillas\CEHis.ods"

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
                        ret = 0 ' SetForegroundWindow(lnHwnd)
                    Else
                        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                        'Crea nueva hoja
                        Set oWorkBook = oExcel.Workbooks.Add
                        'Abre, copia y cierra la plantilla
                        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEHis.xls")
                        oWorkBookPlantilla.Worksheets("his").Copy Before:=oWorkBook.Sheets(1)
                        oWorkBookPlantilla.Close
                        'Activa la primera hoja
                        Set oWorkSheet = oWorkBook.Sheets(1)
                        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
                    End If
                    Set rsReporte = lcBuscaParametro.DevuelveUbigeoDetalladoDelHospital()
                    If rsReporte.RecordCount > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(1, 1).setFormula(lcBuscaParametro.SeleccionaFilaParametro(205))
                            Call Feuille.getcellbyposition(16, 4).setFormula("'" & Trim(str(rsReporte.Fields!IdDepartamento)) & " (" & Trim(rsReporte.Fields!ndpto) & ")")
                            Call Feuille.getcellbyposition(6, 5).setFormula("5 UPS")
                        Else
                            oWorkSheet.Cells(2, 2).Value = lcBuscaParametro.SeleccionaFilaParametro(205)
                            oWorkSheet.Cells(5, 17).Value = "5 UPS"
                            oWorkSheet.Cells(6, 7).Value = "'" & Trim(str(rsReporte.Fields!IdDepartamento)) & " (" & Trim(rsReporte.Fields!ndpto) & ")"
                        End If
                        
                        lcProcedencia = Trim(str(rsReporte.Fields!IdDepartamento))
                        lcPaciente = Trim(str(rsReporte.Fields!IdProvincia))
                        
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(7, 5).setFormula("'" & Mid(lcPaciente, Len(lcProcedencia) + 1, 10))
                        Else
                            oWorkSheet.Cells(6, 8).Value = "'" & Mid(lcPaciente, Len(lcProcedencia) + 1, 10)
                        End If
                        lcProcedencia = Trim(str(rsReporte.Fields!IdDistrito))
                        
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(9, 5).setFormula("'" & Mid(lcProcedencia, Len(lcPaciente) + 1, 10))
                            Call Feuille.getcellbyposition(11, 5).setFormula("'" & rsReporte.Fields!Codigo)
                        Else
                            oWorkSheet.Cells(6, 10).Value = "'" & Mid(lcProcedencia, Len(lcPaciente) + 1, 10)
                            oWorkSheet.Cells(6, 12).Value = "'" & rsReporte.Fields!Codigo
                            
                        End If
                    End If
                    rsReporte.Close
                    
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(16, 5).setFormula(ml_TextoDelFiltro1)
                        Call Feuille.getcellbyposition(18, 5).setFormula(UCase(ml_TextoDelFiltro))
                        Call Feuille.getcellbyposition(21, 5).setFormula(lcCodigoHIsResponsable) 'Actualizado 14102014
                        Call Feuille.getcellbyposition(2, 6).setFormula(Year(CDate(mda_FechaInicio)))
                        Call Feuille.getcellbyposition(3, 6).setFormula(Month(CDate(mda_FechaInicio)))
                    Else
                        oWorkSheet.Cells(6, 17).Value = ml_TextoDelFiltro1
                        oWorkSheet.Cells(6, 19).Value = UCase(ml_TextoDelFiltro)
                        oWorkSheet.Cells(6, 22).Value = lcCodigoHIsResponsable 'Actualizado 14102014
                        oWorkSheet.Cells(7, 3).Value = Year(CDate(mda_FechaInicio))
                        oWorkSheet.Cells(7, 4).Value = Month(CDate(mda_FechaInicio))
                    End If
                    iFila = 13
                    lbCargaUnaVezPorHoja = False
                End If 'Fin de carga una vez por hoja his
                

                'CIE10
                lcFilter = "idPuntoCArga=1"
                If SighEntidades.Parametro561 = "S" Then
                   lcFilter = lcFilter + " or idPuntoCarga=99"
                End If
                If lnoRsImpHIStodos = 0 Then
                    Set rsDx = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(mrs_Tmp.Fields!idAtencion)
                    lnTotalDxCpt = rsDx.RecordCount
                     'CPT
                     If lbAgregaOrdenesMedicas = True Or ml_AgregaCPT = True Then
                        Set rsCpt = mo_AdminAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(mrs_Tmp.Fields!idCuentaAtencion, 0)
                        If lbAgregaOrdenesMedicas = False Then
                           rsCpt.Filter = lcFilter
                        End If
                        lnTotalDxCpt = lnTotalDxCpt + rsCpt.RecordCount
                    End If
                 Else
                    Set rsDx1 = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(mrs_Tmp.Fields!idAtencion)
                    'CPT
                    If lbAgregaOrdenesMedicas = True Or ml_AgregaCPT = True Then
                        Set rsCpt = mo_AdminAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(mrs_Tmp.Fields!idCuentaAtencion, 0)
                        If lbAgregaOrdenesMedicas = False Then
                           rsCpt.Filter = lcFilter
                        End If
                    End If
                    mo_AdminAdmision.CreaTemporalYcargaDxYcptXgrupo rsDx, rsDx1, rsCpt, mrs_Tmp!idAtencion, True, lcAD040
                    lnTotalDxCpt = rsDx.RecordCount
               End If
               '
               If lnTotalDxCpt > 0 Then
                   lcProcedencia = ""
                   If Not IsNull(mrs_Tmp.Fields!distrito) Then
                      lcProcedencia = lcProcedencia & Trim(mrs_Tmp.Fields!distrito)
                   End If
                   If Not IsNull(mrs_Tmp.Fields!provincia) Then
                      lcProcedencia = lcProcedencia & " (" & Trim(mrs_Tmp.Fields!provincia) & ")"
                   End If
                   If Not IsNull(mrs_Tmp.Fields!dpto) Then
                      lcProcedencia = lcProcedencia & " (" & Trim(mrs_Tmp.Fields!dpto) & ")"
                   End If
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(Day(mrs_Tmp.Fields("fechaIngreso").Value))
                    Else
                        oWorkSheet.Cells(iFila, 3).Value = Day(mrs_Tmp.Fields("fechaIngreso").Value)
                    End If
                   lcHcFicha = mrs_Tmp.Fields("nroHistoriaClinica").Value
                   If Not IsNull(mrs_Tmp.Fields("fichaFamiliar").Value) And lbEsCS = True Then
                      lcHcFicha = "FF:" & mrs_Tmp.Fields("fichaFamiliar").Value
                   End If
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lcHcFicha)
                    Else
                        oWorkSheet.Cells(iFila, 4).Value = lcHcFicha
                    End If
                   If mrs_Tmp.Fields("IdDocIdentidad").Value = 1 And Len(mrs_Tmp.Fields("NroDocumento").Value) > 0 Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(3, iFila - 1).setFormula(lcHcFicha & "    Dni:" & mrs_Tmp.Fields("NroDocumento").Value)
                    Else
                        oWorkSheet.Cells(iFila, 4).Value = lcHcFicha & "    Dni:" & mrs_Tmp.Fields("NroDocumento").Value
                    End If
                   End If
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(mrs_Tmp.Fields("CodigoHIS").Value)
                        Call Feuille.getcellbyposition(5, iFila - 1).setFormula(mrs_Tmp.Fields("IdEtnia").Value)
                        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lcProcedencia)
                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(mrs_Tmp.Fields("edad").Value & " (" & Trim(mrs_Tmp.Fields!tEdad) & ")")
                        Call Feuille.getcellbyposition(8, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoSexo").Value = 1, "X", ""))
                        Call Feuille.getcellbyposition(9, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoSexo").Value = 1, "", "X"))
                    Else
                        oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("CodigoHIS").Value
                        oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp.Fields("IdEtnia").Value
                        oWorkSheet.Cells(iFila, 7).Value = lcProcedencia
                        oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp.Fields("edad").Value & " (" & Trim(mrs_Tmp.Fields!tEdad) & ")"
                        oWorkSheet.Cells(iFila, 9).Value = IIf(mrs_Tmp.Fields("idTipoSexo").Value = 1, "X", "")
                        oWorkSheet.Cells(iFila, 10).Value = IIf(mrs_Tmp.Fields("idTipoSexo").Value = 1, "", "X")
                    End If
If mrs_Tmp!idCuentaAtencion = 34977 Then
ml_TextoDelFiltro1 = "xxx"
End If
                    
                   If Not IsNull(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value) Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(13, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 1, "X", ""))    'Nuevo
                        Call Feuille.getcellbyposition(14, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 3, "X", ""))   'Continuador
                        Call Feuille.getcellbyposition(15, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 2, "X", ""))   'Reingreso)
                    Else
                      oWorkSheet.Cells(iFila, 14).Value = IIf(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 1, "X", "")    'Nuevo
                      oWorkSheet.Cells(iFila, 15).Value = IIf(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 3, "X", "")   'Continuador
                      oWorkSheet.Cells(iFila, 16).Value = IIf(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 2, "X", "")   'Reingreso
                    End If
                   End If
                   If Not IsNull(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value) Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(10, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 1, "X", ""))    'Nuevo
                        Call Feuille.getcellbyposition(11, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 3, "X", ""))   'Continuador
                        Call Feuille.getcellbyposition(12, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 2, "X", ""))     'Reingreso)
                    Else
                      oWorkSheet.Cells(iFila, 11).Value = IIf(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 1, "X", "")      'Nuevo
                      oWorkSheet.Cells(iFila, 12).Value = IIf(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 3, "X", "")      'Continuador
                      oWorkSheet.Cells(iFila, 13).Value = IIf(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 2, "X", "")      'Reingreso
                    End If
                   End If
                   lnLineas = 1
                   lnSeccionDx = 0
                   
                   If rsDx.RecordCount > 0 Then
                        rsDx.MoveFirst
                        Do While Not rsDx.EOF
    '                       If lnLineas <= 6 Then
    '                        lnSeccionDx = IIf(Round(lnLineas / 6) < lnLineas / 6, Round(lnLineas / 6) + 1, Round(lnLineas / 6))
                            lnLineas = IIf(lnLineas Mod 7 = 0, lnLineas + 1, lnLineas)
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(16, iFila + lnLineas - 2).setFormula(rsDx.Fields("dDiagnostico").Value)
                                Call Feuille.getcellbyposition(19, iFila + lnLineas - 2).setFormula(IIf(rsDx.Fields("idSubClasificacionDx").Value = 101, "X", ""))  'Presuntivo)
                                Call Feuille.getcellbyposition(20, iFila + lnLineas - 2).setFormula(IIf(rsDx.Fields("idSubClasificacionDx").Value = 102, "X", ""))
                                Call Feuille.getcellbyposition(21, iFila + lnLineas - 2).setFormula(IIf(rsDx.Fields("idSubClasificacionDx").Value = 103, "X", ""))
                                Call Feuille.getcellbyposition(22, iFila + lnLineas - 2).setFormula(Trim(IIf(IsNull(rsDx.Fields("labConfHIS").Value), "", rsDx.Fields("labConfHIS").Value)))
                                Call Feuille.getcellbyposition(23, iFila + lnLineas - 2).setFormula(Trim(rsDx.Fields("codigoCie10").Value))
                            Else
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 17).Value = rsDx.Fields("dDiagnostico").Value
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 20).Value = IIf(rsDx.Fields("idSubClasificacionDx").Value = 101, "X", "") 'Presuntivo
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 21).Value = IIf(rsDx.Fields("idSubClasificacionDx").Value = 102, "X", "") 'Definitivo
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 22).Value = IIf(rsDx.Fields("idSubClasificacionDx").Value = 103, "X", "") 'Repetido
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 23).Value = Trim(rsDx.Fields("labConfHIS").Value)
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 24).Value = Trim(rsDx.Fields("codigoCie10").Value)
                            End If
                            
    '                        End If
                           lnLineas = lnLineas + 1
                           rsDx.MoveNext
                        Loop
                   End If
                   rsDx.Close
                   'CPT
                   If ml_AgregaCPT = True Then
                        'Set rsCpt = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(mrs_Tmp.Fields!idCuentaAtencion)
                        If rsCpt.RecordCount > 0 Then
                             rsCpt.MoveFirst
                             Do While Not rsCpt.EOF
    '                           If lnLineas <= 6 Then
                                lnLineas = IIf(lnLineas Mod 7 = 0, lnLineas + 1, lnLineas)
                                If lbEsOpenOffice = True Then
                                    'Actualizado 22092014
                                    Call Feuille.getcellbyposition(16, iFila + lnLineas - 2).setFormula(rsCpt.Fields("Nombre").Value)
                                    Call Feuille.getcellbyposition(20, iFila + lnLineas - 2).setFormula("X")
                                    Call Feuille.getcellbyposition(23, iFila + lnLineas - 2).setFormula(Trim(rsCpt.Fields("codigo").Value))
                                    Call Feuille.getcellbyposition(22, iFila + lnLineas - 2).setFormula(Trim(rsCpt.Fields("labConfHIS").Value))
                                    '''''''''''''''''''''
                                Else
                                    oWorkSheet.Cells(iFila + lnLineas - 1, 17).Value = rsCpt.Fields("Nombre").Value
                                    oWorkSheet.Cells(iFila + lnLineas - 1, 21).Value = "X"
                                    oWorkSheet.Cells(iFila + lnLineas - 1, 24).Value = Trim(rsCpt.Fields("codigo").Value)
                                    oWorkSheet.Cells(iFila + lnLineas - 1, 23).Value = Trim(rsCpt.Fields("labConfHIS").Value)
                                End If
    '                           End If
                                lnLineas = lnLineas + 1
                                rsCpt.MoveNext
                             Loop
                        End If
                        rsCpt.Close
                   End If
                   lnLineas = lnLineas - 1
                   '32   '41-34
                    iFila = iFila + lnLineas + (7 - lnLineas Mod 7)
                    If Round((iFila - 12) / 7) < (iFila - 12) / 7 Then
                         If Round((iFila - 12) / 7) + 1 = 25 Then
                             lbCargaUnaVezPorHoja = True
                         End If
                    Else
                         If Round((iFila - 12) / 7) = 25 Then
                             lbCargaUnaVezPorHoja = True
                         End If
                    End If
               End If
               mrs_Tmp.MoveNext
'               iFila = iFila + 7
            Loop

            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 24
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$10"
                If oWorkSheet.PageSetup.PrintArea <> "" Then
                   oWorkSheet.PageSetup.PrintArea = SighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila + 3)
                End If
                If lbHaceVistaPrevia = False Then
                   oWorkSheet.PrintOut
                Else
                   oExcel.Visible = True
                   oWorkSheet.PrintPreview
                End If
            End If
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
                Set oWorkSheet = Nothing
                Set oExcel = Nothing
                Set oWorkBookPlantilla = Nothing
                Set oWorkBook = Nothing
            End If
        
        End If
        oConexionExterna.Close
        Set oConexionExterna = Nothing
Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    
    Exit Sub
    Resume
End Sub





Sub CrearReporte_excel2017(lnHwnd As Long, lbAgregaOrdenesMedicas As Boolean, lcHoraInicio As String, lcHoraFinal As String)
Dim iFila As Long
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim lnLineas As Integer: Dim lnIdAtencion  As Long
Dim lcProcedencia As String, lcHcFicha As String, lbEsCS As Boolean
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String
Dim lnSeccionDx As Integer 'Frank 15092014
Dim oDOPaciente As New DOPaciente
Dim oDOAtencion As New DOAtencion
Dim rsReporte As New Recordset
Dim mo_ReglasLaboratorio As New ReglasLaboratorio
Dim rsDx As New Recordset
Dim rsCpt As New Recordset
Dim rsDx1 As New Recordset
Dim oRsTmp987 As New Recordset
Dim mo_ReglasCaja As New ReglasCaja

Dim rsDatosMedico As Recordset 'Actualizado 14102014
Dim oRsImpHIStodos As New Recordset
Dim oConexionExterna As New Connection
Dim oConexion As New Connection
Dim lcCodigoHIsResponsable As String 'Actualizado 14102014
Dim lnoRsImpHIStodos As Long
Dim lcAD040 As String
Dim ldFechaHemoglobina As Date, lcHemoglobina As String, lcFilter As String

'
lcAD040 = lcBuscaParametro.SeleccionaFilaParametro(354)
'
oConexionExterna.CursorLocation = adUseClient
oConexionExterna.CommandTimeout = 150
oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'
oConexion.CursorLocation = adUseClient
oConexion.CommandTimeout = 150
oConexion.Open SighEntidades.CadenaConexion

lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorErrorExcel

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
        
    Dim lnHojas As Long
    Dim lbCargaUnaVezPorHoja As Boolean
    Dim lnTotalDxCpt As Long
        
    'Filtra los Datos
    Set mrs_Tmp = mo_AdminAdmision.BuscaAtencionesCEparaFormatoHIS(ml_IdServicioCE, ml_IdResponsable, mda_FechaInicio, mda_FechaFin, "")
    lcHemoglobina = "HoraIngreso>='" & lcHoraInicio & "' and  HoraIngreso<='" & lcHoraFinal & "'"
    mrs_Tmp.Filter = lcHemoglobina
    Set rsDatosMedico = mo_ReglasDeProgMedica.MedicosSeleccionarXIdMedico(ml_IdResponsable)
'    lcCodigoHIsResponsable ml_IdResponsable
    'Actualizado 14102014
    If rsDatosMedico.RecordCount > 0 Then
        rsDatosMedico.MoveFirst
        Do While Not rsDatosMedico.EOF
            lcCodigoHIsResponsable = "000000000"
            If Not IsNull(rsDatosMedico.Fields!dni) Then
                If Trim(rsDatosMedico.Fields!dni) <> "" Then
                    lcCodigoHIsResponsable = "1" & Trim(rsDatosMedico.Fields!dni)
                End If
            End If
            lcCodigoHIsResponsable = lcCodigoHIsResponsable & Right("00" & rsDatosMedico.Fields!TipoEmpleadoSIS, 2)
            rsDatosMedico.MoveNext
        Loop
    End If
    '''''''''''''''''''''
    
    If mrs_Tmp.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
            
            Set oRsImpHIStodos = mo_AdminAdmision.ServiciosAtenSimultaneaImpHISxUPS(mo_AdminAdmision.BuscaUPSactualDelPaciente(ml_IdServicioCE))
            lnoRsImpHIStodos = oRsImpHIStodos.RecordCount
            lbEsCS = IIf(lcBuscaParametro.SeleccionaFilaParametro(282) = "S", True, False)
            lbCargaUnaVezPorHoja = True
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
                If lbCargaUnaVezPorHoja = True Then
                    If lbEsOpenOffice = True Then
                        'Abre el archivo ExcelOpenOffice
                        lcArchivoExcel = App.Path + "\Plantillas\CEHis.ods"

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
                        Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                        'Crea nueva hoja
                        Set oWorkBook = oExcel.Workbooks.Add
                        'Abre, copia y cierra la plantilla
                        Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEHis2017.xls")
                        oWorkBookPlantilla.Worksheets("his").Copy Before:=oWorkBook.Sheets(1)
                        oWorkBookPlantilla.Close
                        'Activa la primera hoja
                        Set oWorkSheet = oWorkBook.Sheets(1)
                        mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
                    End If
                    Set rsReporte = lcBuscaParametro.DevuelveUbigeoDetalladoDelHospital()
                    If rsReporte.RecordCount > 0 Then
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(1, 1).setFormula(lcBuscaParametro.SeleccionaFilaParametro(205))
                            Call Feuille.getcellbyposition(16, 4).setFormula("'" & Trim(str(rsReporte.Fields!IdDepartamento)) & " (" & Trim(rsReporte.Fields!ndpto) & ")")
                            Call Feuille.getcellbyposition(6, 5).setFormula("5 UPS")
                        Else
                            oWorkSheet.Cells(6, 5).Value = "'" & rsReporte.Fields!Codigo & " " & lcBuscaParametro.SeleccionaFilaParametro(205)    '2017
                            oWorkSheet.Cells(5, 17).Value = "5 UPS"
                            oWorkSheet.Cells(6, 7).Value = "'" & Trim(str(rsReporte.Fields!IdDepartamento)) & " (" & Trim(rsReporte.Fields!ndpto) & ")"
                        End If
                        
                        lcProcedencia = Trim(str(rsReporte.Fields!IdDepartamento))
                        lcPaciente = Trim(str(rsReporte.Fields!IdProvincia))
                        
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(7, 5).setFormula("'" & Mid(lcPaciente, Len(lcProcedencia) + 1, 10))
                        Else
                            oWorkSheet.Cells(6, 8).Value = "'" & Mid(lcPaciente, Len(lcProcedencia) + 1, 10)
                        End If
                        lcProcedencia = Trim(str(rsReporte.Fields!IdDistrito))
                        
                        If lbEsOpenOffice = True Then
                            Call Feuille.getcellbyposition(9, 5).setFormula("'" & Mid(lcProcedencia, Len(lcPaciente) + 1, 10))
                            Call Feuille.getcellbyposition(11, 5).setFormula("'" & rsReporte.Fields!Codigo)
                        Else
                            oWorkSheet.Cells(6, 10).Value = "'" & Mid(lcProcedencia, Len(lcPaciente) + 1, 10)
                            oWorkSheet.Cells(6, 12).Value = "'" & rsReporte.Fields!Codigo
                            
                        End If
                    End If
                    rsReporte.Close
                    
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(16, 5).setFormula(ml_TextoDelFiltro1)
                        Call Feuille.getcellbyposition(18, 5).setFormula(UCase(ml_TextoDelFiltro))
                        Call Feuille.getcellbyposition(21, 5).setFormula(lcCodigoHIsResponsable) 'Actualizado 14102014
                        Call Feuille.getcellbyposition(2, 6).setFormula(Year(CDate(mda_FechaInicio)))
                        Call Feuille.getcellbyposition(3, 6).setFormula(Month(CDate(mda_FechaInicio)))
                    Else
                        oWorkSheet.Cells(6, 10).Value = ml_TextoDelFiltro1         '2017
                        oWorkSheet.Cells(6, 20).Value = UCase(ml_TextoDelFiltro)   '2017
                        oWorkSheet.Cells(6, 19).Value = lcCodigoHIsResponsable 'Actualizado 14102014   '2017
                        oWorkSheet.Cells(6, 3).Value = Year(CDate(mda_FechaInicio))    '2017
                        oWorkSheet.Cells(6, 4).Value = Month(CDate(mda_FechaInicio))   '2017
                    End If
                    iFila = 13
                    lbCargaUnaVezPorHoja = False
                End If 'Fin de carga una vez por hoja his
                

                'CIE10
                lcFilter = "idPuntoCArga=1"
                If SighEntidades.Parametro561 = "S" Then
                   lcFilter = lcFilter + " or idPuntoCarga=99"
                End If
                If lnoRsImpHIStodos > 0 Then
                    Set rsDx = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(mrs_Tmp.Fields!idAtencion)
                    lnTotalDxCpt = rsDx.RecordCount
                     'CPT
                     If lbAgregaOrdenesMedicas = True Or ml_AgregaCPT = True Then
                        Set rsCpt = mo_AdminAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(mrs_Tmp.Fields!idCuentaAtencion, 1)
                        If lbAgregaOrdenesMedicas = False Then
                           rsCpt.Filter = lcFilter
                        End If
                        lnTotalDxCpt = lnTotalDxCpt + rsCpt.RecordCount
                    End If
                 Else
                    Set rsDx1 = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(mrs_Tmp.Fields!idAtencion)
                    'CPT
                    If lbAgregaOrdenesMedicas = True Or ml_AgregaCPT = True Then
                        Set rsCpt = mo_AdminAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(mrs_Tmp!idCuentaAtencion, 1)
                        If lbAgregaOrdenesMedicas = False Then
                           rsCpt.Filter = lcFilter
                        End If
                    End If
                    mo_AdminAdmision.CreaTemporalYcargaDxYcptXgrupo rsDx, rsDx1, rsCpt, mrs_Tmp!idAtencion, True, lcAD040
                    lnTotalDxCpt = rsDx.RecordCount
               End If
               '
               If lnTotalDxCpt > 0 Then
                   '
                   mo_ReglasLaboratorio.DevuelveHemoglobinaDeLaAtencion ldFechaHemoglobina, lcHemoglobina, mrs_Tmp!idPaciente, _
                                        mrs_Tmp!idAtencion, mrs_Tmp!FechaIngreso, oConexion
                   
'                   ldFechaHemoglobina = 0
'                   lcHemoglobina = ""
'                   Set oRsTmp987 = mo_ReglasLaboratorio.LabResultadoPorItemsSeleccionarPorPaciente(mrs_Tmp!idPaciente, 1, oConexion, _
'                                                                                                   sighentidades.DevuelveNroOrdenParaHemoglobina(mrs_Tmp!idAtencion))
'                   If oRsTmp987.RecordCount > 0 Then
'                      oRsTmp987.MoveFirst
'                      oRsTmp987.Find "fecha<='" & mrs_Tmp!fechaIngreso & "'"
'                      If Not oRsTmp987.EOF Then
'                            ldFechaHemoglobina = oRsTmp987!fecha
'                            If Not IsNull(oRsTmp987!valorTexto) Then
'                               lcHemoglobina = oRsTmp987!valorTexto
'                            ElseIf Not IsNull(oRsTmp987!valorNumero) Then
'                               lcHemoglobina = Trim(str(oRsTmp987!valorNumero))
'                            End If
'                      End If
'                   End If
'                   oRsTmp987.Close
                   '
                   Set oRsTmp987 = mo_ReglasCaja.BuscarTriaje(mrs_Tmp!idAtencion)
                   If oRsTmp987.RecordCount > 0 Then
                        If lbEsOpenOffice = True Then
                        Else
                           If Not IsNull(oRsTmp987!triajePeso) Then
                              oWorkSheet.Cells(iFila, 13).Value = oRsTmp987!triajePeso
                           End If
                           If Not IsNull(oRsTmp987!triajeTalla) Then
                              oWorkSheet.Cells(iFila + 2, 13).Value = oRsTmp987!triajeTalla
                           End If
                           If lcHemoglobina <> "" Then
                               oWorkSheet.Cells(iFila + 4, 13).Value = lcHemoglobina
                           End If
                           
                           If Not IsNull(oRsTmp987!TriajePerimCefalico) Then
                              oWorkSheet.Cells(iFila, 11).Value = oRsTmp987!TriajePerimCefalico
                           End If
                           If Not IsNull(oRsTmp987!TriajePerimAbdominal) Then
                              oWorkSheet.Cells(iFila + 3, 11).Value = oRsTmp987!TriajePerimAbdominal
                           End If
                           
                        End If
                   End If
                   oRsTmp987.Close

                   '
                   lcProcedencia = ""
                   If Not IsNull(mrs_Tmp.Fields!distrito) Then
                      lcProcedencia = lcProcedencia & Trim(mrs_Tmp.Fields!distrito)
                   End If
                   If Not IsNull(mrs_Tmp.Fields!provincia) Then
                      lcProcedencia = lcProcedencia & " (" & Trim(mrs_Tmp.Fields!provincia) & ")"
                   End If
                   If Not IsNull(mrs_Tmp.Fields!dpto) Then
                      lcProcedencia = lcProcedencia & " (" & Trim(mrs_Tmp.Fields!dpto) & ")"
                   End If
                   If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(2, iFila - 1).setFormula(Day(mrs_Tmp.Fields("fechaIngreso").Value))
                   Else
                        oWorkSheet.Cells(iFila - 1, 5).Value = mrs_Tmp!Paciente1
                        oWorkSheet.Cells(iFila - 1, 18).Value = "'" & Trim(str(Day(mrs_Tmp!FechaNacimiento))) & "/" & Trim(str(Month(mrs_Tmp!FechaNacimiento))) & "/" & Trim(str(Year(mrs_Tmp!FechaNacimiento)))
                        If lcHemoglobina <> "" Then
                           oWorkSheet.Cells(iFila - 1, 14).Value = "'" & Trim(str(Day(ldFechaHemoglobina))) & "/" & Trim(str(Month(ldFechaHemoglobina))) & "/" & Trim(str(Year(ldFechaHemoglobina)))
                        End If
                        oWorkSheet.Cells(iFila, 3).Value = Day(mrs_Tmp.Fields("fechaIngreso").Value)
                   End If
                   If lbEsOpenOffice = True Then
                   Else
                        If mrs_Tmp.Fields("IdDocIdentidad").Value = 1 And Len(mrs_Tmp.Fields("NroDocumento").Value) > 0 Then
                           oWorkSheet.Cells(iFila, 4).Value = "'" & mrs_Tmp.Fields("NroDocumento").Value
                        End If
                        oWorkSheet.Cells(iFila + 2, 4).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(mrs_Tmp!NroHistoriaClinica)), True)
                        If Not IsNull(mrs_Tmp.Fields("fichaFamiliar").Value) And lbEsCS = True Then
                           oWorkSheet.Cells(iFila + 4, 4).Value = "'" & mrs_Tmp.Fields("fichaFamiliar").Value
                        End If
                   End If
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(4, iFila - 1).setFormula(mrs_Tmp.Fields("CodigoHIS").Value)
                        Call Feuille.getcellbyposition(5, iFila - 1).setFormula(mrs_Tmp.Fields("IdEtnia").Value)
                        Call Feuille.getcellbyposition(6, iFila - 1).setFormula(lcProcedencia)
                        Call Feuille.getcellbyposition(7, iFila - 1).setFormula(mrs_Tmp.Fields("edad").Value & " (" & Trim(mrs_Tmp.Fields!tEdad) & ")")
                        Call Feuille.getcellbyposition(8, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoSexo").Value = 1, "X", ""))
                        Call Feuille.getcellbyposition(9, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoSexo").Value = 1, "", "X"))
                    Else
                        oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("CodigoHIS").Value   'financ 2017
                        oWorkSheet.Cells(iFila + 3, 5).Value = mrs_Tmp.Fields("IdEtnia").Value 'etnia 2017
                        
                        oWorkSheet.Cells(iFila, 6).Value = lcProcedencia
                        If Not IsNull(mrs_Tmp!CentroPoblado) Then
                           oWorkSheet.Cells(iFila + 3, 6).Value = mrs_Tmp!CentroPoblado
                        End If
                        
                        oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields("edad").Value
                        If UCase(Left(mrs_Tmp.Fields!tEdad, 1)) = "A" Then
                           oWorkSheet.Cells(iFila, 8).Value = "A"
                        ElseIf UCase(Left(mrs_Tmp.Fields!tEdad, 1)) = "M" Then
                           oWorkSheet.Cells(iFila + 2, 8).Value = "M"
                        Else
                           oWorkSheet.Cells(iFila + 4, 8).Value = "D"
                        End If
                        If mrs_Tmp.Fields("idTipoSexo").Value = 1 Then
                           oWorkSheet.Cells(iFila, 9).Value = "M"
                        Else
                           oWorkSheet.Cells(iFila + 3, 9).Value = "F"
                        End If
                    End If

                    
                   If Not IsNull(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value) Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(13, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 1, "X", ""))    'Nuevo
                        Call Feuille.getcellbyposition(14, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 3, "X", ""))   'Continuador
                        Call Feuille.getcellbyposition(15, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 2, "X", ""))   'Reingreso)
                    Else
                        If mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 1 Then
                           oWorkSheet.Cells(iFila, 15).Value = "N"
                        ElseIf mrs_Tmp.Fields("idTipoCondicionAlServicio").Value = 3 Then
                           oWorkSheet.Cells(iFila + 2, 15).Value = "C"
                        Else
                           oWorkSheet.Cells(iFila + 4, 15).Value = "R"
                        End If
                    End If
                   End If
                   If Not IsNull(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value) Then
                    If lbEsOpenOffice = True Then
                        Call Feuille.getcellbyposition(10, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 1, "X", ""))    'Nuevo
                        Call Feuille.getcellbyposition(11, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 3, "X", ""))   'Continuador
                        Call Feuille.getcellbyposition(12, iFila - 1).setFormula(IIf(mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 2, "X", ""))     'Reingreso)
                    Else
                      If mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 1 Then
                         oWorkSheet.Cells(iFila, 14).Value = "N"
                      ElseIf mrs_Tmp.Fields("idTipoCondicionAlEstab").Value = 3 Then
                         oWorkSheet.Cells(iFila + 2, 14).Value = "C"
                      Else
                         oWorkSheet.Cells(iFila + 4, 14).Value = "R"
                      End If
                    End If
                   End If
                   lnLineas = 1
                   lnSeccionDx = 0
                   
                   If rsDx.RecordCount > 0 Then
                        rsDx.MoveFirst
                        Do While Not rsDx.EOF
    '                       If lnLineas <= 6 Then
    '                        lnSeccionDx = IIf(Round(lnLineas / 6) < lnLineas / 6, Round(lnLineas / 6) + 1, Round(lnLineas / 6))
                            lnLineas = IIf(lnLineas Mod 7 = 0, lnLineas + 1, lnLineas)
                            If lbEsOpenOffice = True Then
                                Call Feuille.getcellbyposition(16, iFila + lnLineas - 2).setFormula(rsDx.Fields("dDiagnostico").Value)
                                Call Feuille.getcellbyposition(19, iFila + lnLineas - 2).setFormula(IIf(rsDx.Fields("idSubClasificacionDx").Value = 101, "X", ""))  'Presuntivo)
                                Call Feuille.getcellbyposition(20, iFila + lnLineas - 2).setFormula(IIf(rsDx.Fields("idSubClasificacionDx").Value = 102, "X", ""))
                                Call Feuille.getcellbyposition(21, iFila + lnLineas - 2).setFormula(IIf(rsDx.Fields("idSubClasificacionDx").Value = 103, "X", ""))
                                Call Feuille.getcellbyposition(22, iFila + lnLineas - 2).setFormula(Trim(IIf(IsNull(rsDx.Fields("labConfHIS").Value), "", rsDx.Fields("labConfHIS").Value)))
                                Call Feuille.getcellbyposition(23, iFila + lnLineas - 2).setFormula(Trim(rsDx.Fields("codigoCie10").Value))
                            Else
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 17).Value = IIf(SighEntidades.Lx_LabVacio = rsDx.Fields("dDiagnostico").Value, "", rsDx.Fields("dDiagnostico").Value)
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 20).Value = IIf(rsDx.Fields("idSubClasificacionDx").Value = 101, "X", "") 'Presuntivo
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 21).Value = IIf(rsDx.Fields("idSubClasificacionDx").Value = 102, "X", "") 'Definitivo
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 22).Value = IIf(rsDx.Fields("idSubClasificacionDx").Value = 103, "X", "") 'Repetido
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 23).Value = Trim(rsDx.Fields("labConfHIS").Value)
                                 oWorkSheet.Cells(iFila + lnLineas - 1, 24).Value = Trim(IIf(SighEntidades.Lx_LabVacio = rsDx.Fields("codigoCie10").Value, "", rsDx.Fields("codigoCie10").Value))
                            End If
                            
    '                        End If
                           lnLineas = lnLineas + 1
                           rsDx.MoveNext
                        Loop
                   End If
                   rsDx.Close
                   'CPT
                   If ml_AgregaCPT = True Then
                        'Set rsCpt = mo_AdminAdmision.BuscaAtencionesCptCEparaFormatoHIS(mrs_Tmp.Fields!idCuentaAtencion)
                        If rsCpt.RecordCount > 0 Then
                             rsCpt.MoveFirst
                             Do While Not rsCpt.EOF
    '                           If lnLineas <= 6 Then
                                lnLineas = IIf(lnLineas Mod 7 = 0, lnLineas + 1, lnLineas)
                                If lbEsOpenOffice = True Then
                                    'Actualizado 22092014
                                    Call Feuille.getcellbyposition(16, iFila + lnLineas - 2).setFormula(rsCpt.Fields("Nombre").Value)
                                    Call Feuille.getcellbyposition(20, iFila + lnLineas - 2).setFormula("X")
                                    Call Feuille.getcellbyposition(23, iFila + lnLineas - 2).setFormula(Trim(rsCpt.Fields("codigo").Value))
                                    Call Feuille.getcellbyposition(22, iFila + lnLineas - 2).setFormula(Trim(rsCpt.Fields("labConfHIS").Value))
                                    '''''''''''''''''''''
                                Else
                                    oWorkSheet.Cells(iFila + lnLineas - 1, 17).Value = rsCpt.Fields("Nombre").Value
                                    oWorkSheet.Cells(iFila + lnLineas - 1, 21).Value = "X"
                                    oWorkSheet.Cells(iFila + lnLineas - 1, 24).Value = Trim(rsCpt.Fields("codigo").Value)
                                    oWorkSheet.Cells(iFila + lnLineas - 1, 23).Value = Trim(rsCpt.Fields("labConfHIS").Value)
                                End If
    '                           End If
                                lnLineas = lnLineas + 1
                                rsCpt.MoveNext
                             Loop
                        End If
                        rsCpt.Close
                   End If
                   lnLineas = lnLineas - 1
                   '32   '41-34
                    iFila = iFila + lnLineas + (7 - lnLineas Mod 7)
                    If Round((iFila - 12) / 7) < (iFila - 12) / 7 Then
                         If Round((iFila - 12) / 7) + 1 = 25 Then
                             lbCargaUnaVezPorHoja = True
                         End If
                    Else
                         If Round((iFila - 12) / 7) = 25 Then
                             lbCargaUnaVezPorHoja = True
                         End If
                    End If
               End If
               mrs_Tmp.MoveNext
'               iFila = iFila + 7
            Loop

            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 24
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$10"
                    If oWorkSheet.PageSetup.PrintArea <> "" Then
                       oWorkSheet.PageSetup.PrintArea = SighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila + 3)
                    End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
            End If
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
                Set oWorkSheet = Nothing
                Set oExcel = Nothing
                Set oWorkBookPlantilla = Nothing
                Set oWorkBook = Nothing
            End If
        
        End If
        oConexionExterna.Close
        Set oConexionExterna = Nothing
        Set oRsTmp987 = Nothing
        Set mo_ReglasCaja = Nothing
        Set mo_ReglasLaboratorio = Nothing
        Set oDOPaciente = Nothing
        Set oDOAtencion = Nothing
        Set rsReporte = Nothing
        Set rsDatosMedico = Nothing
        Set oRsImpHIStodos = Nothing

Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    
    Exit Sub
    Resume
End Sub



Sub RegistroDiarioAtencionNino(lnHwnd As Long, ldFecha As Date)
Dim iFila As Long
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim lnLineas As Integer: Dim lnIdAtencion  As Long
Dim lcProcedencia As String, lcHcFicha As String, lbEsCS As Boolean
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String
Dim lnSeccionDx As Integer 'Frank 15092014
Dim oDOPaciente As New DOPaciente
Dim oDOAtencion As New DOAtencion
Dim rsReporte As New Recordset
Dim mo_ReglasLaboratorio As New ReglasLaboratorio
Dim rsDx As New Recordset
Dim rsCpt As New Recordset
Dim rsDx1 As New Recordset
Dim oRsTmp987 As New Recordset
Dim mo_ReglasCaja As New ReglasCaja
Dim rsDatosMedico As Recordset 'Actualizado 14102014
Dim oRsImpHIStodos As New Recordset
Dim oRsVacunas As New Recordset
Dim oConexionExterna As New Connection
Dim oConexion As New Connection
Dim lcCodigoHIsResponsable As String 'Actualizado 14102014
Dim lnoRsImpHIStodos As Long
Dim lcAD040 As String
Dim ldFechaHemoglobina As Date, lcHemoglobina As String
Dim lnHistoria As Long, lcDNI As String, lcTipoSeguro As String, lcEdad As String, ldFechaNacimiento As Date, lcDx As String
Dim lnIdTipoSexo As Long, lcCred As String, lcHB As String, lcMMM As String, lcVacunas As String
'
lcAD040 = lcBuscaParametro.SeleccionaFilaParametro(354)
'
oConexionExterna.CursorLocation = adUseClient
oConexionExterna.CommandTimeout = 150
oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'
oConexion.CursorLocation = adUseClient
oConexion.CommandTimeout = 150
oConexion.Open SighEntidades.CadenaConexion

'lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorErrorExcel

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
    mda_FechaInicio = ldFecha
    mda_FechaFin = ldFecha
    'Filtra los Datos
    Set mrs_Tmp = mo_AdminAdmision.BuscaAtencionesCEparaFormatoHIS(ml_IdServicioCE, ml_IdResponsable, mda_FechaInicio, _
                                                mda_FechaFin, "")          'cred
    
    Set oRsVacunas = mo_AdminAdmision.BuscaAtencionesCEparaFormatoHIS(ml_IdServicioCE, ml_IdResponsable, mda_FechaInicio, _
                                                mda_FechaFin, "codigoServicioHIS='301204'")          'vacuna
    mrs_Tmp.Filter = "fechaEgreso<>null"
    If mrs_Tmp.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
            If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\CEHis.ods"

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
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                'Crea nueva hoja
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\CEatencionNino.xls")
                oWorkBookPlantilla.Worksheets("CEatencionNino").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
            iFila = 2
            oWorkSheet.Cells(iFila, 2).Value = "Fecha: " & mda_FechaInicio & "      Responsable de la atención: " & ml_TextoDelFiltro & "      Consultorio: " & ml_TextoDelFiltro1
            iFila = iFila + 4
            lnLineas = 1
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               lnHistoria = mrs_Tmp!NroHistoriaClinica
               lcDNI = IIf(IsNull(mrs_Tmp!nroDocumento), "", mrs_Tmp!nroDocumento)
               lcTipoSeguro = mrs_Tmp!ffinanciamiento
               lcPaciente = mrs_Tmp!Paciente1
               lcEdad = mrs_Tmp!Edad & " " & mrs_Tmp!tEdad
               ldFechaNacimiento = mrs_Tmp!FechaNacimiento
               lnIdTipoSexo = mrs_Tmp!idTipoSexo
               lcDx = ""
               lcCred = "": lcHB = "": lcMMM = ""
               lnIdPaciente = mrs_Tmp!idPaciente
               '
               lcVacunas = "/"
               oRsVacunas.Filter = "idPaciente=" & lnIdPaciente
               If oRsVacunas.RecordCount > 0 Then
                    oRsVacunas.MoveFirst
                    Do While Not oRsVacunas.EOF
                    Set rsCpt = mo_AdminAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(oRsVacunas!idCuentaAtencion, _
                                                 sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
                    rsCpt.Filter = "idPuntoCArga=1"
                    If rsCpt.RecordCount > 0 Then
                        rsCpt.MoveFirst
                        Do While Not rsCpt.EOF
                           lcVacunas = lcVacunas & Trim(rsCpt!Codigo) & "/"
                           rsCpt.MoveNext
                        Loop
                    End If
                    oRsVacunas.MoveNext
                    Loop
               End If
               '
               ldFechaHemoglobina = 0
               lcHemoglobina = ""
               Set oRsTmp987 = mo_ReglasLaboratorio.LabResultadoPorItemsSeleccionarPorPaciente(mrs_Tmp!idPaciente, 1, oConexion, _
                                                                                               SighEntidades.DevuelveNroOrdenParaHemoglobina(mrs_Tmp!idAtencion))
                                                                                               'Val("999" & Trim(Str(mrs_Tmp!idAtencion))))
               If oRsTmp987.RecordCount > 0 Then
                   oRsTmp987.MoveFirst
                   oRsTmp987.Find "fecha<='" & mrs_Tmp!FechaIngreso & "'"
                   If Not oRsTmp987.EOF Then
                         ldFechaHemoglobina = oRsTmp987!fecha
                         If Not IsNull(oRsTmp987!valorTexto) Then
                            lcHemoglobina = oRsTmp987!valorTexto
                         ElseIf Not IsNull(oRsTmp987!valorNumero) Then
                            lcHemoglobina = Trim(str(oRsTmp987!valorNumero))
                         End If
                   End If
               End If
               oRsTmp987.Close
               '
               Do While Not mrs_Tmp.EOF And lnIdPaciente = mrs_Tmp!idPaciente
                    lnIdAtencion = mrs_Tmp!idAtencion
                    Set rsDx1 = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(lnIdAtencion)
                    mo_AdminAdmision.CreaTemporalYcargaDxYcptXgrupo rsDx, rsDx1, rsCpt, lnIdAtencion, True, lcAD040
                    If rsDx.RecordCount > 0 Then
                         rsDx.MoveFirst
                         Do While Not rsDx.EOF
                             If Trim(rsDx.Fields("codigoCie10").Value) <> "" Then
                                If InStr(rsDx!CodigoCIE10, ".") > 0 Then
                                   lcDx = lcDx & rsDx.Fields("codigoCie10").Value & "/"
                                End If
                             End If
                             Select Case UCase(Trim(rsDx!CodigoCIE10))
                             Case "Z00.1"
                                If Not IsNull(rsDx!labConfHIS) Then
                                lcCred = rsDx!labConfHIS
                                End If
                             Case "D50.9"
                                If Not IsNull(rsDx!labConfHIS) Then
                                lcHB = rsDx!labConfHIS
                                End If
                             Case "Z29.8"
                                If Not IsNull(rsDx!labConfHIS) Then
                                lcMMM = rsDx!labConfHIS
                                End If
                                
                             End Select
                             rsDx.MoveNext
                         Loop
                    End If
                    '
                    Do While Not mrs_Tmp.EOF And lnIdPaciente = mrs_Tmp!idPaciente And lnIdAtencion = mrs_Tmp!idAtencion
                        mrs_Tmp.MoveNext
                        If mrs_Tmp.EOF Then
                           Exit Do
                        End If
                    Loop
                    If mrs_Tmp.EOF Then
                       Exit Do
                    End If
               Loop
               If lcVacunas = "/" Then
                  lcVacunas = ""
               End If
               If lbEsOpenOffice = True Then
               Else
                    oWorkSheet.Cells(iFila, 2).Value = "'" & lnLineas
                    oWorkSheet.Cells(iFila, 3).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(lnHistoria)), True)
                    oWorkSheet.Cells(iFila, 4).Value = lcDNI
                    oWorkSheet.Cells(iFila, 5).Value = lcTipoSeguro
                    oWorkSheet.Cells(iFila, 6).Value = lcPaciente
                    oWorkSheet.Cells(iFila, IIf(lnIdTipoSexo = 2, 7, 8)).Value = "X"
                    oWorkSheet.Cells(iFila, 9).Value = Format(ldFechaNacimiento, SighEntidades.DevuelveFechaSoloFormato_DMY)
                    oWorkSheet.Cells(iFila, 10).Value = lcVacunas
                    oWorkSheet.Cells(iFila, 11).Value = lcCred
                    oWorkSheet.Cells(iFila, 12).Value = lcMMM
                    oWorkSheet.Cells(iFila, 14).Value = lcHemoglobina
                    
                    oWorkSheet.Cells(iFila, 15).Value = lcDx
                End If
                lnLineas = lnLineas + 1
                iFila = iFila + 1
            Loop
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 24
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$10"
                    If oWorkSheet.PageSetup.PrintArea <> "" Then
                       oWorkSheet.PageSetup.PrintArea = SighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila + 3)
                    End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
            End If
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
                Set oWorkSheet = Nothing
                Set oExcel = Nothing
                Set oWorkBookPlantilla = Nothing
                Set oWorkBook = Nothing
            End If
        
        End If
        oConexionExterna.Close
        Set oConexionExterna = Nothing
        Set oRsTmp987 = Nothing
        Set mo_ReglasCaja = Nothing
        Set mo_ReglasLaboratorio = Nothing
        Set oDOPaciente = Nothing
        Set oDOAtencion = Nothing
        Set rsReporte = Nothing
        Set rsDatosMedico = Nothing
        Set oRsImpHIStodos = Nothing

Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    
    Exit Sub
    Resume
End Sub


Sub RegistroDiarioVacunacionMujeres(lnHwnd As Long, ldFecha As Date)
Dim iFila As Long
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim lnLineas As Integer: Dim lnIdAtencion  As Long
Dim lcProcedencia As String, lcHcFicha As String, lbEsCS As Boolean
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String
Dim lnSeccionDx As Integer 'Frank 15092014
Dim oDOPaciente As New DOPaciente
Dim oDOAtencion As New DOAtencion
Dim rsReporte As New Recordset
Dim mo_ReglasLaboratorio As New ReglasLaboratorio
Dim rsDx As New Recordset
Dim rsCpt As New Recordset
Dim rsDx1 As New Recordset
Dim oRsTmp987 As New Recordset
Dim mo_ReglasCaja As New ReglasCaja
Dim rsDatosMedico As Recordset 'Actualizado 14102014
Dim oRsImpHIStodos As New Recordset
Dim oConexionExterna As New Connection
Dim oConexion As New Connection
Dim lcCodigoHIsResponsable As String 'Actualizado 14102014
Dim lnoRsImpHIStodos As Long
Dim lcAD040 As String
Dim ldFechaHemoglobina As Date, lcHemoglobina As String
Dim lnHistoria As Long, lcDNI As String, lcTipoSeguro As String, lcEdad As String, ldFechaNacimiento As Date, lcDx As String
Dim lnIdTipoSexo As Long
Dim lcDireccion As String, lcDpto As String, lcProvincia As String, lcDistrito As String, lnEdadEnMeses As Long
Dim lnTot1 As Integer, lnTot2 As Integer, lnTot3 As Integer, lnTot4 As Integer, lnTot5 As Integer, lnTot6 As Integer
Dim lnTot11 As Integer, lnTot12 As Integer, lnTot13 As Integer, lnTot14 As Integer, lnTot15 As Integer, lnTot16 As Integer
Dim lnTot21 As Integer, lnTot22 As Integer, lnTot23 As Integer, lnTot24 As Integer, lnTot25 As Integer, lnTot26 As Integer
Dim lnNinoAdolescenteJoven As Integer, lcLab As String
Dim lcEdadAtencion As String, oEdad As Edad, lnHoras As Integer
Dim lcBCG As String, lcHvB As String, lcPenta As String, lcIPV As String, lcAPO As String
Dim lcSPR As String, lcHepaB As String
'
lcAD040 = lcBuscaParametro.SeleccionaFilaParametro(354)
'
oConexionExterna.CursorLocation = adUseClient
oConexionExterna.CommandTimeout = 150
oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'
oConexion.CursorLocation = adUseClient
oConexion.CommandTimeout = 150
oConexion.Open SighEntidades.CadenaConexion

'lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorErrorExcel

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
    mda_FechaInicio = ldFecha
    mda_FechaFin = ldFecha
    'Filtra los Datos
    Set mrs_Tmp = mo_AdminAdmision.BuscaAtencionesCEparaFormatoHIS(ml_IdServicioCE, ml_IdResponsable, mda_FechaInicio, _
                                                mda_FechaFin, "")          'vacuna
    mrs_Tmp.Filter = "fechaEgreso<>null"
    If mrs_Tmp.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
            If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\CEHis.ods"

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
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                'Crea nueva hoja
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ceVacunacion.xls")
                oWorkBookPlantilla.Worksheets("ceVacunacion").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
            lnNumTotal = 9
            oWorkSheet.Cells(2, 2).Value = "Responsable de la atención: " & ml_TextoDelFiltro & "      Consultorio: " & ml_TextoDelFiltro1
            iFila = 7
            oWorkSheet.Cells(iFila, 3).Value = "E- FECHA DE VACUNACION:    " & mda_FechaInicio
            iFila = iFila + 6
            lnLineas = 1
            lnTot1 = 0: lnTot2 = 0: lnTot3 = 0: lnTot4 = 0: lnTot5 = 0: lnTot6 = 0
            lnTot11 = 0: lnTot12 = 0: lnTot13 = 0: lnTot14 = 0: lnTot15 = 0: lnTot16 = 0
            lnTot21 = 0: lnTot22 = 0: lnTot23 = 0: lnTot24 = 0: lnTot25 = 0: lnTot26 = 0
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               lnIdAtencion = mrs_Tmp!idAtencion
               lnHistoria = mrs_Tmp!NroHistoriaClinica
               lcDNI = IIf(IsNull(mrs_Tmp!nroDocumento), "", mrs_Tmp!nroDocumento)
               lcTipoSeguro = mrs_Tmp!ffinanciamiento
               lcPaciente = mrs_Tmp!Paciente1
               lcEdad = mrs_Tmp!Edad & " " & mrs_Tmp!tEdad
               ldFechaNacimiento = mrs_Tmp!FechaNacimiento
               lnIdTipoSexo = mrs_Tmp!idTipoSexo
               lcDireccion = IIf(IsNull(mrs_Tmp!DireccionDomicilio), "", mrs_Tmp!DireccionDomicilio)
               lcDpto = IIf(IsNull(mrs_Tmp!dpto), "", mrs_Tmp!dpto)
               lcProvincia = IIf(IsNull(mrs_Tmp!provincia), "", mrs_Tmp!provincia)
               lcDistrito = IIf(IsNull(mrs_Tmp!distrito), "", mrs_Tmp!distrito)
               lcDx = ""
               lcEdadAtencion = getDescripcionEdad(mrs_Tmp!Edad, mrs_Tmp!idTipoEdad, _
                                                   mrs_Tmp!FechaNacimiento, CDate(Format(mrs_Tmp!FechaIngreso & " " & mrs_Tmp!HoraIngreso, SighEntidades.DevuelveFechaSoloFormato_DMY_HMS)))
               oEdad = calcularEdadDisgregada(mrs_Tmp!FechaNacimiento, CDate(Format(mrs_Tmp!FechaIngreso & " " & mrs_Tmp!HoraIngreso, SighEntidades.DevuelveFechaSoloFormato_DMY_HMS)))
               If oEdad.EdadAnio = 0 And oEdad.EdadMes = 0 Then
                  lnHoras = DateDiff("h", mrs_Tmp!FechaNacimiento, CDate(Format(mrs_Tmp!FechaIngreso & " " & mrs_Tmp!HoraIngreso, SighEntidades.DevuelveFechaSoloFormato_DMY_HMS)))
               Else
                  lnHoras = 0
               End If
               '
               lcBCG = "": lcHvB = "": lcPenta = "": lcIPV = "": lcAPO = "": lcSPR = "": lcHepaB = ""
               Set rsCpt = mo_AdminAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(mrs_Tmp.Fields!idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
               If rsCpt.RecordCount > 0 Then
                  rsCpt.MoveFirst
                  Do While Not rsCpt.EOF
                     Select Case rsCpt!Codigo
                     Case "90585"
                         lcBCG = "X"
                     Case "90744"
                         lcHvB = "X"
                     Case "90723"
                         lcPenta = IIf(IsNull(rsCpt!labConfHIS), "", rsCpt!labConfHIS)
                     Case "90713"
                         lcIPV = IIf(IsNull(rsCpt!labConfHIS), "", rsCpt!labConfHIS)
                     Case "90712"
                         lcAPO = IIf(IsNull(rsCpt!labConfHIS), "", rsCpt!labConfHIS)
                     Case "90707"
                         lcSPR = IIf(IsNull(rsCpt!labConfHIS), "", rsCpt!labConfHIS)
                   '  Case "90744"
                    '     lcHepaB = IIf(IsNull(rsCpt!labConfHIS), "", rsCpt!labConfHIS)
                     End Select
                     rsCpt.MoveNext
                  Loop
               End If
               rsCpt.Close
               '
               lcLab = ""
               lnIdPaciente = mrs_Tmp!idPaciente
'               Do While Not mrs_Tmp.EOF And lnIdPaciente = mrs_Tmp!idPaciente
'                    lnIdAtencion = mrs_Tmp!idAtencion
'                    Set rsDx1 = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(lnIdAtencion)
'                    mo_AdminAdmision.CreaTemporalYcargaDxYcptXgrupo rsDx, rsDx1, rsCpt, lnIdAtencion, True, lcAD040
'                    If rsDx.RecordCount > 0 Then
'                         rsDx.MoveFirst
'                         Do While Not rsDx.EOF
'                             If Trim(rsDx.Fields("codigoCie10").Value) <> "" Then
'                                lcDx = lcDx & rsDx.Fields("codigoCie10").Value & "/"
'                             End If
'                             If UCase(Trim(rsDx!CodigoCIE10)) = "Z27.81" And Not IsNull(rsDx!labConfHIS) Then
'                                lcLab = rsDx!labConfHIS
'                                If lcLab = "G" Then
'                                   lcLab = "3"
'                                Else
'                                   lcLab = "0"
'                                End If
'                                Exit Do
'                             End If
'                             rsDx.MoveNext
'                         Loop
'                    End If
                    '
'                    Do While Not mrs_Tmp.EOF And lnIdPaciente = mrs_Tmp!idPaciente And lnIdAtencion = mrs_Tmp!idAtencion
'                        mrs_Tmp.MoveNext
'                        If mrs_Tmp.EOF Then
'                           Exit Do
'                        End If
'                    Loop
'                    If mrs_Tmp.EOF Then
'                       Exit Do
'                    End If
'               Loop
               '
                If lbEsOpenOffice = True Then
                Else
                     oWorkSheet.Cells(iFila, 2).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(lnHistoria)), True)
                     oWorkSheet.Cells(iFila, 3).Value = lcDNI
                     oWorkSheet.Cells(iFila, 4).Value = lcPaciente
                     oWorkSheet.Cells(iFila, 5).Value = lcDireccion
                     oWorkSheet.Cells(iFila, 6).Value = Format(ldFechaNacimiento, SighEntidades.DevuelveFechaSoloFormato_DMY)
                     oWorkSheet.Cells(iFila, 7).Value = lcDistrito
                     oWorkSheet.Cells(iFila, 8).Value = lcEdadAtencion
                     If lnHoras > 0 And lnHoras < 24 Then
                        '************horas de nacido
                        oWorkSheet.Cells(iFila, 9).Value = lcBCG
                        If lnHoras < 12 Then
                           oWorkSheet.Cells(iFila, 12).Value = lcHvB
                        ElseIf lnHoras < 24 Then
                           oWorkSheet.Cells(iFila, 13).Value = lcHvB
                        End If
                        If lcIPV <> "" Then
                           oWorkSheet.Cells(iFila, IIf(lcIPV = "1", 14, 15)).Value = "X"
                        End If
                        If lcAPO <> "" Then
                           oWorkSheet.Cells(iFila, 16).Value = "X"
                        End If
                        If lcPenta <> "" Then
                           oWorkSheet.Cells(iFila, IIf(lcPenta = "1", 19, IIf(lcPenta = "2", 20, 21))).Value = "X"
                        End If
                     Else
                        If oEdad.EdadAnio > 0 Then
                            '************Edad en años
                            If lcSPR <> "" Then
                               oWorkSheet.Cells(iFila, 44).Value = "X"
                            End If
                            If lcIPV <> "" Then
                               Select Case oEdad.EdadAnio
                               Case 1
                                    oWorkSheet.Cells(iFila, IIf(lcIPV = "1", 47, 48)).Value = "X"
                               Case 2
                                    oWorkSheet.Cells(iFila, IIf(lcIPV = "1", 65, 66)).Value = "X"
                               Case 3
                                    oWorkSheet.Cells(iFila, IIf(lcIPV = "1", 83, 84)).Value = "X"
                               Case 4
                                    oWorkSheet.Cells(iFila, IIf(lcIPV = "1", 102, 102)).Value = "X"
                               End Select
                            End If
                            If lcAPO <> "" Then
                               Select Case oEdad.EdadAnio
                               Case 1
                                     oWorkSheet.Cells(iFila, 49).Value = "X"
                               Case 2
                                     oWorkSheet.Cells(iFila, 67).Value = "X"
                               Case 3
                                     oWorkSheet.Cells(iFila, 85).Value = "X"
                               Case 4
                                     oWorkSheet.Cells(iFila, 103).Value = "X"
                               End Select
                            End If
                            If lcPenta <> "" Then
                               Select Case oEdad.EdadAnio
                               Case 1
                                    oWorkSheet.Cells(iFila, IIf(lcPenta = "1", 52, IIf(lcPenta = "2", 53, 54))).Value = "X"
                               Case 2
                                    oWorkSheet.Cells(iFila, IIf(lcPenta = "1", 70, IIf(lcPenta = "2", 71, 72))).Value = "X"
                               Case 3
                                    oWorkSheet.Cells(iFila, IIf(lcPenta = "1", 88, IIf(lcPenta = "2", 89, 90))).Value = "X"
                               Case 4
                                    oWorkSheet.Cells(iFila, IIf(lcPenta = "1", 106, IIf(lcPenta = "2", 107, 108))).Value = "X"
                               End Select
                            End If
                            If lcHepaB <> "" Then
                               If oEdad.EdadAnio > 4 Then
                                  oWorkSheet.Cells(iFila, IIf(lcHepaB = "1", 120, IIf(lcHepaB = "2", 121, 122))).Value = "X"
                               End If
                            End If
                        ElseIf oEdad.EdadMes > 0 Then
                            '************edad en meses
                            oWorkSheet.Cells(iFila, 11).Value = lcBCG
                        Else
                            '************edad en dias
                            If oEdad.EdadDia <= 31 Then
                              oWorkSheet.Cells(iFila, 10).Value = lcBCG
                            End If
                            If lcIPV <> "" Then
                               oWorkSheet.Cells(iFila, IIf(lcIPV = "1", 14, 15)).Value = "X"
                            End If
                            If lcAPO <> "" Then
                               oWorkSheet.Cells(iFila, 16).Value = "X"
                            End If
                            If lcPenta <> "" Then
                               oWorkSheet.Cells(iFila, IIf(lcPenta = "1", 19, IIf(lcPenta = "2", 20, 21))).Value = "X"
                            End If
                        End If
                    End If
                     
                 End If
                 iFila = iFila + 1
                 Do While Not mrs_Tmp.EOF And lnIdPaciente = mrs_Tmp!idPaciente
                        mrs_Tmp.MoveNext
                        If mrs_Tmp.EOF Then
                           Exit Do
                        End If
                 Loop
            Loop
            iFila = iFila + 1
            'lnNinoAdolescenteJoven = lnNumTotal
            
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 24
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$10"
                    If oWorkSheet.PageSetup.PrintArea <> "" Then
                       oWorkSheet.PageSetup.PrintArea = SighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila + 3)
                    End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
            End If
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
                Set oWorkSheet = Nothing
                Set oExcel = Nothing
                Set oWorkBookPlantilla = Nothing
                Set oWorkBook = Nothing
            End If
        
        End If
        oConexionExterna.Close
        Set oConexionExterna = Nothing
        Set oRsTmp987 = Nothing
        Set mo_ReglasCaja = Nothing
        Set mo_ReglasLaboratorio = Nothing
        Set oDOPaciente = Nothing
        Set oDOAtencion = Nothing
        Set rsReporte = Nothing
        Set rsDatosMedico = Nothing
        Set oRsImpHIStodos = Nothing

Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    
    Exit Sub
    Resume
End Sub



Sub RegistroDiarioNutricional(lnHwnd As Long, ldFecha As Date)
Dim iFila As Long
Dim lnNumHistorias As Long
Dim lnIdPaciente As Long
Dim lcPaciente As String
Dim lnNumTotal As Long
Dim lnLineas As Integer: Dim lnIdAtencion  As Long
Dim lcProcedencia As String, lcHcFicha As String, lbEsCS As Boolean
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String
Dim lnSeccionDx As Integer 'Frank 15092014
Dim oDOPaciente As New DOPaciente
Dim oDOAtencion As New DOAtencion
Dim rsReporte As New Recordset
Dim mo_ReglasLaboratorio As New ReglasLaboratorio
Dim rsDx As New Recordset
Dim rsCpt As New Recordset
Dim rsDx1 As New Recordset
Dim oRsTmp987 As New Recordset
Dim mo_ReglasCaja As New ReglasCaja
Dim rsDatosMedico As Recordset 'Actualizado 14102014
Dim oRsImpHIStodos As New Recordset
Dim oConexionExterna As New Connection
Dim oConexion As New Connection
Dim lcCodigoHIsResponsable As String 'Actualizado 14102014
Dim lnoRsImpHIStodos As Long
Dim lcAD040 As String
Dim ldFechaHemoglobina As Date, lcHemoglobina As String
Dim lnHistoria As Long, lcDNI As String, lcTipoSeguro As String, lcEdad As String, ldFechaNacimiento As Date, lcDx As String
Dim lnIdTipoSexo As Long
Dim lcDireccion As String, lcDpto As String, lcProvincia As String, lcDistrito As String, lnEdadEnMeses As Long
Dim lnPeso As Double, lnTalla As Integer, lcSachets As String, lnIdTipoFinanacimiento As Long
Dim lcControles As String, lcConsejNutri As String
'
lcAD040 = lcBuscaParametro.SeleccionaFilaParametro(354)
'
oConexionExterna.CursorLocation = adUseClient
oConexionExterna.CommandTimeout = 150
oConexionExterna.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
'
oConexion.CursorLocation = adUseClient
oConexion.CommandTimeout = 150
oConexion.Open SighEntidades.CadenaConexion

'lbEsOpenOffice = IIf(lcBuscaParametro.SeleccionaFilaParametro(284) = "S", True, False)
On Error GoTo ManejadorErrorExcel

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
    mda_FechaInicio = ldFecha
    mda_FechaFin = ldFecha
    'Filtra los Datos
    Set mrs_Tmp = mo_AdminAdmision.BuscaAtencionesCEparaFormatoHIS(ml_IdServicioCE, ml_IdResponsable, mda_FechaInicio, _
                                                mda_FechaFin, "")          'cred
    mrs_Tmp.Filter = "fechaEgreso<>null"
    If mrs_Tmp.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
            If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\CEHis.ods"

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
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                'Crea nueva hoja
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\ceNUTRICIONAL.xls")
                oWorkBookPlantilla.Worksheets("ceNUTRICIONAL").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
            oWorkSheet.Cells(2, 2).Value = "Fecha: " & mda_FechaInicio & "            Responsable de la atención: " & ml_TextoDelFiltro & "      Consultorio: " & ml_TextoDelFiltro1
            iFila = 5
            'oWorkSheet.Cells(iFila, 12).Value = "Fecha: " & mda_FechaInicio
            iFila = iFila + 5
            lnLineas = 1
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               lnIdTipoFinanacimiento = mrs_Tmp!IdFormaPago
               lnHistoria = mrs_Tmp!NroHistoriaClinica
               lcDNI = IIf(IsNull(mrs_Tmp!nroDocumento), "", mrs_Tmp!nroDocumento)
               lcTipoSeguro = mrs_Tmp!ffinanciamiento
               lcPaciente = mrs_Tmp!Paciente1
               lcEdad = mrs_Tmp!Edad & " " & mrs_Tmp!tEdad
               ldFechaNacimiento = mrs_Tmp!FechaNacimiento
               lnIdTipoSexo = mrs_Tmp!idTipoSexo
               lcDireccion = IIf(IsNull(mrs_Tmp!DireccionDomicilio), "", mrs_Tmp!DireccionDomicilio)
               lcDpto = IIf(IsNull(mrs_Tmp!dpto), "", mrs_Tmp!dpto)
               lcProvincia = IIf(IsNull(mrs_Tmp!provincia), "", mrs_Tmp!provincia)
               lcDistrito = IIf(IsNull(mrs_Tmp!distrito), "", mrs_Tmp!distrito)
               lnEdadEnMeses = SighEntidades.DevuelveEdadEnMeses(mrs_Tmp!FechaNacimiento, mrs_Tmp!FechaIngreso)
               lcDx = ""
               lcControles = ""
               lcSachets = ""
               Set rsDx1 = mo_AdminAdmision.BuscaAtencionesDxCEparaFormatoHIS(mrs_Tmp.Fields!idAtencion)
               mo_AdminAdmision.CreaTemporalYcargaDxYcptXgrupo rsDx, rsDx1, rsCpt, mrs_Tmp!idAtencion, True, lcAD040
               If rsDx.RecordCount > 0 Then
                    rsDx.MoveFirst
                    Do While Not rsDx.EOF
                        If Trim(rsDx.Fields("codigoCie10").Value) <> "" Then
                           lcDx = lcDx & rsDx.Fields("codigoCie10").Value & "/"
                        End If
                        Select Case UCase(Trim(rsDx!CodigoCIE10))
                        Case "Z00.1"
                             If Not IsNull(rsDx!labConfHIS) Then
                                lcControles = rsDx!labConfHIS
                             End If
                        Case "Z29.8"
                             'If Not IsNull(rsDx!labConfHIS) Then
                                lcSachets = "30"
                             'End If
                        End Select
                        rsDx.MoveNext
                    Loop
               End If
               '
               ldFechaHemoglobina = 0
               lcHemoglobina = ""
               Set oRsTmp987 = mo_ReglasLaboratorio.LabResultadoPorItemsSeleccionarPorPaciente(mrs_Tmp!idPaciente, 1, oConexion, _
                                                                                               SighEntidades.DevuelveNroOrdenParaHemoglobina(mrs_Tmp!idAtencion))
                                                                                               'Val("999" & Trim(Str(mrs_Tmp!idAtencion))))
               If oRsTmp987.RecordCount > 0 Then
                   oRsTmp987.MoveFirst
                   oRsTmp987.Find "fecha<='" & mrs_Tmp!FechaIngreso & "'"
                   If Not oRsTmp987.EOF Then
                         ldFechaHemoglobina = oRsTmp987!fecha
                         If Not IsNull(oRsTmp987!valorTexto) Then
                            lcHemoglobina = oRsTmp987!valorTexto
                         ElseIf Not IsNull(oRsTmp987!valorNumero) Then
                            lcHemoglobina = Trim(str(oRsTmp987!valorNumero))
                         End If
                   End If
               End If
               oRsTmp987.Close
               '
               lcConsejNutri = ""
               Set rsCpt = mo_AdminAdmision.BuscaOtrosCPTyExamenesparaFormatoHIS(mrs_Tmp.Fields!idCuentaAtencion, sghPuntosCargaBasicos.sghPtoCargaServicioHospitalizacion)
               If rsCpt.RecordCount > 0 Then
                  rsCpt.MoveFirst
                  Do While Not rsCpt.EOF
                     Select Case rsCpt!Codigo
                     Case "99403"
                         lcConsejNutri = "1"
                     End Select
                     rsCpt.MoveNext
                  Loop
               End If
               rsCpt.Close
               '
               lnPeso = 0: lnTalla = 0
               Set oRsTmp987 = mo_AdminAdmision.AtencionCESeleccionarPorIdAtencion(mrs_Tmp!idAtencion, oConexionExterna)
               If oRsTmp987.RecordCount > 0 Then
                  If Not IsNull(oRsTmp987!triajePeso) Then
                  lnPeso = CCur(oRsTmp987!triajePeso)
                  End If
                  If Not IsNull(oRsTmp987!triajeTalla) Then
                  lnTalla = Val(oRsTmp987!triajeTalla)
                  End If
               End If
               oRsTmp987.Close
               '
               lnIdPaciente = mrs_Tmp!idPaciente
               Do While Not mrs_Tmp.EOF And lnIdPaciente = mrs_Tmp!idPaciente
                  mrs_Tmp.MoveNext
                  If mrs_Tmp.EOF Then
                     Exit Do
                  End If
               Loop
               If lbEsOpenOffice = True Then
               Else
                    oWorkSheet.Cells(iFila, 2).Value = lnLineas
                    oWorkSheet.Cells(iFila, 3).Value = lcDNI
                    oWorkSheet.Cells(iFila, 4).Value = HCigualDNI_DevuelveHistoriaConCerosIzquierda(Trim(str(lnHistoria)), True)
                    oWorkSheet.Cells(iFila, IIf(lnIdTipoSexo = 1, 5, 6)).Value = "X"
                    oWorkSheet.Cells(iFila, 7).Value = Format(ldFechaNacimiento, SighEntidades.DevuelveFechaSoloFormato_DMY)
                    oWorkSheet.Cells(iFila, 8).Value = lnEdadEnMeses
                    
                    oWorkSheet.Cells(iFila, 9).Value = lnPeso
                    oWorkSheet.Cells(iFila, 10).Value = lnTalla
                    If lcHemoglobina <> "" Then
                    oWorkSheet.Cells(iFila, 11).Value = lcHemoglobina
                    oWorkSheet.Cells(iFila, 12).Value = Format(ldFechaHemoglobina, SighEntidades.DevuelveFechaSoloFormato_DMY)
                    End If
                    oWorkSheet.Cells(iFila, 13).Value = lcControles
                    oWorkSheet.Cells(iFila, 14).Value = lcSachets
                    oWorkSheet.Cells(iFila, 15).Value = lcConsejNutri
                    If lnIdTipoFinanacimiento = 2 Then
                       oWorkSheet.Cells(iFila, 19).Value = "X"
                    End If
                    
                    oWorkSheet.Cells(iFila, 22).Value = lcProvincia
                    oWorkSheet.Cells(iFila, 23).Value = lcDistrito
                    oWorkSheet.Cells(iFila, 24).Value = lcDireccion
                    
                    
               ldFechaHemoglobina = 0
               lcHemoglobina = ""
                    
                End If
                lnLineas = lnLineas + 1
                iFila = iFila + 1
            Loop
            If lbEsOpenOffice = True Then
                Set PrintArea(0) = ServiceManager.Bridge_GetStruct("com.sun.star.table.CellRangeAddress")
                PrintArea(0).Sheet = 0
                PrintArea(0).startcolumn = 1
                PrintArea(0).StartRow = 0
                PrintArea(0).EndColumn = 24
                PrintArea(0).EndRow = iFila
                Call Feuille.SetPrintAreas(PrintArea())
                Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
                MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
            Else
                oWorkSheet.PageSetup.PrintTitleRows = "$1:$10"
                    If oWorkSheet.PageSetup.PrintArea <> "" Then
                       oWorkSheet.PageSetup.PrintArea = SighEntidades.DevuelveRangoExcelAimprimir(oWorkSheet.PageSetup.PrintArea, iFila + 3)
                    End If
                oExcel.Visible = True
                oWorkSheet.PrintPreview
            End If
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
                Set oWorkSheet = Nothing
                Set oExcel = Nothing
                Set oWorkBookPlantilla = Nothing
                Set oWorkBook = Nothing
            End If
        
        End If
        oConexionExterna.Close
        Set oConexionExterna = Nothing
        Set oRsTmp987 = Nothing
        Set mo_ReglasCaja = Nothing
        Set mo_ReglasLaboratorio = Nothing
        Set oDOPaciente = Nothing
        Set oDOAtencion = Nothing
        Set rsReporte = Nothing
        Set rsDatosMedico = Nothing
        Set oRsImpHIStodos = Nothing

Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    
    Exit Sub
    Resume
End Sub


