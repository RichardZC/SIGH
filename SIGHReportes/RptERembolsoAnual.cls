VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptERembolsoAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reembolsos
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_idUsuario As Long
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim mo_CabeceraReportes As New SIGHNegocios.ReglasComunes

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Sub EjecutaFormulario()
    Dim oFormulario As New EReembolsoAnual
    oFormulario.idUsuario = ml_idUsuario
    oFormulario.Show 1
End Sub

Sub CrearReporte_excel(ml_TextoDelFiltro As String, lnAnio As Long, lnIdAreaTramitaSeguros As Long, _
                       lnIdFuenteFinanciamiento As Long, lnHwnd As Long)
Dim rsReporte As New Recordset
Dim mrs_Tmp As New Recordset
Dim iFila As Long, lnTotal As Double
Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim lbEsOpenOffice As Boolean
Dim lcNombre As String, lcSql As String

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
    'Filtra los Datos
    Set rsReporte = mo_ReglasFacturacion.FactReembolsosSelecionarPorAnioYAreatsYPlan(lnAnio, lnIdAreaTramitaSeguros, lnIdFuenteFinanciamiento)
    If rsReporte.RecordCount = 0 Then
        MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
    Else
        With mrs_Tmp
              .Fields.Append "Reembolso", adVarChar, 50
              .Fields.Append "Ene", adDouble
              .Fields.Append "Feb", adDouble
              .Fields.Append "Mar", adDouble
              .Fields.Append "Abr", adDouble
              .Fields.Append "May", adDouble
              .Fields.Append "Jun", adDouble
              .Fields.Append "Jul", adDouble
              .Fields.Append "Ago", adDouble
              .Fields.Append "Sep", adDouble
              .Fields.Append "Oct", adDouble
              .Fields.Append "Nov", adDouble
              .Fields.Append "Dic", adDouble
              .LockType = adLockOptimistic
              .Open
              'saldo Inicial
              .AddNew
              .Fields!Reembolso = "1"
              .Update
              'Produccion
              .AddNew
              .Fields!Reembolso = "2"
              .Update
              'produccion sin reembolsar
              .AddNew
              .Fields!Reembolso = "3"
              .Update
              'Reembolsado
              .AddNew
              .Fields!Reembolso = "4"
              .Update
              'saldo Final
              .AddNew
              .Fields!Reembolso = "5"
              .Update
        End With
        rsReporte.MoveFirst
        Do While Not rsReporte.EOF
           Select Case rsReporte.Fields!Mes
           Case 1
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Ene = mrs_Tmp.Fields!Ene + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Ene = mrs_Tmp.Fields!Ene + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Ene = mrs_Tmp.Fields!Ene + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Ene = mrs_Tmp.Fields!Ene + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Ene = mrs_Tmp.Fields!Ene + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 2
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Feb = mrs_Tmp.Fields!Feb + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Feb = mrs_Tmp.Fields!Feb + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Feb = mrs_Tmp.Fields!Feb + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Feb = mrs_Tmp.Fields!Feb + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Feb = mrs_Tmp.Fields!Feb + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 3
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Mar = mrs_Tmp.Fields!Mar + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Mar = mrs_Tmp.Fields!Mar + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Mar = mrs_Tmp.Fields!Mar + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Mar = mrs_Tmp.Fields!Mar + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Mar = mrs_Tmp.Fields!Mar + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 4
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Abr = mrs_Tmp.Fields!Abr + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Abr = mrs_Tmp.Fields!Abr + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Abr = mrs_Tmp.Fields!Abr + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Abr = mrs_Tmp.Fields!Abr + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Abr = mrs_Tmp.Fields!Abr + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 5
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!May = mrs_Tmp.Fields!May + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!May = mrs_Tmp.Fields!May + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!May = mrs_Tmp.Fields!May + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!May = mrs_Tmp.Fields!May + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!May = mrs_Tmp.Fields!May + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 6
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Jun = mrs_Tmp.Fields!Jun + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Jun = mrs_Tmp.Fields!Jun + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Jun = mrs_Tmp.Fields!Jun + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Jun = mrs_Tmp.Fields!Jun + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Jun = mrs_Tmp.Fields!Jun + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 7
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Jul = mrs_Tmp.Fields!Jul + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Jul = mrs_Tmp.Fields!Jul + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Jul = mrs_Tmp.Fields!Jul + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Jul = mrs_Tmp.Fields!Jul + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Jul = mrs_Tmp.Fields!Jul + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 8
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Ago = mrs_Tmp.Fields!Ago + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Ago = mrs_Tmp.Fields!Ago + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Ago = mrs_Tmp.Fields!Ago + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Ago = mrs_Tmp.Fields!Ago + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Ago = mrs_Tmp.Fields!Ago + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 9
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Sep = mrs_Tmp.Fields!Sep + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Sep = mrs_Tmp.Fields!Sep + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Sep = mrs_Tmp.Fields!Sep + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Sep = mrs_Tmp.Fields!Sep + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Sep = mrs_Tmp.Fields!Sep + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 10
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Oct = mrs_Tmp.Fields!Oct + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Oct = mrs_Tmp.Fields!Oct + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Oct = mrs_Tmp.Fields!Oct + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Oct = mrs_Tmp.Fields!Oct + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Oct = mrs_Tmp.Fields!Oct + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 11
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Nov = mrs_Tmp.Fields!Nov + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Nov = mrs_Tmp.Fields!Nov + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Nov = mrs_Tmp.Fields!Nov + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Nov = mrs_Tmp.Fields!Nov + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Nov = mrs_Tmp.Fields!Nov + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           Case 12
              'saldo inicial
               mrs_Tmp.MoveFirst
               mrs_Tmp.Fields!Dic = mrs_Tmp.Fields!Dic + rsReporte.Fields!saldoInicial
               mrs_Tmp.Update
              'Produccion
               mrs_Tmp.Move 1, adBookmarkFirst
               mrs_Tmp.Fields!Dic = mrs_Tmp.Fields!Dic + rsReporte.Fields!ConsumoPorReembolsar
               mrs_Tmp.Update
              'Produccion sin Reembolsar
               mrs_Tmp.Move 2, adBookmarkFirst
               mrs_Tmp.Fields!Dic = mrs_Tmp.Fields!Dic + rsReporte.Fields!ReembolsoPorPagar
               mrs_Tmp.Update
              'Reembolsado
               mrs_Tmp.Move 3, adBookmarkFirst
               mrs_Tmp.Fields!Dic = mrs_Tmp.Fields!Dic + rsReporte.Fields!ReembolsoPagado
               mrs_Tmp.Update
              'Saldo Final
               mrs_Tmp.Move 4, adBookmarkFirst
               mrs_Tmp.Fields!Dic = mrs_Tmp.Fields!Dic + rsReporte.Fields!SaldoFinal
               mrs_Tmp.Update
           End Select
           rsReporte.MoveNext
        Loop
        If mrs_Tmp.RecordCount = 0 Then
            MsgBox "No existe información con esos Datos", vbInformation, "Resultado"
        Else
            If lbEsOpenOffice = True Then
                'Abre el archivo ExcelOpenOffice
                lcArchivoExcel = App.Path + "\Plantillas\HReembolsoAnual.ods"
'                FileCopy lcArchivoExcel, App.Path + "\Plantillas\OpenOffice.ods"
'                Chemin = "file:///" & App.Path & "\Plantillas\"
'                Chemin = Replace(Chemin, "\", "/")
'                Fichier = Chemin & "/OpenOffice.ods"
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
                'Encabezado de Pagina
                mo_CabeceraReportes.CabeceraReportes Document, True
                ' Pone la ventana en primer plano, pasándole el Hwnd
                ret = SetForegroundWindow(lnHwnd)
            Else
                Set oExcel = GalenhosExcelApplication()  'New Excel.Application
                'Crea nueva hoja
                Set oWorkBook = oExcel.Workbooks.Add
                'Abre, copia y cierra la plantilla
                Set oWorkBookPlantilla = oExcel.Workbooks.Open(App.Path + "\Plantillas\HReembolsoAnual.xls")
                oWorkBookPlantilla.Worksheets("HReembolsoAnual").Copy Before:=oWorkBook.Sheets(1)
                oWorkBookPlantilla.Close
                'Activa la primera hoja
                Set oWorkSheet = oWorkBook.Sheets(1)
                mo_CabeceraReportes.CabeceraReportes oWorkSheet, False
            End If
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(1, 2).setFormula(ml_TextoDelFiltro)
                Else
                    oWorkSheet.Cells(3, 2).Value = ml_TextoDelFiltro
                End If
            iFila = 7: lnTotal = 0
            mrs_Tmp.MoveFirst
            Do While Not mrs_Tmp.EOF
               lnTotal = mrs_Tmp.Fields("Ene").Value + mrs_Tmp.Fields("Feb").Value + mrs_Tmp.Fields("Mar").Value _
                         + mrs_Tmp.Fields("Abr").Value + mrs_Tmp.Fields("May").Value + mrs_Tmp.Fields("Jun").Value _
                         + mrs_Tmp.Fields("Jul").Value + mrs_Tmp.Fields("Ago").Value + mrs_Tmp.Fields("Sep").Value _
                         + mrs_Tmp.Fields("Oct").Value + mrs_Tmp.Fields("Nov").Value + mrs_Tmp.Fields("Dic").Value
               
                If lbEsOpenOffice = True Then
                    Call Feuille.getcellbyposition(2, iFila - 1).setFormula(mrs_Tmp.Fields("Ene").Value)
                    Call Feuille.getcellbyposition(3, iFila - 1).setFormula(mrs_Tmp.Fields("Feb").Value)
                    Call Feuille.getcellbyposition(4, iFila - 1).setFormula(mrs_Tmp.Fields("Mar").Value)
                    Call Feuille.getcellbyposition(5, iFila - 1).setFormula(mrs_Tmp.Fields("Abr").Value)
                    Call Feuille.getcellbyposition(6, iFila - 1).setFormula(mrs_Tmp.Fields("May").Value)
                    Call Feuille.getcellbyposition(7, iFila - 1).setFormula(mrs_Tmp.Fields("Jun").Value)
                    Call Feuille.getcellbyposition(8, iFila - 1).setFormula(mrs_Tmp.Fields("Jul").Value)
                    Call Feuille.getcellbyposition(9, iFila - 1).setFormula(mrs_Tmp.Fields("Ago").Value)
                    Call Feuille.getcellbyposition(10, iFila - 1).setFormula(mrs_Tmp.Fields("Sep").Value)
                    Call Feuille.getcellbyposition(11, iFila - 1).setFormula(mrs_Tmp.Fields("Oct").Value)
                    Call Feuille.getcellbyposition(12, iFila - 1).setFormula(mrs_Tmp.Fields("Nov").Value)
                    Call Feuille.getcellbyposition(13, iFila - 1).setFormula(mrs_Tmp.Fields("Dic").Value)
                    Call Feuille.getcellbyposition(14, iFila - 1).setFormula(lnTotal)
                Else
                    oWorkSheet.Cells(iFila, 3).Value = mrs_Tmp.Fields("Ene").Value
                    oWorkSheet.Cells(iFila, 4).Value = mrs_Tmp.Fields("Feb").Value
                    oWorkSheet.Cells(iFila, 5).Value = mrs_Tmp.Fields("Mar").Value
                    oWorkSheet.Cells(iFila, 6).Value = mrs_Tmp.Fields("Abr").Value
                    oWorkSheet.Cells(iFila, 7).Value = mrs_Tmp.Fields("May").Value
                    oWorkSheet.Cells(iFila, 8).Value = mrs_Tmp.Fields("Jun").Value
                    oWorkSheet.Cells(iFila, 9).Value = mrs_Tmp.Fields("Jul").Value
                    oWorkSheet.Cells(iFila, 10).Value = mrs_Tmp.Fields("Ago").Value
                    oWorkSheet.Cells(iFila, 11).Value = mrs_Tmp.Fields("Sep").Value
                    oWorkSheet.Cells(iFila, 12).Value = mrs_Tmp.Fields("Oct").Value
                    oWorkSheet.Cells(iFila, 13).Value = mrs_Tmp.Fields("Nov").Value
                    oWorkSheet.Cells(iFila, 14).Value = mrs_Tmp.Fields("Dic").Value
                    oWorkSheet.Cells(iFila, 15).Value = lnTotal
                End If
                    iFila = iFila + 1
                    mrs_Tmp.MoveNext
            Loop
            If lbEsOpenOffice = True Then
               Call Document.getCurrentController.getFrame.getContainerWindow.setVisible(True)
            Else
              oWorkSheet.PageSetup.PrintTitleRows = "$1:$6"
              oExcel.Visible = True
              oWorkSheet.PrintPreview
            End If
        End If
        If lbEsOpenOffice = True Then
            MsgBox "El Reporte se generó en forma exitosa: " & lcArchivoExcel, vbInformation
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
    End If
Exit Sub
ManejadorErrorExcel:
    Select Case Err.Number
    Case 1004
        MsgBox "No hay impresoras instaladas. Para instalar una impresora, elija Configuración en el menú Inicio de Windows, haga clic en Impresoras y después haga doble clic en Agregar impresora. Siga las instrucciones del asistente.", vbExclamation, "Reporte de historia clínica"
    Case Else
        MsgBox Err.Description
    End Select
    Exit Sub
End Sub



