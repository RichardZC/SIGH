VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptAHhcNOusadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Historias no usadas
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasArchivoClinico As New SIGHNegocios.ReglasArchivoClinico

Sub EjecutaFormulario()
    Dim oFormulario As New AHhcNOusadas
    oFormulario.Show 1
    Set oFormulario = Nothing
End Sub


Sub CreaDatosParaReporte(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, lnHCmaximasXlinea As Integer, lnNroLineasMaximas As Long, lbOrdenMenorAmayor As Boolean, lnHwnd As Long)
     Dim mrs_Tmp As New Recordset
     Dim mrs_Tmp1 As New Recordset
     Dim lnUltimaHC As Long, lnContHC As Long
     Dim oConexion As New Connection
     Dim lbHcVacia As Boolean
     Dim lclinea As String, lnNroHistoriasXlinea As Integer
     Dim lnLinea As Long
     Dim lcColumna1 As String, lcColumna2 As String, lcColumna3 As String, lcColumna4 As String
     Dim lcColumna5 As String, lcColumna6 As String, lcColumna7 As String
     If lnHCmaximasXlinea > 20 Then
        lnHCmaximasXlinea = 20
     End If
     If lnNroLineasMaximas > 320000 Then
        lnNroLineasMaximas = 310000
     End If
     oConexion.CommandTimeout = 300
     oConexion.CursorLocation = adUseClient
     oConexion.Open sighEntidades.CadenaConexion
     Set mrs_Tmp1 = mo_ReglasArchivoClinico.GeneradorNroHistoriaClinicaXidTipoNumeracion(1, oConexion)
     lnUltimaHC = 0
     If mrs_Tmp1.RecordCount > 0 Then
        lnUltimaHC = mrs_Tmp1.Fields!nrohistoriaclinica
     End If
     mrs_Tmp1.Close
     If lnUltimaHC > 0 Then
         With mrs_Tmp
          .Fields.Append "Columna1", adVarChar, 20, adFldIsNullable
          .Fields.Append "Columna2", adVarChar, 20, adFldIsNullable
          .Fields.Append "Columna3", adVarChar, 20, adFldIsNullable
          .Fields.Append "Columna4", adVarChar, 20, adFldIsNullable
          .Fields.Append "Columna5", adVarChar, 20, adFldIsNullable
          .Fields.Append "Columna6", adVarChar, 20, adFldIsNullable
          .Fields.Append "Columna7", adVarChar, 20, adFldIsNullable
          .LockType = adLockOptimistic
          .Open
         End With
         lnContHC = 1
         lclinea = "": lnNroHistoriasXlinea = 0
         If lbOrdenMenorAmayor = True Then
            lnContHC = 1
         Else
            lnContHC = lnUltimaHC
         End If
         lnLinea = 0
         lcColumna1 = "": lcColumna2 = "": lcColumna3 = "": lcColumna4 = "": lcColumna5 = "": lcColumna6 = "": lcColumna7 = ""
         Do While True
            '
            Set mrs_Tmp1 = mo_ReglasArchivoClinico.HistoriaClinicasXidentificador(lnContHC, oConexion)
            lbHcVacia = True
            If mrs_Tmp1.RecordCount > 0 Then
               If mrs_Tmp1.Fields!IdPaciente > 0 Then
                  lbHcVacia = False
               End If
            End If
            mrs_Tmp1.Close
            If lbHcVacia = True Then
               If lclinea = "" Then
                  lclinea = Trim(Str(lnContHC))
               Else
                  lclinea = lclinea & " " & Trim(Str(lnContHC))
               End If
               lnNroHistoriasXlinea = lnNroHistoriasXlinea + 1
               Select Case lnNroHistoriasXlinea
               Case 1
                  lcColumna1 = Trim(Str(lnContHC))
               Case 2
                  lcColumna2 = Trim(Str(lnContHC))
               Case 3
                  lcColumna3 = Trim(Str(lnContHC))
               Case 4
                  lcColumna4 = Trim(Str(lnContHC))
               Case 5
                  lcColumna5 = Trim(Str(lnContHC))
               Case 6
                  lcColumna6 = Trim(Str(lnContHC))
               Case Else
                  lcColumna7 = Trim(Str(lnContHC))
               End Select
               
               If lnNroHistoriasXlinea >= lnHCmaximasXlinea Then
                  mrs_Tmp.AddNew
                  mrs_Tmp.Fields!columna1 = lcColumna1
                  mrs_Tmp.Fields!columna2 = lcColumna2
                  mrs_Tmp.Fields!columna3 = lcColumna3
                  mrs_Tmp.Fields!columna4 = lcColumna4
                  mrs_Tmp.Fields!columna5 = lcColumna5
                  mrs_Tmp.Fields!columna6 = lcColumna6
                  mrs_Tmp.Fields!columna7 = lcColumna7
                  mrs_Tmp.Update
                  lclinea = ""
                  lnNroHistoriasXlinea = 0
                  lnLinea = lnLinea + 1
                  lcColumna1 = "": lcColumna2 = "": lcColumna3 = "": lcColumna4 = "": lcColumna5 = "": lcColumna6 = "": lcColumna7 = ""
               End If
            End If
            '
            If lbOrdenMenorAmayor = True Then
               lnContHC = lnContHC + 1
            Else
               lnContHC = lnContHC - 1
            End If
            '
            If lnLinea >= lnNroLineasMaximas Then
               Exit Do
            End If
            '
         Loop
         If lclinea <> "" Then
'            mrs_Tmp.AddNew
'            mrs_Tmp.Fields!HistoriasVacias = lclinea
'            mrs_Tmp.Update
         End If
     Else
         MsgBox "No se actualizó el último correlativo de HISTORIAS GENERADAS AUTOMATICAMENTE", vbInformation, "Reporte"
         Exit Sub
     End If
     If mrs_Tmp.RecordCount = 0 Then
        MsgBox "No existe Historias Clínicas NO USADAS", vbInformation, "Reporte"
     Else
        If lbEnExcel = True Then
           Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, "", lnHwnd
'        Else
'            Set RpAHCauditoria.DataSource = mrs_Tmp
'            RpAHCauditoria.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
'            RpAHCauditoria.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
'            Set RpAHCauditoria.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
'            RpAHCauditoria.Sections("pie").Controls("lblPie").Caption = ""
'            RpAHCauditoria.Orientation = rptOrientLandscape
'            RpAHCauditoria.Show 1
        End If
     End If
     
End Sub

Sub CreaDatosParaReporteHcTemporales(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, lnHwnd As Long)
     Dim mrs_Tmp As New Recordset
     Dim lcSql As String, lcPie As String
     Set mrs_Tmp = mo_ReglasArchivoClinico.PacientesConHistoriasTemporales
     lcPie = Trim(Str(mrs_Tmp.RecordCount))
     If Val(lcPie) = 0 Then
        MsgBox "No existen temporales", vbInformation, "Reporte"
     Else
        lcPie = "N° Registros: " & lcPie
        If lbEnExcel = True Then
           Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
           mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
'        Else
'            Set RpAHCauditoria.DataSource = mrs_Tmp
'            RpAHCauditoria.Sections("cabecera").Controls("lblTitulo").Caption = lcTitulo
'            RpAHCauditoria.Sections("cabecera").Controls("lblSubTitulo").Caption = lcSubTitulo
'            Set RpAHCauditoria.Sections("cabecera").Controls("image1").Picture = LoadPicture(App.Path & "\imagenes\Imagen de reportes.jpg")
'            RpAHCauditoria.Sections("pie").Controls("lblPie").Caption = ""
'            RpAHCauditoria.Orientation = rptOrientLandscape
'            RpAHCauditoria.Show 1
        End If
     End If
     
     
End Sub



