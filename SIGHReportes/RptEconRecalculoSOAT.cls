VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RptEconRecalculoSOAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Recalculo del SOAT
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
'***************daniel barrantes**************
'***************Filtro de Datos, Configuración y Emisión del Reporte
'***************Consideraciones:
Option Explicit
Dim ml_idUsuario  As Long
Const lnIdEspecialidad As Long = 23    'Neonatologia
Property Let idUsuario(lValue As Long)
    ml_idUsuario = lValue
End Property

Sub EjecutaFormulario()
    Dim oERecalculoSOAT As New ERecalculoSOAT
    oERecalculoSOAT.Show 1
    Set oERecalculoSOAT = Nothing
End Sub


Sub CreaDatosParaReporte(lbEnExcel As Boolean, lcTitulo As String, lcSubTitulo As String, ml_idFuenteFinanciamiento As Long, _
                         mda_FechaInicio As Date, mda_FechaFin As Date, lnHwnd As Long)
    Dim mrs_Tmp As New Recordset
    Dim mrs_Tmp1 As New Recordset
    Dim mrs_Tmp2 As New Recordset
    Dim mrs_Tmp3 As New Recordset
    Dim mrs_Tmp4 As New Recordset
    Dim mrs_Cab As New Recordset
    Dim mrs_Det As New Recordset
    Dim mrs_Shape As New Recordset
    Dim ldFecha As Date, ldFechaSolo As Date
    Dim lcHoraSolo As String, lcUsuario As String, lcNusuario As String, lcAccion As String
    Dim lcPc As String, lcTexto1 As String, lcTexto2 As String, lcPie As String
    Dim lnIdListItem As Long, lnNroCuenta As Long, ldFechaIngreso As Date
    Dim lcHoraIngreso As String, lcPaciente As String, lcSexo As String
    Dim lcMotivo As String, lcEstancia As String, ldFechaEgreso As Date
    Dim lcHoraEgreso As String, lcDestino As String, lcRecibe As String
    Dim lcMedicoIE As String, lnImporte As Double, lcDescripcion As String
    Dim lcNombreTablaCab As String, lcNombreTablaDet As String
    Dim lnEstancia As Integer, lnIdServicio As Long, lnTotalG As Double
    Dim mo_AdminReportes As New SIGHNegocios.ReglasReportes
    Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
    Dim mo_ReglasDeSeguridad As New SIGHNegocios.ReglasDeSeguridad
    Dim mo_AdminCaja As New SIGHNegocios.ReglasCaja
    Dim mo_ReglasAdmision As New SIGHNegocios.ReglasAdmision
    Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
    Dim oDoCatalogoServicioHosp As New DOFinanciamientoCatalogoServ
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim lcSql As String, lnIdEstancia As Long
    Set mrs_Tmp1 = mo_ReglasDeSeguridad.AuditoriaFiltrarSoloRecalculos
    If mrs_Tmp1.RecordCount > 0 Then
       With mrs_Tmp
           .Fields.Append "idCuentaAtencion", adInteger
           .Fields.Append "Paciente", adVarChar, 160, adFldIsNullable
           .Fields.Append "Descripcion", adVarChar, 100, adFldIsNullable
           .Fields.Append "Cantidad", adInteger
           .Fields.Append "PrecioBoleta", adDouble
           .Fields.Append "ImporteBoleta", adDouble
           .Fields.Append "PrecioParticular", adDouble
           .Fields.Append "ImporteParticular", adDouble
           .Fields.Append "ImpBoletaMenosImpParticular", adDouble
           .LockType = adLockOptimistic
           .Open
       End With
       lnImporte = 0

       mrs_Tmp1.MoveFirst
       
       Do While Not mrs_Tmp1.EOF
          lnIdListItem = mrs_Tmp1.Fields!idRegistro
If mrs_Tmp1.Fields!idRegistro = 33559 Then
lcSql = ""
End If
          If InStr(mrs_Tmp1.Fields!observaciones, "IAFA: SOAT") > 0 Then

             Set mrs_Tmp2 = mo_ReglasAdmision.atencionesSeleccionarPorIdCuentaAtencion(mrs_Tmp1.Fields!idRegistro, mda_FechaInicio, mda_FechaFin)
             If mrs_Tmp2.RecordCount > 0 Then
                 mrs_Tmp2.Close
                 lnIdEstancia = lcBuscaParametro.SeleccionaFilaParametro(202)
                 'Consumo de Servicios
                 Set mrs_Tmp2 = mo_AdminReportes.BoletasDeServicioParaSOATxCuenta(mrs_Tmp1.Fields!idRegistro)
                 If mrs_Tmp2.RecordCount > 0 Then
                    mrs_Tmp2.MoveFirst
                    Do While Not mrs_Tmp2.EOF
                       If mrs_Tmp2.Fields!Codigo <> "F00001" Then
                            Set mrs_Tmp3 = mo_ReglasComunes.FactCatalogoServiciosHospXfiltro("idTipoFinanciamiento=1 and idProducto=" & mrs_Tmp2.Fields!idProducto)
                            lnTotalG = 0
                            If mrs_Tmp3.RecordCount > 0 Then
                               lnTotalG = mrs_Tmp3.Fields!PrecioUnitario
                               If lnIdEstancia = mrs_Tmp2.Fields!idProducto Then
                                  lnTotalG = Round(lnTotalG / 24, 4)
                               End If
                            End If
                            mrs_Tmp3.Close
                            mrs_Tmp.AddNew
                            mrs_Tmp.Fields!idCuentaAtencion = mrs_Tmp2.Fields!idCuentaAtencion
                            mrs_Tmp.Fields!Paciente = Left(Trim(Str(mrs_Tmp2.Fields!nrohistoriaclinica)) & " - " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre), 100)
                            mrs_Tmp.Fields!Descripcion = Left("S: " & Trim(mrs_Tmp2.Fields!Codigo) & " " & mrs_Tmp2.Fields!Nombre, 100)
                            mrs_Tmp.Fields!Cantidad = mrs_Tmp2.Fields!Cantidad
                            mrs_Tmp.Fields!PrecioBoleta = mrs_Tmp2.Fields!Precio
                            mrs_Tmp.Fields!ImporteBoleta = mrs_Tmp2.Fields!Total
                            mrs_Tmp.Fields!PrecioParticular = lnTotalG
                            mrs_Tmp.Fields!ImporteParticular = Round(mrs_Tmp2.Fields!Cantidad * lnTotalG, 2)
                            mrs_Tmp.Fields!ImpBoletaMenosImpParticular = mrs_Tmp2.Fields!Total - Round(mrs_Tmp2.Fields!Cantidad * lnTotalG, 2)
                            mrs_Tmp.Update
                            lnImporte = lnImporte + mrs_Tmp.Fields!ImpBoletaMenosImpParticular
                       End If
                       mrs_Tmp2.MoveNext
                    Loop
                 End If
                 mrs_Tmp2.Close
                 'Consumo de Farmacia
                 Set mrs_Tmp2 = mo_AdminReportes.BoletasDeFarmaciaParaSOATxCuenta(mrs_Tmp1.Fields!idRegistro)
                 If mrs_Tmp2.RecordCount > 0 Then
                    mrs_Tmp2.MoveFirst
                    Do While Not mrs_Tmp2.EOF
                       Set mrs_Tmp3 = mo_ReglasComunes.FactCatalogoBienesInsumosHospXfiltro("idTipoFinanciamiento=1 and idProducto=" & mrs_Tmp2.Fields!idProducto)
                       lnTotalG = 0
                       If mrs_Tmp3.RecordCount > 0 Then
                          lnTotalG = mrs_Tmp3.Fields!PrecioUnitario
                       End If
                       mrs_Tmp3.Close
                       mrs_Tmp.AddNew
                       mrs_Tmp.Fields!idCuentaAtencion = mrs_Tmp2.Fields!idCuentaAtencion
                       mrs_Tmp.Fields!Paciente = Left(Trim(Str(mrs_Tmp2.Fields!nrohistoriaclinica)) & " - " & Trim(mrs_Tmp2.Fields!ApellidoPaterno) & " " & Trim(mrs_Tmp2.Fields!ApellidoMaterno) & " " & Trim(mrs_Tmp2.Fields!PrimerNombre), 100)
                       mrs_Tmp.Fields!Descripcion = Left("F: " & Trim(mrs_Tmp2.Fields!Codigo) & " " & mrs_Tmp2.Fields!Nombre, 100)
                       mrs_Tmp.Fields!Cantidad = mrs_Tmp2.Fields!Cantidad
                       mrs_Tmp.Fields!PrecioBoleta = mrs_Tmp2.Fields!Precio
                       mrs_Tmp.Fields!ImporteBoleta = mrs_Tmp2.Fields!Total
                       mrs_Tmp.Fields!PrecioParticular = lnTotalG
                       mrs_Tmp.Fields!ImporteParticular = Round(mrs_Tmp2.Fields!Cantidad * lnTotalG, 2)
                       mrs_Tmp.Fields!ImpBoletaMenosImpParticular = mrs_Tmp2.Fields!Total - Round(mrs_Tmp2.Fields!Cantidad * lnTotalG, 2)
                       mrs_Tmp.Update
                       lnImporte = lnImporte + mrs_Tmp.Fields!ImpBoletaMenosImpParticular
                       mrs_Tmp2.MoveNext
                    Loop
                 End If
                 mrs_Tmp2.Close
              Else
                 mrs_Tmp2.Close
              End If
          End If
          Do While Not mrs_Tmp1.EOF And lnIdListItem = mrs_Tmp1.Fields!idRegistro
              mrs_Tmp1.MoveNext
              If mrs_Tmp1.EOF Then
                 Exit Do
              End If
          Loop
       Loop
       If mrs_Tmp.RecordCount = 0 Then
          MsgBox "No existe Información con esos Datos", vbInformation, "Reporte"
       Else
          If lbEnExcel = True Then
             lcPie = "Suma de columna 'ImpBoletaMenosImpParticular' = " & lnImporte
             Dim mo_ReglasReportes As New SIGHNegocios.ReglasReportes
             mo_ReglasReportes.ExportarRecordSetAexcel mrs_Tmp, lcTitulo, lcSubTitulo, lcPie, lnHwnd
          Else
             'falta
          End If
       End If
    Else
       MsgBox "No existe Información con esos Datos", vbInformation, "Reporte"
    End If
End Sub






