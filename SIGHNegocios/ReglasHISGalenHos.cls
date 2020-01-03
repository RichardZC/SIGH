VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasHISGalenos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reglas HIS GALENHOS
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ms_MensajeError As String
Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos  'Valor de Mensaje de Error
Dim lcBuscaParametro As New SIGHDatos.Parametros                                'Valores de Parametros del Sistema
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad                    'Ingreso al registro de Auditoria
Dim mo_AdminAdmision As New SIGHNegocios.ReglasAdmision

Dim mo_Paciente As New SIGHDatos.Pacientes
Dim mo_DatosServiciosPorEstab As New SIGHDatos.HIS_ServEstablecimiento         'Contiene los Servicios dependiendo del Establecimiento

Property Let MensajeError(sValue As String)
  ms_MensajeError = sValue
End Property

Property Get MensajeError() As String
  MensajeError = ms_MensajeError
End Property

'============================================== LISTADOS DE FORMULARIO HIS ==========================================
Public Function ListaMeses() As Recordset
Dim rcs_meses As New Recordset
Set rcs_meses = Nothing

rcs_meses.CursorType = adOpenStatic
rcs_meses.Fields.Append "IdMes", adInteger, 50, adFldIsNullable
rcs_meses.Fields.Append "NombreMes", adVarChar, 20, adFldIsNullable
rcs_meses.Open

Dim I As Integer
For I = 1 To 12
    rcs_meses.AddNew
    rcs_meses.Fields(0) = I
    rcs_meses.Fields(1) = DevuelveNombreMes(I)
    rcs_meses.Update
Next I

rcs_meses.MoveFirst
Set ListaMeses = rcs_meses
End Function

Public Function ListaTurnos() As Recordset
Dim oConexion As New ADODB.Connection
Dim mo_ListaTurnos As New HIS_Turnos
On Error GoTo ListarTurnos_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_ListaTurnos.Conexion = oConexion
Set ListaTurnos = mo_ListaTurnos.ListaTurnos

On Error GoTo 0
Exit Function
ListarTurnos_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListarTurnos of Módulo de clase ReglasHISGalenos"
End Function

Public Function ListaTiposDocumento() As Recordset
Dim oConexion As New ADODB.Connection
Dim mo_His_cabecera As New SIGHDatos.HIS_Cabecera

On Error GoTo ListaTiposDocumento_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion
Set ListaTiposDocumento = mo_His_cabecera.ListaTiposDocumentos

On Error GoTo 0
Exit Function
ListaTiposDocumento_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaTiposDocumento of Módulo de clase ReglasHISGalenos"
End Function

Public Function ListaFuentesFinanciamiento() As Recordset
Dim oConexion As New ADODB.Connection
Dim mo_His_cabecera As New SIGHDatos.HIS_Cabecera
On Error GoTo ListaFuentesFinanciamiento_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion
Set ListaFuentesFinanciamiento = mo_His_cabecera.ListaFuentesFinanciamiento

On Error GoTo 0
Exit Function
ListaFuentesFinanciamiento_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaTiposFinanciamiento of Módulo de clase ReglasHISGalenos"
End Function

Public Function ListaEtnias() As Recordset
Dim oConexion As New ADODB.Connection
Dim mo_His_cabecera As New SIGHDatos.HIS_Cabecera
On Error GoTo ListaEtnias_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion
Set ListaEtnias = mo_His_cabecera.ListaEtnias

On Error GoTo 0
Exit Function
ListaEtnias_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaEtnias of Módulo de clase ReglasHISGalenos"
End Function

Public Function ListaTiposEdad() As Recordset
Dim oConexion As New ADODB.Connection
Dim mo_DatosTipoEdad As New SIGHDatos.HIS_TipoEdad
On Error GoTo ListaTiposEdad_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_DatosTipoEdad.Conexion = oConexion
Set ListaTiposEdad = mo_DatosTipoEdad.ListaTiposEdad

On Error GoTo 0
Exit Function
ListaTiposEdad_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaTiposEdad of Módulo de clase ReglasHISGalenos"
End Function

Public Function ListaTiposSexo() As Recordset
Dim oConexion As New ADODB.Connection
Dim mo_His_cabecera As New SIGHDatos.HIS_Cabecera
On Error GoTo ListaTiposSexo_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion
Set ListaTiposSexo = mo_His_cabecera.ListaGeneros

On Error GoTo 0
Exit Function
ListaTiposSexo_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaTiposSexo of Módulo de clase ReglasHISGalenos"
End Function

Public Function ListaSituacionPaciente() As Recordset
Dim oConexion As New ADODB.Connection
Dim mo_His_cabecera As New SIGHDatos.HIS_Cabecera
On Error GoTo ListaSituacionPacienteFrenteaEstablecimiento_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion
Set ListaSituacionPaciente = mo_His_cabecera.ListaEstadosPaciente

On Error GoTo 0
Exit Function
ListaSituacionPacienteFrenteaEstablecimiento_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaSituacionPacienteFrenteaEstablecimiento of Módulo de clase ReglasHISGalenos"
End Function

Public Function ListaTiposDiagnosticos() As Recordset
Dim oConexion As New ADODB.Connection
Dim mo_DatosSubClDiagnosticos As New SIGHDatos.SubclasificacionDiagnosticos
On Error GoTo ListaTiposDiagnosticos_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_DatosSubClDiagnosticos.Conexion = oConexion
Set ListaTiposDiagnosticos = mo_DatosSubClDiagnosticos.SeleccionarDxConsultaExterna()

On Error GoTo 0
Exit Function
ListaTiposDiagnosticos_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaTiposDiagnosticos of Módulo de clase ReglasHISGalenos"
End Function

'Public Function ObtenerListaCodigosNaciones() As Recordset
'Dim oConexion As New ADODB.Connection
'Dim mo_DatosHIS_Cabecera As New SIGHDatos.HIS_Cabecera
'
'On Error GoTo ObtenerListaCodigosNaciones_Error
'oConexion.Open sighentidades.CadenaConexion
'oConexion.CursorLocation = adUseClient
'Set mo_DatosHIS_Cabecera.Conexion = oConexion
'Set ObtenerListaCodigosNaciones = mo_DatosHIS_Cabecera.ListaPaises
'
'On Error GoTo 0
'Exit Function
'ObtenerListaCodigosNaciones_Error:
'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerListaCodigosNaciones of Módulo de clase ReglasHISGalenos"
'End Function

'MODIFICADO POR YEPE NOVIEMBRE Inicio
'Public Function ObtenerListaCodigosNaciones() As Recordset
Public Function ObtenerListaCodigosNaciones(ByVal Codigo As String, ByVal nombre As String) As Recordset
Dim oConexion As New ADODB.Connection
Dim mo_DatosHIS_Cabecera As New SIGHDatos.HIS_Cabecera

On Error GoTo ObtenerListaCodigosNaciones_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_DatosHIS_Cabecera.Conexion = oConexion
Set ObtenerListaCodigosNaciones = mo_DatosHIS_Cabecera.ListaPaisesPorCodigoNombre(Codigo, nombre)

On Error GoTo 0
Exit Function
ObtenerListaCodigosNaciones_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerListaCodigosNaciones of Módulo de clase ReglasHISGalenos"
End Function
'MODIFICADO POR YEPE NOVIEMBRE  Fin

'============================================== FORMULARIO PRINCIPAL - CARGA DE DATOS DE CABECERA ==========================================
Public Function ObtenerDatosCabecera(IdRegistroHIS As Long) As DOHIS_Cabecera
Dim oConexion As New ADODB.Connection
Dim oTablaHIS_Cabecera As New DOHIS_Cabecera
Dim mo_His_cabecera As New HIS_Cabecera

On Error GoTo ObtenerRegistroHISCabecera_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion

oTablaHIS_Cabecera.IdHisCabecera = IdRegistroHIS
mo_His_cabecera.SeleccionarPorId oTablaHIS_Cabecera
Set ObtenerDatosCabecera = oTablaHIS_Cabecera

On Error GoTo 0
Exit Function
ObtenerRegistroHISCabecera_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerRegistroHISCabecera of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMULARIO PRINCIPAL - CARGA DE DATOS DE DETALLE ===============================================
Public Function ObtenerDatosDetalleAtencion(ml_IdCabeceraHIS As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oRcs_RegistroHISDetalle As New Recordset
Dim mo_HIS_Detalle As New HIS_Detalle

On Error GoTo ObtenerRegistroHISDetalle_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Detalle.Conexion = oConexion

Set oRcs_RegistroHISDetalle = mo_HIS_Detalle.ObtenerDatosDetalleAtencion(ml_IdCabeceraHIS)
If oRcs_RegistroHISDetalle.RecordCount <> 0 Then
    oRcs_RegistroHISDetalle.MoveFirst
End If
Set ObtenerDatosDetalleAtencion = oRcs_RegistroHISDetalle

On Error GoTo 0
Exit Function
ObtenerRegistroHISDetalle_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerRegistroHISDetalle of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMULARIO PRINCIPAL - CARGA DE DATOS DE DETALLE DIAGNOSTICO ===================================
Public Function ObtenerDatosDetalleDiagnostico(ml_IdCabeceraHIS As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oRcs_RegistroHISDetalleDiagnostico As New Recordset
Dim mo_HIS_DetalleDiagnostico As New HIS_DetalleDiagnostico

On Error GoTo ObtenerRegistroHISDetalleDiagnostico_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_DetalleDiagnostico.Conexion = oConexion

Set oRcs_RegistroHISDetalleDiagnostico = mo_HIS_DetalleDiagnostico.ObtenerDatosDetalleDiagnostico(ml_IdCabeceraHIS)
If oRcs_RegistroHISDetalleDiagnostico.RecordCount <> 0 Then
    oRcs_RegistroHISDetalleDiagnostico.MoveFirst
End If
Set ObtenerDatosDetalleDiagnostico = oRcs_RegistroHISDetalleDiagnostico

On Error GoTo 0
Exit Function
ObtenerRegistroHISDetalleDiagnostico_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerRegistroHISDetalleDiagnostico of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMULARIO PRINCIPAL - CARGA DE DATOS DE DETALLE DIAGNOSTICO ===================================
Public Function ObtenerDatosDetalleDiagnosticoPorIdDetalle(ml_IdDetalleHIS As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oRcs_RegistroHISDetalleDiagnostico As New Recordset
Dim mo_HIS_DetalleDiagnostico As New HIS_DetalleDiagnostico

On Error GoTo ObtenerDatosDetalleDiagnosticoPorIdDetalle_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_DetalleDiagnostico.Conexion = oConexion

Set oRcs_RegistroHISDetalleDiagnostico = mo_HIS_DetalleDiagnostico.ObtenerDatosDetalleDiagnosticoPorIdDetalle(ml_IdDetalleHIS)
If oRcs_RegistroHISDetalleDiagnostico.RecordCount <> 0 Then
    oRcs_RegistroHISDetalleDiagnostico.MoveFirst
End If
Set ObtenerDatosDetalleDiagnosticoPorIdDetalle = oRcs_RegistroHISDetalleDiagnostico

On Error GoTo 0
Exit Function
ObtenerDatosDetalleDiagnosticoPorIdDetalle_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosDetalleDiagnosticoPorIdDetalle of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMULARIO PRINCIPAL - BUSQUEDA DE UN DISTRITO DEPENDIENDO DE SU ID ===================================
Function ConsultarDistritoPorId(moDistrito As DODistrito) As DODistrito
Dim oConexion As New ADODB.Connection
Dim mo_Distritos As New Distritos

On Error GoTo ConsultarDistritoPorId_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_Distritos.Conexion = oConexion

mo_Distritos.SeleccionarPorId moDistrito
Set ConsultarDistritoPorId = moDistrito

On Error GoTo 0
Exit Function
ConsultarDistritoPorId_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ConsultarDistritoPorId of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMULARIO PRINCIPAL - OBTENER ID DE SERVICIO POR ID DE SERVICIOESTABLECIMIENTO ===================================
Function ObtenerDatosIdServicioPorIdServEstablecimiento(IdServEstablec As Long) As Long
Dim oConexion As New ADODB.Connection
Dim mo_HIS_ServEstablecimiento As New HIS_ServEstablecimiento

On Error GoTo ObtenerDatosIdServicioPorIdServEstablecimiento_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_ServEstablecimiento.Conexion = oConexion

ObtenerDatosIdServicioPorIdServEstablecimiento = mo_HIS_ServEstablecimiento.ObtenerDatosIdServicioPorIdServEstablecimiento(IdServEstablec)

On Error GoTo 0
Exit Function
ObtenerDatosIdServicioPorIdServEstablecimiento_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosIdServicioPorIdServEstablecimiento of Módulo de clase ReglasHISGalenos"
End Function

'============================================== LISTADO FORMULARIO CODIGOS DE ACTIVIDADES ==========================================
Function ObtenerListaCodigosActividades(ml_CodigoTipoActividad As Integer) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oRcs_RObtenerListaRegistrosActividades As New Recordset
Dim mo_HIS_CodigosActividades As New HIS_CodigosActividades

On Error GoTo ObtenerListaCodigosActividades_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_CodigosActividades.Conexion = oConexion
Set oRcs_RObtenerListaRegistrosActividades = mo_HIS_CodigosActividades.ObtenerListaCodigosActividades(ml_CodigoTipoActividad)
Set ObtenerListaCodigosActividades = oRcs_RObtenerListaRegistrosActividades

On Error GoTo 0
Exit Function
ObtenerListaCodigosActividades_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerListaCodigosActividades of Módulo de clase ReglasHISGalenos"
End Function

'============================================== LISTADO FORMULARIO CODIGOS DE LAB ==========================================
Function ObtenerListaCodigosLAB() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oConexion As New ADODB.Connection
Dim mo_DatosLAB As New SIGHDatos.HIS_situacio
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
    
Set mo_DatosLAB.Conexion = oConexion
Set ObtenerListaCodigosLAB = mo_DatosLAB.ObtenerListaCodigosLAB()
Exit Function
ManejadorDeError:
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

'MODIFICADO POR YEPE NOVIEMBRE
'============================================== LISTADO FORMULARIO CODIGOS DE ACTIVIDADES POR FILTRO ==========================================
Function ObtenerListaCodigosActividadesporCodigoyNombre(ByVal CodigoActividad As String, ByVal Descripcion As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oConexion As New ADODB.Connection
Dim mo_DatosActividades As New SIGHDatos.HIS_CodigosActividades
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
    
Set mo_DatosActividades.Conexion = oConexion
Set ObtenerListaCodigosActividadesporCodigoyNombre = mo_DatosActividades.ObtenerListaCodigosActividadesporCodigoyNombre(CodigoActividad, Descripcion)
Exit Function
ManejadorDeError:
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function
'MODIFICADO POR YEPE NOVIEMBRE

'MODIFICADO POR YEPE NOVIEMBRE
'============================================== LISTADO FORMULARIO CODIGOS DE LAB POR FILTRO ==========================================
Function ObtenerListaCodigosLABporCodigoyNombre(ByVal valores As String, ByVal descripcio As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oConexion As New ADODB.Connection
Dim mo_DatosLAB As New SIGHDatos.HIS_situacio
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
    
Set mo_DatosLAB.Conexion = oConexion
Set ObtenerListaCodigosLABporCodigoyNombre = mo_DatosLAB.ObtenerListaCodigosLABporCodigoyNombre(valores, descripcio)
Exit Function
ManejadorDeError:
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function
'MODIFICADO POR YEPE NOVIEMBRE

'MODIFICADO POR YEPE NOVIEMBRE
'============================================== LISTADO FORMULARIO CODIGOS DE LAB POR FILTRO ==========================================
Function ObtenerListaCodigosProductosHisPorNombreYDescripcion(ByVal codigodiagcpt As String, ByVal descripciondiagcpt As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim mo_DatosProductosHis As New SIGHDatos.HIS_Cabecera
    Dim lnPosSeparador As Integer
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
        
    lnPosSeparador = InStr(codigodiagcpt, "-")
    If lnPosSeparador = 0 Then
    Else
        codigodiagcpt = Trim(Mid(codigodiagcpt, 2, lnPosSeparador - 4))
    End If
    
    Set mo_DatosProductosHis.Conexion = oConexion
    Set ObtenerListaCodigosProductosHisPorNombreYDescripcion = mo_DatosProductosHis.ObtenerListaCodigosProductosHisPorCodigoyNombre(codigodiagcpt, descripciondiagcpt)
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

'============================================== LISTADO FORMULARIO SERVICIOS  - LISTAR ESTABLECIMIENTOS MR ==========================================
Function ObtenerListaEstablecimientosMR() As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_DatosServiciosPorEstablecimientos As New SIGHDatos.HIS_ServEstablecimiento
    
    On Error GoTo ObtenerListaEstablecimientosMR_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
        
    Set mo_DatosServiciosPorEstablecimientos.Conexion = oConexion
    Set ObtenerListaEstablecimientosMR = mo_DatosServiciosPorEstablecimientos.ObtenerListaEstablecimientosMR
    
    On Error GoTo 0
    Exit Function
ObtenerListaEstablecimientosMR_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerListaEstablecimientosMR of Módulo de clase ReglasHISGalenos"
End Function

'============================================== LISTADO FORMULARIO SERVICIOS  - LISTAR SERVICIOS CENTRO MR ==========================================
Public Function ObtenerListaServiciosCentroMR() As Recordset
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_DatosServiciosPorEstablecimientos As New SIGHDatos.HIS_ServEstablecimiento
    
    On Error GoTo ListarServiciosEstablecimientosMR_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set mo_DatosServiciosPorEstablecimientos.Conexion = oConexion
    Set oRcs_Result = mo_DatosServiciosPorEstablecimientos.ObtenerListaServiciosCentroMR
    Set ObtenerListaServiciosCentroMR = oRcs_Result
    
    On Error GoTo 0
    Exit Function
    
ListarServiciosEstablecimientosMR_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListarServiciosEstablecimientosMR of Módulo de clase ReglasHISGalenos"
End Function

'============================================== LISTADO FORMULARIO SERVICIOS  - LISTAR SERVICIOS POR ESTABLECIMIENTO MR ==========================================
Public Function ObtenerListaServiciosPorEstablecimiento(IdEstablecimiento As Long) As Recordset
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_DatosServiciosPorEstablecimiento As New SIGHDatos.HIS_ServEstablecimiento
    
    On Error GoTo ObtenerListaServiciosPorEstablecimiento_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set mo_DatosServiciosPorEstablecimiento.Conexion = oConexion
    Set oRcs_Result = mo_DatosServiciosPorEstablecimiento.ListaServiciosPorEstablecimiento(IdEstablecimiento)
    Set ObtenerListaServiciosPorEstablecimiento = oRcs_Result
    
    On Error GoTo 0
    Exit Function
    
ObtenerListaServiciosPorEstablecimiento_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListarServiciosEstablecimiento of Módulo de clase ReglasHISGalenos"
End Function

'============================================== LISTADO FORMULARIO SERVICIOS - MANTENIMIENTO DE SERVICIOS DE LA MR ==========================================
Public Function ActualizarServiciosPorEstablecimientos(lnIdEstablecimiento As Long, sighOpcion As sghOpciones, oRcs_ServiciosPorEstablecimiento As Recordset) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oTabla_DatosServiciosPorEstablecimientos As New DOHIS_ServEstablecimiento
    Dim mo_DatosServiciosPorEstablecimientos As New SIGHDatos.HIS_ServEstablecimiento
    Dim oDOHis_Establecimientos As New DOHis_Establecimientos
    Dim oHis_Establecimientos As New SIGHDatos.His_Establecimientos
    Dim mb_Resultado As Boolean
    Dim sSql As String
    ms_MensajeError = ""
    mb_Resultado = True
    On Error GoTo ActualizarServiciosPorEstablecimientos_Error
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oHis_Establecimientos.Conexion = oConexion
    Set mo_DatosServiciosPorEstablecimientos.Conexion = oConexion
    oConexion.BeginTrans
    
    Select Case sighOpcion
        Case sghOpciones.sghAgregar
            oDOHis_Establecimientos.IdEstablecimiento = lnIdEstablecimiento
            oDOHis_Establecimientos.IdUsuarioAuditoria = 0
            If oHis_Establecimientos.Insertar(oDOHis_Establecimientos) Then
                mb_Resultado = True
            Else
                mb_Resultado = False
            End If
        Case sghOpciones.sghModificar
'            oTabla_DatosServiciosPorEstablecimientos.idestablecimiento = lnidestablecimiento
'            If mo_DatosServiciosPorEstablecimientos.Eliminar(oTabla_DatosServiciosPorEstablecimientos) Then
'                mb_Resultado = True
'            Else
'                mb_Resultado = False
'            End If
        Case sghOpciones.sghEliminar
            oDOHis_Establecimientos.IdEstablecimiento = lnIdEstablecimiento
            oDOHis_Establecimientos.IdUsuarioAuditoria = 0
            If oHis_Establecimientos.Eliminar(oDOHis_Establecimientos) Then
                mb_Resultado = True
            Else
                mb_Resultado = False
            End If
     End Select

    If mb_Resultado = True And sighOpcion <> sghOpciones.sghEliminar Then
        oRcs_ServiciosPorEstablecimiento.Filter = "IdEstado<>0"
        If oRcs_ServiciosPorEstablecimiento.RecordCount <> 0 Then
            oRcs_ServiciosPorEstablecimiento.MoveFirst
            Do While Not oRcs_ServiciosPorEstablecimiento.EOF
                Select Case oRcs_ServiciosPorEstablecimiento!IdEstado
                    Case sghOpciones.sghAgregar
                        oTabla_DatosServiciosPorEstablecimientos.IdServicio = oRcs_ServiciosPorEstablecimiento!IdServicio
                        oTabla_DatosServiciosPorEstablecimientos.IdEstablecimiento = oRcs_ServiciosPorEstablecimiento!IdEstablecimiento
                        If mo_DatosServiciosPorEstablecimientos.Insertar(oTabla_DatosServiciosPorEstablecimientos) Then
                            mb_Resultado = True
                        Else
                            mb_Resultado = False
                            Exit Do
                        End If
                    Case sghOpciones.sghModificar
                        oTabla_DatosServiciosPorEstablecimientos.IdServicio = oRcs_ServiciosPorEstablecimiento!IdServicio
                        oTabla_DatosServiciosPorEstablecimientos.IdEstablecimiento = oRcs_ServiciosPorEstablecimiento!IdEstablecimiento
                        If mo_DatosServiciosPorEstablecimientos.modificar(oTabla_DatosServiciosPorEstablecimientos) Then
                            mb_Resultado = True
                        Else
                            mb_Resultado = False
                            Exit Do
                        End If
                        
                    Case (sghOpciones.sghEliminar - 1)
                        oTabla_DatosServiciosPorEstablecimientos.IdServicio = oRcs_ServiciosPorEstablecimiento!IdServicio
                        oTabla_DatosServiciosPorEstablecimientos.IdEstablecimiento = oRcs_ServiciosPorEstablecimiento!IdEstablecimiento
                        If mo_DatosServiciosPorEstablecimientos.Eliminar(oTabla_DatosServiciosPorEstablecimientos) Then
                            mb_Resultado = True
                        Else
                            mb_Resultado = False
                            Exit Do
                        End If
                End Select
                oRcs_ServiciosPorEstablecimiento.MoveNext
            Loop
        End If
    End If
        
    If mb_Resultado Then
        oConexion.CommitTrans
        oConexion.Close
        ActualizarServiciosPorEstablecimientos = True
    Else
        oConexion.RollbackTrans
        oConexion.Close
        ActualizarServiciosPorEstablecimientos = False
        ms_MensajeError = mo_DatosServiciosPorEstablecimientos.MensajeError
    End If
    Set oConexion = Nothing
    Set oTabla_DatosServiciosPorEstablecimientos = Nothing
    Set mo_DatosServiciosPorEstablecimientos = Nothing
    
    On Error GoTo 0
    Exit Function
ActualizarServiciosPorEstablecimientos_Error:
    'oConexion.Close
    MsgBox "Error " & Err.Number & " (" & ms_MensajeError & ") in procedure ActualizarServiciosPorEstablecimientos of Módulo de clase ReglasHISGalenos"
End Function


'============================================== LISTADO FORMULARIO SERVICIOS - LISTADO DE SERVICIOS POR DEP Y ESP ==========================================
Public Function ConsultarRegistroServiciosPorEspec(mo_cmbEspecialidad As Long) As Recordset
Dim mo_Servicios As New Servicios
Dim oConexion As New ADODB.Connection

On Error GoTo ConsultarRegistroServiciosPorDepYEsp_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_Servicios.Conexion = oConexion

Set ConsultarRegistroServiciosPorEspec = mo_Servicios.ServiciosSeleccionarConsultoriosPorEspecialidadDEBB(mo_cmbEspecialidad, sghFiltraSoloActivos)

On Error GoTo 0
Exit Function
ConsultarRegistroServiciosPorDepYEsp_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ConsultarRegistroServiciosPorDepYEsp of Módulo de clase ReglasHISGalenos"
End Function

'============================================== PROGRAMACION MEDICA - MANTENIMIENTO INGRESAR CRONOGRAMA ==========================================
Public Function IngresarRegistroProgramacionMedica(oTablaProgramacionMed As DOHIS_ProgMedEstMR, mo_FechasProgramadas As Collection) As Boolean
Dim mo_HIS_ProgMedEstMR As New HIS_ProgMedEstMR
Dim oConexion As New ADODB.Connection
Dim dia As Variant
On Error GoTo IngresarRegistroProgramacionMedica_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_ProgMedEstMR.Conexion = oConexion

oConexion.BeginTrans
For Each dia In mo_FechasProgramadas
    oTablaProgramacionMed.FechaProgramada = dia
    If Not mo_HIS_ProgMedEstMR.Insertar(oTablaProgramacionMed) Then
        Err.Description = mo_HIS_ProgMedEstMR.MensajeError
        GoTo IngresarRegistroProgramacionMedica_Error
    End If
Next
oConexion.CommitTrans
oConexion.Close
Set oConexion = Nothing
Set mo_HIS_ProgMedEstMR = Nothing
IngresarRegistroProgramacionMedica = True
On Error GoTo 0
Exit Function
IngresarRegistroProgramacionMedica_Error:
oConexion.RollbackTrans
oConexion.Close
Set oConexion = Nothing
Set mo_HIS_ProgMedEstMR = Nothing
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IngresarRegistroProgramacionMedica of Módulo de clase ReglasHISGalenos"
End Function

Public Function ModificarRegistroProgramacionMedica(oTablaProgramacionMed As DOHIS_ProgMedEstMR, mo_FechasProgramadas As Collection) As Boolean
Dim mo_HIS_ProgMedEstMR As New HIS_ProgMedEstMR
Dim oConexion As New ADODB.Connection
Dim dia As Variant
On Error GoTo ModificarRegistroProgramacionMedica_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_ProgMedEstMR.Conexion = oConexion

oConexion.BeginTrans
For Each dia In mo_FechasProgramadas
    oTablaProgramacionMed.FechaProgramada = dia
    If Not mo_HIS_ProgMedEstMR.modificar(oTablaProgramacionMed) Then
        Err.Description = mo_HIS_ProgMedEstMR.MensajeError
        GoTo ModificarRegistroProgramacionMedica_Error
    End If
Next
oConexion.CommitTrans
oConexion.Close
Set oConexion = Nothing
Set mo_HIS_ProgMedEstMR = Nothing
ModificarRegistroProgramacionMedica = True
On Error GoTo 0
Exit Function
ModificarRegistroProgramacionMedica_Error:
oConexion.RollbackTrans
oConexion.Close
Set oConexion = Nothing
Set mo_HIS_ProgMedEstMR = Nothing
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ModificarRegistroProgramacionMedica of Módulo de clase ReglasHISGalenos"
End Function


'Frank HIS
Function ConsultarProgramacionMedicaHIS(DOHIS_ProgMedEstMR As DOHIS_ProgMedEstMR) As Boolean
Dim oConexion As New ADODB.Connection
Dim mo_HIS_ProgMedEstMR As New HIS_ProgMedEstMR
ConsultarProgramacionMedicaHIS = False

On Error GoTo ConsultarProgramacionMedicaHIS_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_ProgMedEstMR.Conexion = oConexion
    
ConsultarProgramacionMedicaHIS = mo_HIS_ProgMedEstMR.SeleccionarPorId(DOHIS_ProgMedEstMR)
'Set ConsultarProgramacionMedicaHIS = DOHIS_ProgMedEstMR
On Error GoTo 0
Exit Function
ConsultarProgramacionMedicaHIS_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ConsultarProgramacionMedicaHIS of Módulo de clase ReglasHISGalenos"
End Function

'============================================== PROGRAMACION MEDICA - MANTENIMIENTO ELIMINAR CRONOGRAMA ==========================================
Public Function EliminarRegistroProgramacionMedica(oTablaHIS_ProgMedEstMR As DOHIS_ProgMedEstMR) As Boolean
Dim mo_HIS_ProgMedEstMR As New HIS_ProgMedEstMR
Dim oConexion As New ADODB.Connection
On Error GoTo EliminarRegistroProgramacionMedica_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_ProgMedEstMR.Conexion = oConexion

EliminarRegistroProgramacionMedica = mo_HIS_ProgMedEstMR.Eliminar(oTablaHIS_ProgMedEstMR)
oConexion.Close
Set oConexion = Nothing
On Error GoTo 0
Exit Function
EliminarRegistroProgramacionMedica_Error:
Set oConexion = Nothing
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EliminarRegistroProgramacionMedica of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMULARIO PRINCIPAL - VALIDACION FECHAS MES DE MEDICO PROGRAMADO ==========================================
Function ValidarProgramacionMedica_FechasMesActual(IdMedico As Long, fecha As String) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_ProgMedEstMR As New HIS_ProgMedEstMR

On Error GoTo ValidarProgramacionMedica_FechasMesActual_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_ProgMedEstMR.Conexion = oConexion

Set oRcs_Result = mo_HIS_ProgMedEstMR.ListarProgramacionMedica_FechasMesActual(IdMedico, fecha)

If oRcs_Result.RecordCount <> 0 Or oRcs_Result.RecordCount = -1 Then
    ValidarProgramacionMedica_FechasMesActual = True
Else
    ValidarProgramacionMedica_FechasMesActual = False
End If

On Error GoTo 0
Exit Function
ValidarProgramacionMedica_FechasMesActual_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidarProgramacionMedica_FechasMesActual of Módulo de clase ReglasHISGalenos"
End Function

'============================================== PROGRAMACION MEDICA - OBTIEN LOS DATOS DE LA PROGRAMACION MEDICA DE UN RESPONSABLE EN MES ACTUAL ==========================================
Function ListarProgramacionMedica_FechasMesActual(IdMedico As Long, ms_Fecha As String) As Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_ProgMedEstMR As New HIS_ProgMedEstMR

On Error GoTo ListarProgramacionMedica_FechasMesActual_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_ProgMedEstMR.Conexion = oConexion

Set ListarProgramacionMedica_FechasMesActual = mo_HIS_ProgMedEstMR.ListarProgramacionMedica_FechasMesActual(IdMedico, ms_Fecha)

On Error GoTo 0
Exit Function
ListarProgramacionMedica_FechasMesActual_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListarProgramacionMedica_FechasMesActual of Módulo de clase ReglasHISGalenos"
End Function
 
'============================================== FORMUALRIO PRINCIPAL - OBTIENE PROGRAMACION DE UN MEDICO EN UN ESTABLECIMIENTO DICHO ==========================================
Function ListarProgramacionMedicaPorMedicoYEstablecimiento(ml_IdEstablecimientoActual As Long, IdMedico As Long, IdMes As Integer, IdAnio As Integer) As Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_ProgMedEstMR As New HIS_ProgMedEstMR

On Error GoTo ListarProgramacionMedicaPorMedicoYEstablecimiento_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_ProgMedEstMR.Conexion = oConexion

Set ListarProgramacionMedicaPorMedicoYEstablecimiento = mo_HIS_ProgMedEstMR.ListarProgramacionMedicaPorMedicoYEstablecimiento(ml_IdEstablecimientoActual, IdMedico, IdMes, IdAnio)

On Error GoTo 0
Exit Function

ListarProgramacionMedicaPorMedicoYEstablecimiento_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListarProgramacionMedicaPorMedicoYEstablecimiento of Módulo de clase ReglasHISGalenos"
End Function
 
'============================================== LOTE HIS - INGRESO DE LOTE HIS ==========================================
Function IngresarRegistroLoteHIS(oTablaDOHIS_Lote As DOHIS_Lotes) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo IngresarRegistroLoteHIS_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion
    
IngresarRegistroLoteHIS = mo_HIS_Lotes.Insertar(oTablaDOHIS_Lote)

oConexion.Close
Set oRcs_Result = Nothing
Set oConexion = Nothing
Set mo_HIS_Lotes = Nothing

On Error GoTo 0
Exit Function
IngresarRegistroLoteHIS_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IngresarRegistroLoteHIS of Módulo de clase ReglasHISGalenos"
End Function

'============================================== LOTE HIS - ACTUALIZACION DE LOTE HIS ==========================================
Function ModificarRegistroLoteHIS(oTablaDOHIS_Lote As DOHIS_Lotes) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo ModificarRegistroLoteHIS_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion
    
ModificarRegistroLoteHIS = mo_HIS_Lotes.modificar(oTablaDOHIS_Lote)

oConexion.Close
Set oRcs_Result = Nothing
Set oConexion = Nothing
Set mo_HIS_Lotes = Nothing

On Error GoTo 0
Exit Function
ModificarRegistroLoteHIS_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ModificarRegistroLoteHIS of Módulo de clase ReglasHISGalenos"
End Function

'============================================== LOTE HIS - ELIMINAR DE LOTE HIS ==========================================
Function EliminarRegistroLoteHIS(oTablaDOHIS_Lote As DOHIS_Lotes) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo EliminarRegistroLoteHIS_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion
    
EliminarRegistroLoteHIS = mo_HIS_Lotes.Eliminar(oTablaDOHIS_Lote)

oConexion.Close
Set oRcs_Result = Nothing
Set oConexion = Nothing
Set mo_HIS_Lotes = Nothing

On Error GoTo 0
Exit Function
EliminarRegistroLoteHIS_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EliminarRegistroLoteHIS of Módulo de clase ReglasHISGalenos"
End Function

'============================================== LOTE HIS - CONSULTAR DE LOTE HIS ==========================================
Function ConsultarRegistroLoteHIS(oTablaDOHIS_Lote As DOHIS_Lotes) As DOHIS_Lotes
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo ConsultarRegistroLoteHIS_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion
    
mo_HIS_Lotes.SeleccionarPorId oTablaDOHIS_Lote
Set ConsultarRegistroLoteHIS = oTablaDOHIS_Lote

On Error GoTo 0
Exit Function
ConsultarRegistroLoteHIS_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ConsultarRegistroLoteHIS of Módulo de clase ReglasHISGalenos"
End Function

Function ModificarEstadoLoteHIS(oTablaDOHIS_Lote As DOHIS_Lotes) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo ModificarRegistroLoteHIS_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion
    
ModificarEstadoLoteHIS = mo_HIS_Lotes.ModificarEstadoLote(oTablaDOHIS_Lote)

oConexion.Close
Set oRcs_Result = Nothing
Set oConexion = Nothing
Set mo_HIS_Lotes = Nothing

On Error GoTo 0
Exit Function
ModificarRegistroLoteHIS_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ModificarEstadoLoteHIS of Módulo de clase ReglasHISGalenos"
End Function

Function ConsultarRegistrosLotesPorIdEstablec(IdEstablecimiento As Long) As ADODB.Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo ConsultarRegistrosLotesHIS_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion

Set ConsultarRegistrosLotesPorIdEstablec = mo_HIS_Lotes.ConsultarRegistrosLotesPorIdEstablec(IdEstablecimiento)

On Error GoTo 0
Exit Function
ConsultarRegistrosLotesHIS_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ConsultarRegistrosLotesHIS of Módulo de clase ReglasHISGalenos"
End Function

'============================================== LOTE HIS - VALIDAR SI EL LOTE EXISTE ==========================================
Function ValidarLoteHIS_LoteExiste(oTablaDOHIS_Lote As DOHIS_Lotes) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo ValidarLoteHIS_LoteExiste_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion

ValidarLoteHIS_LoteExiste = mo_HIS_Lotes.ValidarLoteHIS_LoteExiste(oTablaDOHIS_Lote)

On Error GoTo 0
Exit Function
ValidarLoteHIS_LoteExiste_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidarLoteHIS_LoteExiste of Módulo de clase ReglasHISGalenos"
End Function

'============================================== LOTE HIS - VALIDAR SI NO SE CERRO EL ULTIMO LOTE ==========================================
Function ValidarLoteHIS_UltimoLoteNoCerrado(oTablaDOHIS_Lote As DOHIS_Lotes) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo ValidarLoteHIS_UltimoLoteNoCerrado_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion
    
ValidarLoteHIS_UltimoLoteNoCerrado = mo_HIS_Lotes.ValidarLoteHIS_UltimoLoteNoCerrado(oTablaDOHIS_Lote)

On Error GoTo 0
Exit Function
ValidarLoteHIS_UltimoLoteNoCerrado_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidarLoteHIS_UltimoLoteNoCerrado of Módulo de clase ReglasHISGalenos"
End Function

'============================================== MANTENIMIENTO PRINCIPAL - CONSULTAR DEPARTAMENTO POR USUARIO ==========================================
Function ObtenerDatosEstablecimientoPorUsuario(ml_idUsuario As Long) As ADODB.Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_His_cabecera As New HIS_Cabecera

On Error GoTo ObtenerIdDepPorUsuario_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion
    
Set ObtenerDatosEstablecimientoPorUsuario = mo_His_cabecera.ObtenerDatosEstablecimientoPorUsuario(ml_idUsuario)

On Error GoTo 0
Exit Function
ObtenerIdDepPorUsuario_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosEstablecimientoPorUsuario of Módulo de clase ReglasHISGalenos"
End Function

'============================================== MANTENIMIENTO PRINCIPAL - VALIDAR LOTE HIS ULTIMO CERRADO ==========================================
Function ValidarLoteHIS_UltimoLoteCerrado(ml_IdEstablecimientoActual As Long) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lote As New HIS_Lotes

On Error GoTo ValidarLoteHIS_UltimoLoteCerrado_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lote.Conexion = oConexion

ValidarLoteHIS_UltimoLoteCerrado = mo_HIS_Lote.ValidarLoteHIS_UltimoLoteCerrado(ml_IdEstablecimientoActual)

On Error GoTo 0
Exit Function
ValidarLoteHIS_UltimoLoteCerrado_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidarLoteHIS_UltimoLoteCerrado of Módulo de clase ReglasHISGalenos"
End Function

'============================================== MANTENIMIENTO PRINCIPAL - VALIDAR LOTE HIS PSAO EL LIMITE DE PAGINAS ==========================================
Function ValidarLoteHIS_PasoMaximoPaginasPermitido(ml_IdEstablecimientoActual As Long) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lote As New HIS_Lotes

On Error GoTo ValidarLoteHIS_PasoMaximoPaginasPermitido_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lote.Conexion = oConexion
ValidarLoteHIS_PasoMaximoPaginasPermitido = mo_HIS_Lote.ValidarLoteHIS_PasoMaximoPaginasPermitido(ml_IdEstablecimientoActual)

On Error GoTo 0
Exit Function
ValidarLoteHIS_PasoMaximoPaginasPermitido_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ValidarLoteHIS_PasoMaximoPaginasPermitido of Módulo de clase ReglasHISGalenos"
End Function

'============================================== MANTENIMIENTO PRINCIPAL - CONSULTAR DATOS LOTE HIS POR USUARIO ==========================================
Function ObtenerDatosLotePorEstablecimiento(ml_IdEstablecimientoActual As Long) As ADODB.Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lote As New HIS_Lotes

On Error GoTo ObtenerDatosLotePorUsuario_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lote.Conexion = oConexion
Set ObtenerDatosLotePorEstablecimiento = mo_HIS_Lote.ObtenerDatosLotePorEstablecimiento(ml_IdEstablecimientoActual)

On Error GoTo 0
Exit Function
ObtenerDatosLotePorUsuario_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosLotePorUsuario of Módulo de clase ReglasHISGalenos"
End Function

'============================================== MANTENIMIENTO PRINCIPAL - CONSULTAR DATOS LOTE HIS POR ID ==========================================
Function ObtenerDatosLotePorIdLote(DoHisLote As DOHIS_Lotes) As DOHIS_Lotes
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lote As New HIS_Lotes

On Error GoTo ObtenerDatosLotePorIdLote_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lote.Conexion = oConexion

mo_HIS_Lote.SeleccionarPorId DoHisLote

Set ObtenerDatosLotePorIdLote = DoHisLote

On Error GoTo 0
Exit Function
ObtenerDatosLotePorIdLote_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosLotePorIdLote of Módulo de clase ReglasHISGalenos"
End Function

'============================================== MANTENIMIENTO PRINCIPAL - CONSULTAR DATOS DIGITADOR HIS ==========================================
Function ObtenerDatosDigitador(ml_idUsuario As Long) As ADODB.Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_His_cabecera As New HIS_Cabecera

On Error GoTo ObtenerDatosDigitador_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion

Set ObtenerDatosDigitador = mo_His_cabecera.ObtenerDatosDigitador(ml_idUsuario)

On Error GoTo 0
Exit Function
ObtenerDatosDigitador_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosDigitador of Módulo de clase ReglasHISGalenos"
End Function

'============================================== MANTENIMIENTO PRINCIPAL - CONSULTAR DATO HOJA LIBRE ==========================================
Function ObtenerDatosLoteNroHojaLibre(ml_IdLote As Long) As Integer
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lote As New HIS_Lotes

On Error GoTo ObtenerDatosLoteHojaLibre_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lote.Conexion = oConexion

ObtenerDatosLoteNroHojaLibre = mo_HIS_Lote.ObtenerDatosLoteNroHojaLibre(ml_IdLote)

On Error GoTo 0
Exit Function
ObtenerDatosLoteHojaLibre_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosLoteHojaLibre of Módulo de clase ReglasHISGalenos"
End Function

'Function ObtenerListaHojasDisponibles(ml_IdLote As Long) As ADODB.Recordset
'Dim oRcs_Result As New Recordset
'Dim oConexion As New ADODB.Connection
'Dim mo_HIS_Lote As New HIS_Lotes
'
'On Error GoTo ObtenerDatosLoteHojaLibre_Error
'oConexion.Open sighentidades.CadenaConexion
'oConexion.CursorLocation = adUseClient
'Set mo_HIS_Lote.Conexion = oConexion
'
'Set ObtenerListaHojasDisponibles = mo_HIS_Lote.ObtenerListaHojasDisponibles(ml_IdLote)
'
'On Error GoTo 0
'Exit Function
'ObtenerDatosLoteHojaLibre_Error:
'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerListaHojasDisponibles of Módulo de clase ReglasHISGalenos"
'End Function

'============================================== MANTENIMIENTO PRINCIPAL - LISTA DE SERVICIOS POR ESTABLECIMIENTO ==========================================
Function ListaServiciosPorEstablecimiento(ml_IdEstablecimientoActual As Long) As Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_ServEstablecimiento As New HIS_ServEstablecimiento

On Error GoTo ListaServiciosPorEstablecimiento_Error
oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open sighentidades.CadenaConexion

Set mo_HIS_ServEstablecimiento.Conexion = oConexion

Set ListaServiciosPorEstablecimiento = mo_HIS_ServEstablecimiento.ListaServiciosPorEstablecimiento(ml_IdEstablecimientoActual)

On Error GoTo 0
Exit Function
ListaServiciosPorEstablecimiento_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaServiciosPorEstablecimiento of Módulo de clase ReglasHISGalenos"
End Function

'============================================== MANTENIMIENTO PRINCIPAL - OBTIENE EL NRO DE FORMATO LIBRE ==========================================
Function ObtenerDatosNroFormatoLibre(mi_anio As Integer, ml_IdEstablecimientoActual As Long) As String
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_His_cabecera As New HIS_Cabecera

On Error GoTo ObtenerDatosNroFormatoLibre_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion

ObtenerDatosNroFormatoLibre = mo_His_cabecera.ObtenerDatosNroFormatoLibre(mi_anio, ml_IdEstablecimientoActual)
    
On Error GoTo 0
Exit Function
ObtenerDatosNroFormatoLibre_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosNroFormatoLibre of Módulo de clase ReglasHISGalenos"
End Function

'============================================== MANTENIMIENTO PRINCIPAL - OBTIENE EL ID DE SERVICIO ESCOGIDO ==========================================
'Revisar Frank
Public Function ObtenerDatosIdServEstablecimiento(IdServicio As Long, IdEstablecimiento As Long) As Long
    Dim oConexion As New ADODB.Connection
    Dim oTablaHIS_ServiciosPorEstablec As New DOHIS_ServEstablecimiento
    Dim mo_DatosServiciosPorEstablecimientos As New SIGHDatos.HIS_ServEstablecimiento
    
    On Error GoTo ObtenerDatosIdServEstablecimiento_Error
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oTablaHIS_ServiciosPorEstablec.IdEstablecimiento = IdEstablecimiento
    oTablaHIS_ServiciosPorEstablec.IdServicio = IdServicio
    Set mo_DatosServiciosPorEstablecimientos.Conexion = oConexion
    
    mo_DatosServiciosPorEstablecimientos.ObtenerDatosIdServEstablecimiento oTablaHIS_ServiciosPorEstablec
    ObtenerDatosIdServEstablecimiento = oTablaHIS_ServiciosPorEstablec.IdServicio
    
    On Error GoTo 0
    Exit Function
ObtenerDatosIdServEstablecimiento_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosIdServEstablecimiento of Módulo de clase ReglasHISGalenos"
End Function

'======================================================================================================================================================
'======================================================================================================================================================
Function IngresarHojaHIS(RegCabeceraHIS As DOHIS_Cabecera) As Integer
Dim mi_CodigoCabeceraHIS As Integer
Dim oConexion As New ADODB.Connection
Dim oCabeceraHIS As New HIS_Cabecera
Dim sSql As String

mi_CodigoCabeceraHIS = 0

On Error GoTo ManejadorDeError
ms_MensajeError = ""

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set oCabeceraHIS.Conexion = oConexion

If oCabeceraHIS.Insertar(RegCabeceraHIS) Then
    'Despues de grabar obten el id de cabecera
    mi_CodigoCabeceraHIS = RegCabeceraHIS.IdHisCabecera
Else
    GoTo ManejadorDeError:
End If

IngresarHojaHIS = mi_CodigoCabeceraHIS

Exit Function
ManejadorDeError:
oConexion.Close
Set oConexion = Nothing
IngresarHojaHIS = False
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

Function ModificarHojaHIS(RegCabeceraHIS As DOHIS_Cabecera) As Integer
Dim mi_CodigoCabeceraHIS As Integer
Dim oConexion As New ADODB.Connection
Dim oCabeceraHIS As New HIS_Cabecera
Dim sSql As String

mi_CodigoCabeceraHIS = 0

On Error GoTo ManejadorDeError
ms_MensajeError = ""

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set oCabeceraHIS.Conexion = oConexion

If oCabeceraHIS.modificar(RegCabeceraHIS) Then
    'Despues de grabar obten el id de cabecera
    mi_CodigoCabeceraHIS = RegCabeceraHIS.IdHisCabecera
Else
    GoTo ManejadorDeError:
End If

ModificarHojaHIS = mi_CodigoCabeceraHIS

Exit Function
ManejadorDeError:
oConexion.Close
Set oConexion = Nothing
ModificarHojaHIS = False
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

Function EliminarHojaHIS(RegCabeceraHIS As DOHIS_Cabecera) As Boolean
Dim oConexion As New ADODB.Connection
Dim mo_CabeceraHIS As New HIS_Cabecera
Dim mo_HIS_Detalle As New HIS_Detalle
Dim mo_HIS_DetalleDiagnosticos As New HIS_DetalleDiagnostico
Dim mo_HIS_Lote As New HIS_Lotes

Dim oRcs_ListaIdDiagnosticos As New ADODB.Recordset
Dim oRcs_ListaIdAtenciones As New ADODB.Recordset
Dim sSql As String

On Error GoTo ManejadorDeError
ms_MensajeError = ""

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient

Set mo_CabeceraHIS.Conexion = oConexion
Set mo_HIS_Detalle.Conexion = oConexion
Set mo_HIS_DetalleDiagnosticos.Conexion = oConexion
Set mo_HIS_Lote.Conexion = oConexion

oConexion.BeginTrans

'------- OBTENCION DE IDS DE ATENCIONES DE CABECERA ----------------
Set oRcs_ListaIdAtenciones = ListaIdsAtencionesPorIdCabecera(RegCabeceraHIS)

'------- OBTENCION DE IDS DE DIAGNOSTICOS DE CABECERA --------------
Set oRcs_ListaIdDiagnosticos = ListaIdsDiagnosticosPorIdCabecera(RegCabeceraHIS)

'------- ELIMINACION DE DIAGNOSTICOS -------------------------------
Dim oTablaDOHIS_DetalleDiagnostico As New DOHIS_DetalleDiagnostico
If oRcs_ListaIdDiagnosticos.RecordCount <> 0 Then
    oRcs_ListaIdDiagnosticos.MoveFirst
    Do While Not oRcs_ListaIdDiagnosticos.EOF
        oTablaDOHIS_DetalleDiagnostico.idhisdetallediagnostico = oRcs_ListaIdDiagnosticos.Fields!idhisdetallediagnostico
        mo_HIS_DetalleDiagnosticos.Eliminar oRcs_ListaIdDiagnosticos.Fields!IdHisDetalle
        Exit Do
        oRcs_ListaIdDiagnosticos.MoveNext
    Loop
End If
'------- ELIMINACION DE ATENCIONES ---------------------------------
If oRcs_ListaIdAtenciones.RecordCount <> 0 Then
    Dim oTablaDOHIS_Detalle As New DOHIS_Detalle
    oRcs_ListaIdAtenciones.MoveFirst
    
    Do While Not oRcs_ListaIdAtenciones.EOF
        oTablaDOHIS_Detalle.IdHisDetalle = oRcs_ListaIdAtenciones!IdHisDetalle
        mo_HIS_Detalle.Eliminar oTablaDOHIS_Detalle
        oRcs_ListaIdAtenciones.MoveNext
    Loop
End If
'------- ELIMINACION DE CABECERA -----------------------------------
mo_CabeceraHIS.Eliminar RegCabeceraHIS

'------- ACTUALIZACION DE NUMERO DE PAGINAS EN EL LOTE -------------
'mo_HIS_Lote.DisminuirNroHojaHIS RegCabeceraHIS.IdHisLote

oConexion.CommitTrans
oConexion.Close
Set oConexion = Nothing
Set mo_CabeceraHIS = Nothing
Set mo_HIS_Detalle = Nothing
Set mo_HIS_DetalleDiagnosticos = Nothing
Set mo_HIS_Lote = Nothing
Set oRcs_ListaIdDiagnosticos = Nothing
Set oRcs_ListaIdAtenciones = Nothing



EliminarHojaHIS = True
Exit Function

ManejadorDeError:
oConexion.RollbackTrans
oConexion.Close
Set oConexion = Nothing
Set mo_CabeceraHIS = Nothing
Set mo_HIS_Detalle = Nothing
Set mo_HIS_DetalleDiagnosticos = Nothing
Set mo_HIS_Lote = Nothing
Set oRcs_ListaIdDiagnosticos = Nothing
Set oRcs_ListaIdAtenciones = Nothing

ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

'======================================================================================================================================================
'======================================================================================================================================================
Function IngresarRegistroHIS(mi_CodigoCabeceraHIS As Integer, idUsuario As Long, ListaAtenciones As ADODB.Recordset, ListaDiagnosticos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, ml_idEstablecimiento As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTablaHIS_Detalle As New DOHIS_Detalle
Dim oTablaHIS_DetalleDiagnostico As New DOHIS_DetalleDiagnostico

Dim oHIS_Detalle As New HIS_Detalle
Dim oHIS_DetalleDiagnosticos As New HIS_DetalleDiagnostico

Dim bProcesoOK As Boolean
Dim sSql As String

On Error GoTo ManejadorDeError
ms_MensajeError = ""
bProcesoOK = True

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient

Set oHIS_Detalle.Conexion = oConexion
Set oHIS_DetalleDiagnosticos.Conexion = oConexion

oConexion.BeginTrans
' ================================== REGISTRO DE ATENCION ==============================
ListaAtenciones.MoveFirst
oHIS_Detalle.SetDefaults oTablaHIS_Detalle

'Datos comunes
oTablaHIS_Detalle.IdHisCabecera = mi_CodigoCabeceraHIS 'Se toma del dato externo.
oTablaHIS_Detalle.IdTipoAtencion = ListaAtenciones!IdTipoAtencion
oTablaHIS_Detalle.DiaAtencion = ListaAtenciones!DiaAtencion
oTablaHIS_Detalle.NroRegistroHoja = ListaAtenciones!NroRegistroHoja
oTablaHIS_Detalle.NroRegistroLote = 0
oTablaHIS_Detalle.IdDistrito = ListaAtenciones!IdDistrito
oTablaHIS_Detalle.IdUsuarioAuditoria = idUsuario

'Analizamos el tipo de entrada de los regsitro dependiendo del tipo de atencion
Select Case CInt(ListaAtenciones!IdTipoAtencion)

Case sghHISTipoActividad.Atencion
    'Verificacion de ID de Paciente
    oTablaHIS_Detalle.IdHisPaciente = ProcesarPaciente(CLng(ListaAtenciones!IdPacienteGalenHos), _
                                                       CInt(ListaAtenciones!IdNacionalidad), _
                                                       CLng(ListaAtenciones!IdTipoDocIdentidad), _
                                                       CStr(ListaAtenciones!NroDocIdentidad), _
                                                       Trim(CStr(ListaAtenciones!NroHijo)), _
                                                       CStr(ListaAtenciones!HC_FF_COD), _
                                                       CStr(ListaAtenciones!IdEtnia), _
                                                       CInt(ListaAtenciones!Sexo), _
                                                       idUsuario, _
                                                       oConexion, _
                                                       ml_idEstablecimiento)
    
    oTablaHIS_Detalle.IdTipoFinanciamiento = ListaAtenciones!IdFinanciador
    oTablaHIS_Detalle.Edad = ListaAtenciones!Edad
    oTablaHIS_Detalle.IdTipoEdad = ListaAtenciones!TipoEdad
    oTablaHIS_Detalle.Talla = ListaAtenciones!Talla
    oTablaHIS_Detalle.Peso = ListaAtenciones!Peso
    oTablaHIS_Detalle.IdEstadoaEstablec = ListaAtenciones!IdEstadoaEstablec
    oTablaHIS_Detalle.IdEstadoaServicio = ListaAtenciones!IdEstadoaServicio
    
Case sghHISTipoActividad.ActividadMasiva
    oTablaHIS_Detalle.Peso = 0
    oTablaHIS_Detalle.Edad = ListaAtenciones!Edad
    oTablaHIS_Detalle.CodigoActividad = CStr(ListaAtenciones!HC_FF_COD)
    
Case sghHISTipoActividad.ActividadPreventivaPromocional, sghHISTipoActividad.ActividadConAnimales
    oTablaHIS_Detalle.Peso = 0
    oTablaHIS_Detalle.CodigoActividad = CStr(ListaAtenciones!HC_FF_COD)
End Select

If Not oHIS_Detalle.Insertar(oTablaHIS_Detalle) Then
    bProcesoOK = False
    GoTo ManejadorDeError:
End If

' ================================== REGISTRO DE DIAGNOSTICOS ==============================
If ListaDiagnosticos.RecordCount > 0 Then
    ListaDiagnosticos.MoveFirst
    Do While Not ListaDiagnosticos.EOF
        oHIS_DetalleDiagnosticos.SetDefaults oTablaHIS_DetalleDiagnostico
        
        oTablaHIS_DetalleDiagnostico.IdHisDetalle = oTablaHIS_Detalle.IdHisDetalle 'contiene el id existente en la BD despues de ingresar
        oTablaHIS_DetalleDiagnostico.IdCIE = ListaDiagnosticos!IdCIE
        oTablaHIS_DetalleDiagnostico.CodLAB = IIf(IsNull(ListaDiagnosticos!CodLAB), "", ListaDiagnosticos!CodLAB)
        oTablaHIS_DetalleDiagnostico.IdSubclasificacionDx = ListaDiagnosticos!IdSubclasificacionDx
            
        If Not oHIS_DetalleDiagnosticos.Insertar(oTablaHIS_DetalleDiagnostico) Then
            bProcesoOK = False
        End If
        
        ListaDiagnosticos.MoveNext
    Loop
End If

oConexion.CommitTrans
oConexion.Close
Set oConexion = Nothing
Set oTablaHIS_Detalle = Nothing
Set oTablaHIS_DetalleDiagnostico = Nothing
Set oHIS_Detalle = Nothing
Set oHIS_DetalleDiagnosticos = Nothing

IngresarRegistroHIS = True
Exit Function

ManejadorDeError:
oConexion.RollbackTrans
oConexion.Close
Set oConexion = Nothing
IngresarRegistroHIS = False

ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"

End Function

'======================================================================================================================================================
'======================================================================================================================================================

Function ActualizarRegistroHIS(mi_CodigoCabeceraHIS As Integer, idUsuario As Long, ListaAtenciones As ADODB.Recordset, ListaDiagnosticos As ADODB.Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, ml_idEstablecimiento As Long) As Boolean
Dim oConexion As New ADODB.Connection
Dim oTablaHIS_Detalle As New DOHIS_Detalle
Dim oTablaHIS_DetalleDiagnostico As New DOHIS_DetalleDiagnostico
Dim oHIS_Detalle As New HIS_Detalle
Dim oHIS_DetalleDiagnosticos As New HIS_DetalleDiagnostico

Dim bProcesoOK As Boolean
Dim sSql As String

On Error GoTo ManejadorDeError
ms_MensajeError = ""
bProcesoOK = True

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient

Set oHIS_Detalle.Conexion = oConexion
Set oHIS_DetalleDiagnosticos.Conexion = oConexion

oConexion.BeginTrans
' ================================== REGISTRO DE ATENCION ==============================
ListaAtenciones.MoveFirst
oHIS_Detalle.SetDefaults oTablaHIS_Detalle

'Datos comunes
oTablaHIS_Detalle.IdHisCabecera = mi_CodigoCabeceraHIS 'Se toma del dato externo.
'-------------------------------------------------------------
oTablaHIS_Detalle.IdHisDetalle = ListaAtenciones!IdHisDetalle
'-------------------------------------------------------------
oTablaHIS_Detalle.IdTipoAtencion = ListaAtenciones!IdTipoAtencion
oTablaHIS_Detalle.DiaAtencion = ListaAtenciones!DiaAtencion
oTablaHIS_Detalle.NroRegistroHoja = ListaAtenciones!NroRegistroHoja
oTablaHIS_Detalle.NroRegistroLote = 0
oTablaHIS_Detalle.IdDistrito = ListaAtenciones!IdDistrito
oTablaHIS_Detalle.IdUsuarioAuditoria = idUsuario

'Analizamos el tipo de entrada de los regsitro dependiendo del tipo de atencion
Select Case CInt(ListaAtenciones!IdTipoAtencion)

Case sghHISTipoActividad.Atencion
    'Analisis de un paciente    -   NAC - TIPO DOC. - NROIDENTIFICACION - HC/FF - SEXO - ETNIA
    oTablaHIS_Detalle.IdHisPaciente = ProcesarPaciente(CLng(ListaAtenciones!IdPacienteGalenHos), _
                                                       CInt(ListaAtenciones!IdNacionalidad), _
                                                       CLng(ListaAtenciones!IdTipoDocIdentidad), _
                                                       CStr(ListaAtenciones!NroDocIdentidad), _
                                                       CStr(ListaAtenciones!NroHijo), _
                                                       CStr(ListaAtenciones!HC_FF_COD), _
                                                       CStr(ListaAtenciones!IdEtnia), _
                                                       CInt(ListaAtenciones!Sexo), _
                                                       idUsuario, _
                                                       oConexion, _
                                                       ml_idEstablecimiento)
            
    oTablaHIS_Detalle.IdTipoFinanciamiento = ListaAtenciones!IdFinanciador
    oTablaHIS_Detalle.Edad = ListaAtenciones!Edad
    oTablaHIS_Detalle.IdTipoEdad = ListaAtenciones!TipoEdad
'    oTablaHIS_Detalle.CodigoActividad = CStr(ListaAtenciones!HC_FF_COD)
    oTablaHIS_Detalle.Talla = ListaAtenciones!Talla
    oTablaHIS_Detalle.Peso = ListaAtenciones!Peso
    oTablaHIS_Detalle.IdEstadoaEstablec = ListaAtenciones!IdEstadoaEstablec
    oTablaHIS_Detalle.IdEstadoaServicio = ListaAtenciones!IdEstadoaServicio
    
Case sghHISTipoActividad.ActividadMasiva
    oTablaHIS_Detalle.IdTipoEdad = ListaAtenciones!TipoEdad
    oTablaHIS_Detalle.Edad = ListaAtenciones!Edad
    oTablaHIS_Detalle.CodigoActividad = CStr(ListaAtenciones!HC_FF_COD)
    
Case sghHISTipoActividad.ActividadPreventivaPromocional, sghHISTipoActividad.ActividadConAnimales
    oTablaHIS_Detalle.CodigoActividad = CStr(ListaAtenciones!HC_FF_COD)
End Select

If Not oHIS_Detalle.modificar(oTablaHIS_Detalle) Then
    bProcesoOK = False
    GoTo ManejadorDeError:
End If

' ================================== ACTUALIZACION DE DIAGNOSTICOS ==============================
'Eliminamos todos los DX
If Not oHIS_DetalleDiagnosticos.Eliminar(oTablaHIS_Detalle.IdHisDetalle) Then
    bProcesoOK = False
End If

If ListaDiagnosticos.RecordCount > 0 Then
    ListaDiagnosticos.MoveFirst
    Do While Not ListaDiagnosticos.EOF
    oHIS_DetalleDiagnosticos.SetDefaults oTablaHIS_DetalleDiagnostico
    
        oTablaHIS_DetalleDiagnostico.IdHisDetalle = oTablaHIS_Detalle.IdHisDetalle 'contiene el id existente en la BD despues de ingresar
        oTablaHIS_DetalleDiagnostico.IdCIE = ListaDiagnosticos!IdCIE
        oTablaHIS_DetalleDiagnostico.CodLAB = IIf(IsNull(ListaDiagnosticos!CodLAB), "", ListaDiagnosticos!CodLAB)
        oTablaHIS_DetalleDiagnostico.IdSubclasificacionDx = ListaDiagnosticos!IdSubclasificacionDx
           
        If Not oHIS_DetalleDiagnosticos.Insertar(oTablaHIS_DetalleDiagnostico) Then
            bProcesoOK = False
        End If
        ListaDiagnosticos.MoveNext
    Loop
End If

oConexion.CommitTrans
oConexion.Close
Set oConexion = Nothing
Set oTablaHIS_Detalle = Nothing
Set oTablaHIS_DetalleDiagnostico = Nothing
Set oHIS_Detalle = Nothing
Set oHIS_DetalleDiagnosticos = Nothing


ActualizarRegistroHIS = True
Exit Function

ManejadorDeError:
oConexion.RollbackTrans
oConexion.Close
Set oConexion = Nothing
ActualizarRegistroHIS = False
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

'======================================================================================================================================================
'======================================================================================================================================================

Function EliminarRegistroHIS(idAtencion As Long, idUsuario As Long, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Detalle As New HIS_Detalle
Dim mo_HIS_DetalleDiagnosticos As New HIS_DetalleDiagnostico

Dim oRcs_ListaIdDiagnosticos As New ADODB.Recordset
Dim bProcesoOK As Boolean
Dim sSql As String

On Error GoTo ManejadorDeError
ms_MensajeError = ""

bProcesoOK = True
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient

Set mo_HIS_Detalle.Conexion = oConexion
Set mo_HIS_DetalleDiagnosticos.Conexion = oConexion

oConexion.BeginTrans

'------- OBTENCION DE IDS DE DIAGNOSTICOS DE CABECERA --------------
Set oRcs_ListaIdDiagnosticos = ObtenerDatosDetalleDiagnosticoPorIdDetalle(idAtencion)

'------- ELIMINACION DE DIAGNOSTICOS -------------------------------
If oRcs_ListaIdDiagnosticos.RecordCount <> 0 Then
    Dim oTablaDOHIS_DetalleDiagnostico As New DOHIS_DetalleDiagnostico
    oRcs_ListaIdDiagnosticos.MoveFirst
    
    Do While Not oRcs_ListaIdDiagnosticos.EOF
        oTablaDOHIS_DetalleDiagnostico.idhisdetallediagnostico = oRcs_ListaIdDiagnosticos.Fields!idhisdetallediagnostico
        mo_HIS_DetalleDiagnosticos.Eliminar oRcs_ListaIdDiagnosticos.Fields!IdHisDetalle
        Exit Do
        oRcs_ListaIdDiagnosticos.MoveNext
    Loop
End If
'------- ELIMINACION DE ATENCION ---------------------------------
Dim oTablaDOHIS_Detalle As New DOHIS_Detalle
oTablaDOHIS_Detalle.IdHisDetalle = idAtencion
If Not mo_HIS_Detalle.Eliminar(oTablaDOHIS_Detalle) Then
    GoTo ManejadorDeError
End If

oConexion.CommitTrans
oConexion.Close
Set oConexion = Nothing
EliminarRegistroHIS = True
Exit Function

ManejadorDeError:
oConexion.RollbackTrans
oConexion.Close
'oConexion = Null
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

'Ingresa un paciente en la Base de Datos del HIS_GALENHOS
Private Function ProcesarPaciente(IdPacienteGalenHos As Long, _
                                  IdNacionalidad As Long, _
                                  IdTipoDoc As Long, _
                                  NroDocIdentidad As String, _
                                  NroHijo As String, _
                                  HC_FF_COD As String, _
                                  IdEtnia As String, _
                                  Sexo As Integer, _
                                  idUsuario As Long, _
                                  oConexion As ADODB.Connection, _
                                  ml_idEstablecimiento As Long)

Dim mo_DatosPaciente As New SIGHDatos.HIS_Paciente
Dim mo_DatosEstablecimientoPaciente As New SIGHDatos.HIS_EstablecPacienteHIS
Dim oTablaHIS_Paciente As New DOHIS_Paciente
Dim oTablaHIS_EstabPacienteHIS As New DOHIS_EstablecPacienteHIS

Dim mb_ExistePacHIS As Boolean
Dim oHIS_Paciente As New HIS_Paciente
'Dim mo_RcsPacienteHIS As New Recordset

Set mo_DatosPaciente.Conexion = oConexion
Set mo_DatosEstablecimientoPaciente.Conexion = oConexion
Set oHIS_Paciente.Conexion = oConexion

'Verificacion de Datos del Paciente en HIS
Dim oTablaHIS_PacienteTemp As New DOHIS_Paciente
If IdTipoDoc = 8 Then
    Set oTablaHIS_PacienteTemp = mo_DatosPaciente.VerificarPacienteNroHijo(NroDocIdentidad, CInt(IdTipoDoc), Val(Trim(NroHijo)))
Else
    Set oTablaHIS_PacienteTemp = mo_DatosPaciente.VerificarPaciente(NroDocIdentidad, CInt(IdTipoDoc))
End If

If oTablaHIS_PacienteTemp.IdHisPaciente <> 0 Then
    mb_ExistePacHIS = True
Else
    mb_ExistePacHIS = False
End If

'Proceso de Ingreso de Datos del Paciente
If mb_ExistePacHIS Then
    'Verifica si : HIS PAciente no tiene ID PAciente GalenHos y se encontro el Id Paciente GalenHos
'    If oTablaHIS_PacienteTemp.IdPacienteGalenHos = 0 And IdPacienteGalenHos <> 0 Then
'        'Actualiza Todos los registros que tengas el DNI Encontrado - Debido a que pueden
'        'haber ingresado la misma persona en diferentes establecimientos
'        ActualizarPacientesHIS NroDocIdentidad, IdPacienteGalenHos
'    End If
    
    'Devuelve ID de HIS Paciente
    ProcesarPaciente = oTablaHIS_PacienteTemp.IdHisPaciente
    
    'llenamos de datos y modificamos
    oTablaHIS_Paciente.IdHisPaciente = ProcesarPaciente
    oTablaHIS_Paciente.IdNacionalidad = IdNacionalidad
    oTablaHIS_Paciente.IdTipoDocumento = IdTipoDoc
    oTablaHIS_Paciente.NroDocIdentidad = NroDocIdentidad
    oTablaHIS_Paciente.NroHijo = NroHijo
    oTablaHIS_Paciente.nrohc_ff = HC_FF_COD
    oTablaHIS_Paciente.IdEtnia = IdEtnia
    oTablaHIS_Paciente.Sexo = Sexo
    oTablaHIS_Paciente.IdUsuarioAuditoria = idUsuario
    
    'Ingresa dato del ID de Paciente en GalenHos si hubiera
    If IdPacienteGalenHos <> 0 Then
        oTablaHIS_Paciente.IdPacienteGalenHos = IdPacienteGalenHos
    End If
   
    oHIS_Paciente.modificar oTablaHIS_Paciente
    ProcesarPaciente = oTablaHIS_Paciente.IdHisPaciente
   
   'INGRESAMOS UN REGISTRO MAS EN LA ENTIDAD QUE RELACIONA
   'EL ESTABLECIMIENTO Y EL PACIENTE CON SU HC RESPECTIVA
    
    oTablaHIS_EstabPacienteHIS.IdEstablecimiento = ml_idEstablecimiento
    oTablaHIS_EstabPacienteHIS.IdHisPaciente = oTablaHIS_Paciente.IdHisPaciente
    oTablaHIS_EstabPacienteHIS.nrohc_ff = HC_FF_COD
    
    mo_DatosEstablecimientoPaciente.modificar oTablaHIS_EstabPacienteHIS
Else
    'llenamos de datos y ingresamos
    oTablaHIS_Paciente.IdNacionalidad = IdNacionalidad
    oTablaHIS_Paciente.IdTipoDocumento = IdTipoDoc
    oTablaHIS_Paciente.NroDocIdentidad = NroDocIdentidad
    oTablaHIS_Paciente.NroHijo = NroHijo
    oTablaHIS_Paciente.IdEtnia = IdEtnia
    oTablaHIS_Paciente.Sexo = Sexo
    oTablaHIS_Paciente.IdUsuarioAuditoria = idUsuario
    
    'Ingresa dato del ID de Paciente en GalenHos si hubiera
    If IdPacienteGalenHos <> 0 Then
        oTablaHIS_Paciente.IdPacienteGalenHos = IdPacienteGalenHos
    End If
   
    oHIS_Paciente.Insertar oTablaHIS_Paciente
    ProcesarPaciente = oTablaHIS_Paciente.IdHisPaciente
   
   'INGRESAMOS UN REGISTRO MAS EN LA ENTIDAD QUE RELACIONA
   'EL ESTABLECIMIENTO Y EL PACIENTE CON SU HC RESPECTIVA
   
    oTablaHIS_EstabPacienteHIS.IdEstablecimiento = ml_idEstablecimiento
    oTablaHIS_EstabPacienteHIS.IdHisPaciente = oTablaHIS_Paciente.IdHisPaciente
    oTablaHIS_EstabPacienteHIS.nrohc_ff = HC_FF_COD
    
    mo_DatosEstablecimientoPaciente.Insertar oTablaHIS_EstabPacienteHIS
End If
End Function

Private Function ActualizarPacientesHIS(NroDocumentoIdentidad As String, IdPacienteGalenHos As Long)
Dim oConexion As New ADODB.Connection
Dim mo_DatosPaciente As New SIGHDatos.HIS_Paciente

On Error GoTo ActualizarPacientesHIS_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_DatosPaciente.Conexion = oConexion

ActualizarPacientesHIS = mo_DatosPaciente.ActualizarPacienteHIS(NroDocumentoIdentidad, IdPacienteGalenHos)

On Error GoTo 0
Exit Function
ActualizarPacientesHIS_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ActualizarPacientesHIS of Módulo de clase ReglasHISGalenos"
End Function

Private Function ListaIdsAtencionesPorIdCabecera(RegCabeceraHIS As DOHIS_Cabecera) As ADODB.Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Detalle As New HIS_Detalle

On Error GoTo ListaIdsAtencionesPorIdCabecera_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Detalle.Conexion = oConexion

Set ListaIdsAtencionesPorIdCabecera = mo_HIS_Detalle.ListaIdsAtencionesPorIdCabecera(RegCabeceraHIS)

On Error GoTo 0
Exit Function
ListaIdsAtencionesPorIdCabecera_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaIdsAtencionesPorIdCabecera of Módulo de clase ReglasHISGalenos"
End Function

Private Function ListaIdsDiagnosticosPorIdCabecera(RegCabeceraHIS As DOHIS_Cabecera) As ADODB.Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_DetalleDiagnostico As New HIS_DetalleDiagnostico

On Error GoTo ListaIdsDiagnosticosPorIdCabecera_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_DetalleDiagnostico.Conexion = oConexion

Set ListaIdsDiagnosticosPorIdCabecera = mo_HIS_DetalleDiagnostico.ListaIdsDiagnosticosPorIdCabecera(RegCabeceraHIS)

On Error GoTo 0
Exit Function
ListaIdsDiagnosticosPorIdCabecera_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaIdsDiagnosticosPorIdCabecera of Módulo de clase ReglasHISGalenos"
End Function

'======================================================================================================================================================
'======================================================================================================================================================

'============================================== USERCONTROL ATENCIONES - LISTA DE ATENCIONES DE MR ==========================================
Function ConsultarRegistroFiltroAtenciones(IdEstablecimiento As Long, NombreLote As String, Anio As String, IdMes As Integer) As Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_His_cabecera As New HIS_Cabecera

On Error GoTo ConsultarRegistroFiltroAtenciones_Error
oConexion.CommandTimeout = 300
oConexion.CursorLocation = adUseClient
oConexion.Open sighentidades.CadenaConexion

Set mo_His_cabecera.Conexion = oConexion
    
Set ConsultarRegistroFiltroAtenciones = mo_His_cabecera.ConsultarRegistroFiltroAtenciones(IdEstablecimiento, NombreLote, Anio, IdMes)

On Error GoTo 0
Exit Function
ConsultarRegistroFiltroAtenciones_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ConsultarRegistroFiltroAtenciones of Módulo de clase ReglasHISGalenos"
End Function

'============================================== USERCONTROL LOTES - LISTA DE LOTES MR ==========================================
Function ConsultarRegistroFiltroLotes(IdEstablecimiento As Long, Anio As Integer, IdMes As Integer, Lote As String, IdCerrado As Boolean) As Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo ConsultarRegistroFiltroLotes_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion
    
Set ConsultarRegistroFiltroLotes = mo_HIS_Lotes.ConsultarRegistroFiltroLotes(IdEstablecimiento, Anio, IdMes, Lote, IdCerrado)

On Error GoTo 0
Exit Function
ConsultarRegistroFiltroLotes_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ConsultarRegistroFiltroLotes of Módulo de clase ReglasHISGalenos"
End Function

Function His_ConsultarHojasRegistradas(IdEstablecimiento As Long, IdLote As Long) As Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo His_ConsultarHojasRegistradas_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion
    
Set His_ConsultarHojasRegistradas = mo_HIS_Lotes.His_ConsultarHojasRegistradas(IdEstablecimiento, IdLote)

On Error GoTo 0
Exit Function
His_ConsultarHojasRegistradas_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure His_ConsultarHojasRegistradas of Módulo de clase ReglasHISGalenos"
End Function

Function His_ConsultarTotalRegistrosLote(IdEstablecimiento As Long, IdLote As Long) As Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo His_ConsultarTotalRegistrosLote_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion
    
Set His_ConsultarTotalRegistrosLote = mo_HIS_Lotes.His_ConsultarTotalRegistrosLote(IdEstablecimiento, IdLote)

On Error GoTo 0
Exit Function
His_ConsultarTotalRegistrosLote_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure His_ConsultarTotalRegistrosLote of Módulo de clase ReglasHISGalenos"
End Function
'
Function HIS_ConsultarRegMuestraLotes(IdLote As Long) As Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo HIS_ConsultarRegMuestraLotes_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion
    
Set HIS_ConsultarRegMuestraLotes = mo_HIS_Lotes.HIS_ConsultarRegMuestraLotes(IdLote)

On Error GoTo 0
Exit Function
HIS_ConsultarRegMuestraLotes_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HIS_ConsultarRegMuestraLotes of Módulo de clase ReglasHISGalenos"
End Function

'============================================== USERCONTROL PROGRAMACION - LISTA DE MEDICOS DE MR ==========================================
Public Function ObtenerListaMedicosMR(IdEspecialidad As Long) As Recordset
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_His_cabecera As New HIS_Cabecera

On Error GoTo ListarMedicosMR_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion

Set ObtenerListaMedicosMR = mo_His_cabecera.ObtenerListaMedicosMR(IdEspecialidad)

On Error GoTo 0
Exit Function
ListarMedicosMR_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerListaMedicosMR of Módulo de clase ReglasHISGalenos"
End Function

'============================================== USERCONTROL PROGRAMACION - LISTA SERVICIOS POR ESTBLECIMIENTOS ELEGIDO ==========================================
Public Function ListarEspecialidadesEstablecimientosExternos() As Recordset
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    
    On Error GoTo ManejadorDeError
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_DatosServiciosPorEstab.Conexion = oConexion
    
    Set ListarEspecialidadesEstablecimientosExternos = mo_DatosServiciosPorEstab.ListarEspecialidadesEstablecimientosExternos()
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

'============================================== USERCONTROL PROGRAMACION - OBTENER DATOS DE PROGRAMACION MEDICA ==========================================
Public Function ObtenerDatosProgramacionMedica(IdEstablecimiento As Long, IdServicio As Long, IdMedico As Long, Anio As Integer, IdMes As Integer, IdTurno As Integer) As ADODB.Recordset
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_HIS_ProgMedEstMR As New HIS_ProgMedEstMR
    
    On Error GoTo ObtenerProgramacionMedica_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_HIS_ProgMedEstMR.Conexion = oConexion
    
    Set ObtenerDatosProgramacionMedica = mo_HIS_ProgMedEstMR.ObtenerDatosProgramacionMedica(IdEstablecimiento, IdServicio, IdMedico, Anio, IdMes, IdTurno)
    
    On Error GoTo 0
    Exit Function
ObtenerProgramacionMedica_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerProgramacionMedica of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMUALRIO PROGRAMACION - LISTAR SERVICIOS POR ESPECIALIDAD Y ESTABLECIMIENTO ==========================================
Function ListaServiciosPorEstablecimientoYEspecialidad(IdEspecialidad As Long, IdEstablecimiento As Long) As Recordset
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_HIS_ProgMedEstMR As New HIS_ProgMedEstMR
    
    On Error GoTo ListaServiciosPorEstablecimientoYEspecialidad_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_HIS_ProgMedEstMR.Conexion = oConexion
    
    Set ListaServiciosPorEstablecimientoYEspecialidad = mo_HIS_ProgMedEstMR.ListaServiciosPorEstablecimientoYEspecialidad(IdEspecialidad, IdEstablecimiento)
    
    On Error GoTo 0
    Exit Function
ListaServiciosPorEstablecimientoYEspecialidad_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaServiciosPorEstablecimientoYEspecialidad of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMUALRIO LISTA LOTES - OBTENER NUMERO DE HOJAS UTILIZADAS EN UN LOTE ==========================================
Function ObtenerDatosNumeroHojasUtilizadas(idhislote As Long) As Integer
    Dim oConexion As New ADODB.Connection
    Dim mo_HIS_Lotes As New HIS_Lotes
    On Error GoTo ObtenerDatosNumeroHojasUtilizadas_Error
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_HIS_Lotes.Conexion = oConexion
    ObtenerDatosNumeroHojasUtilizadas = mo_HIS_Lotes.ObtenerDatosNumeroHojasUtilizadas(idhislote)
    
    On Error GoTo 0
    Exit Function
ObtenerDatosNumeroHojasUtilizadas_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosNumeroHojasUtilizadas of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMUALRIO PRINCIPAL - OBTIENE EL ID DE LA ESPECIALIDAD A PARTIR DE UN SERVICIO VINCULADO ==========================================
Function ConsultarIdEspecialidadPorIdServicio(IdServicio As Long) As Long
    Dim oConexion As New ADODB.Connection
    Dim mo_DatosServicios As New SIGHDatos.Servicios
    
    On Error GoTo ConsultarIdEspecialidadPorIdServicio_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_DatosServicios.Conexion = oConexion
    
    ConsultarIdEspecialidadPorIdServicio = mo_DatosServicios.DevuelveEspecialidadDelServicio(IdServicio)
    
    On Error GoTo 0
    Exit Function
ConsultarIdEspecialidadPorIdServicio_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ConsultarIdEspecialidadPorIdServicio of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMUALRIO PRINCIPAL - VALIDA SI EL MEDICO ACTUAL ES DEL ESTABLECIMIENTO ==========================================
Public Function ValidarMedicoEstablecimiento(ml_IdEmpleado As Long, ml_idEstablecimiento As Long) As Boolean
    On Error GoTo ManejadorDeError
    Dim oConexion As New ADODB.Connection
    Dim mo_ProgMedEstMR As New SIGHDatos.HIS_ProgMedEstMR
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set mo_ProgMedEstMR.Conexion = oConexion
    ValidarMedicoEstablecimiento = mo_ProgMedEstMR.ValidarMedicoEstablecimiento(ml_IdEmpleado, ml_idEstablecimiento)
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

'============================================== FORMUALRIO PRINCIPAL - OBTIENE LOS DATOS DEL MEDICO ACTUAL ==========================================
Public Function ObtenerDatosMedicoPorId(oTablaMedico As DOMedico, oTablaEmpleado As DOEmpleado) As Boolean
Dim oConexion As New ADODB.Connection
Dim mb_ObtenerMedicoPorId As Boolean
Dim mo_DatosMedicos As New SIGHDatos.Medicos
Dim mo_DatosEmpleados As New SIGHDatos.Empleados

On Error GoTo ObtenerDatosMedicoPorId_Error
mb_ObtenerMedicoPorId = False
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_DatosMedicos.Conexion = oConexion
Set mo_DatosEmpleados.Conexion = oConexion

If mo_DatosMedicos.SeleccionarPorId(oTablaMedico) Then
    oTablaEmpleado.idEmpleado = oTablaMedico.idEmpleado
    If mo_DatosEmpleados.SeleccionarPorId(oTablaEmpleado) Then
        mb_ObtenerMedicoPorId = True
    End If
End If
ObtenerDatosMedicoPorId = mb_ObtenerMedicoPorId

On Error GoTo 0
Exit Function
ObtenerDatosMedicoPorId_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosMedicoPorId of Módulo de clase ReglasHISGalenos"
End Function

'============================================== FORMUALRIO PRINCIPAL - OBTIENE CODIGO NACIONALIDAD POR EL ID DE NACIONALIDAD ==========================================
Function ObtenerDatosCodNacPorIdNac(IdNacionalidad As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim mb_ObtenerMedicoPorId As Boolean
Dim mo_His_cabecera As New HIS_Cabecera
On Error GoTo ObtenerDatosCodNacPorIdNac_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion
Set ObtenerDatosCodNacPorIdNac = mo_His_cabecera.ObtenerDatosPais(IdNacionalidad)

On Error GoTo 0
Exit Function
ObtenerDatosCodNacPorIdNac_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosCodNacPorIdNac of Módulo de clase ReglasHISGalenos"
End Function

Function ObtenerDatosCodNacPorCodigo(Codigo As String) As Recordset
Dim oConexion As New ADODB.Connection
Dim mb_ObtenerMedicoPorId As Boolean
Dim mo_His_cabecera As New HIS_Cabecera
On Error GoTo ObtenerDatosCodNacPorCodigo_Error

oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_His_cabecera.Conexion = oConexion
Set ObtenerDatosCodNacPorCodigo = mo_His_cabecera.His_ObtenerIdPaisPorCodNac(Codigo)

On Error GoTo 0
Exit Function
ObtenerDatosCodNacPorCodigo_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosCodNacPorCodigo of Módulo de clase ReglasHISGalenos"
End Function


'============================================== FORMUALRIO PRINCIPAL - ACTUALIZA LA HOJA DEL LOTE DEL HIS ==========================================
Public Function ModificarRegistroLoteHISAdicionarHoja(ml_IdLote As Long, NumHoja As Integer) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo ModificarRegistroLoteHISAdicionarHoja_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion

ModificarRegistroLoteHISAdicionarHoja = mo_HIS_Lotes.ModificarRegistroLoteHISAdicionarHoja(ml_IdLote, NumHoja)

On Error GoTo 0
Exit Function
ModificarRegistroLoteHISAdicionarHoja_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ModificarRegistroLoteHISAdicionarHoja of Módulo de clase ReglasHISGalenos"
End Function

'===========================================================================================================
'                                   METODOS DE EXPORTACION DE REGISTROS DEL HIS
'===========================================================================================================

'metodo de extraccion de informacion de atenciones
Public Function ExportacionHIS_Atenciones(idUsuario As Long, ml_mes As Long, mi_anio As Long) As Recordset
Dim oRcs_Atenciones As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Atenciones As New HIS_Detalle

On Error GoTo ExportacionHIS_Atenciones_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Atenciones.Conexion = oConexion

Set ExportacionHIS_Atenciones = mo_HIS_Atenciones.ExportacionHIS_Atenciones(idUsuario, ml_mes, mi_anio)

On Error GoTo 0
Exit Function
ExportacionHIS_Atenciones_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExportacionHIS_Atenciones of Módulo de clase ReglasHISGalenos"
End Function

'metodo de extraccion de informacion de diagnosticos
Public Function ExportacionHIS_Diagnosticos(idUsuario As Long, ml_mes As Long, mi_anio As Long) As Recordset
Dim oRcs_Diagnosticos As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Diagnosticos As New HIS_DetalleDiagnostico

On Error GoTo ExportacionHIS_Diagnosticos_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Diagnosticos.Conexion = oConexion

Set ExportacionHIS_Diagnosticos = mo_HIS_Diagnosticos.ExportacionHIS_Diagnosticos(idUsuario, ml_mes, mi_anio)

On Error GoTo 0
Exit Function
ExportacionHIS_Diagnosticos_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ExportacionHIS_Diagnosticos of Módulo de clase ReglasHISGalenos"
End Function

'metodo para averiguar los datos de establecimiento desde el formualrio de ingrso de atenciones
Public Function BusquedaDistrito(IdDistrito As Long) As DODistrito
Dim mo_doDistrito As New DODistrito
Dim mo_Distrito As New Distritos
Dim oConexion As New ADODB.Connection

On Error GoTo BusquedaDistrito_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_Distrito.Conexion = oConexion

mo_doDistrito.IdDistrito = IdDistrito
mo_Distrito.SeleccionarPorId mo_doDistrito
Set BusquedaDistrito = mo_doDistrito

On Error GoTo 0
Exit Function
BusquedaDistrito_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure BusquedaDistrito of Módulo de clase ReglasHISGalenos"
End Function

'MODIFICADO POR YEPE NOVIEMBRE
Function SisafilicianesObtenerNombreApellidosPorNroDocumentoYtipo(lcNroDocumento As String, lnIdDocIdentidad As Long) As ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oRs As New ADODB.Recordset
Dim oConexionExterna As New Connection
Dim wxParametroJAMO As String
wxParametroJAMO = lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)

    oConexionExterna.CommandTimeout = 300
    oConexionExterna.CursorLocation = adUseClient
    oConexionExterna.Open wxParametroJAMO

Set SisafilicianesObtenerNombreApellidosPorNroDocumentoYtipo = Nothing
With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexionExterna
        .CommandTimeout = 150
        .CommandText = "SisafilicianesObtenerNombreApellidosPorNroDocumentoYtipo"
        Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 10, Left(lcNroDocumento, 10)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@DocumentoTipo", adChar, adParamInput, 1, lnIdDocIdentidad): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
End With
   
Set SisafilicianesObtenerNombreApellidosPorNroDocumentoYtipo = oRs
   oConexionExterna.Close
   Set oCommand = Nothing
   Set oConexionExterna = Nothing
End Function

'INGRESADO POR EYPE 22 NOVIEMBRE
Function ValidaConsistenciaDiagnosticosHis(IdLote As Long, DiaAtencion As Integer, NroPagina As Integer, Edad As Integer, TipoEdadHis As String, SexoHis As String, PesoHis As Currency, _
                                            FichaFamHis As String, Establecimiento As String, Servicio As String, ListaDiagnosticos As ADODB.Recordset) As String
' ================================== VALIDAR CONSISTENCIA DE REGISTRO DE DIAGNOSTICOS ==============================
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRsBuscarInconsistencia As New ADODB.Recordset
    Dim oRcs_DetalleProductosHis As New ADODB.Recordset
    Dim oRsValidaciones As New ADODB.Recordset
    Dim lcMensajeValidacionDiag As String
    Dim lnNumeroDiagnostico As Integer
    
    ValidaConsistenciaDiagnosticosHis = ""
        
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    'Valida ingreso de diagnosticos
    If ListaDiagnosticos.RecordCount <= 0 Then
        oConexion.Close
        Set oConexion = Nothing
        ValidaConsistenciaDiagnosticosHis = "Ingrese los diagnósticos"
        Exit Function
    Else
        If ListaDiagnosticos.RecordCount > 6 Then
                oConexion.Close
                Set oConexion = Nothing
                ValidaConsistenciaDiagnosticosHis = "Solo se debe ingresar como máximo 6 diagnósticos"
                Exit Function
        End If
        ListaDiagnosticos.MoveFirst
        Do While Not ListaDiagnosticos.EOF
        If IsNull(ListaDiagnosticos!descripcion_cie) Then
            oConexion.Close
            Set oConexion = Nothing
            ValidaConsistenciaDiagnosticosHis = "Antes de grabar asegurese tener la descripción del diagnóstico"
            Exit Function
        Else
'            If CStr(ListaDiagnosticos!descripcion_cie) = "..." Or CStr(ListaDiagnosticos!descripcion_cie) = "" Then
            If CStr(ListaDiagnosticos!descripcion_cie) = "" Then
                oConexion.Close
                Set oConexion = Nothing
                ValidaConsistenciaDiagnosticosHis = "Ingrese los diagnósticos"
                Exit Function
            Else
                Set oRcs_DetalleProductosHis = mr_ReglasHIS.ObtenerListaCodigosProductosHisPorNombreYDescripcion(ListaDiagnosticos!descripcion_cie, "")
                If ListaDiagnosticos!IdCIE = 0 Or IsNull(ListaDiagnosticos!IdCIE) Then
                    If oRcs_DetalleProductosHis.RecordCount = 0 Then
                        oConexion.Close
                        Set oConexion = Nothing
                        ValidaConsistenciaDiagnosticosHis = "El código de diagnóstico ingresado no existe"
                        Exit Function
                    Else
                        ListaDiagnosticos!IdCIE = oRcs_DetalleProductosHis.Fields!IdDiagCpt
                    End If
                End If
                'Valida Diagnostico por Sexo
                If oRcs_DetalleProductosHis.Fields!DxSexo <> 0 Then
                    If oRcs_DetalleProductosHis.Fields!DxSexo <> SexoHis Then
                        oConexion.Close
                        Set oConexion = Nothing
                        ValidaConsistenciaDiagnosticosHis = "El código de diagnóstico '" & ListaDiagnosticos!descripcion_cie & "' no es válido para el TIPO de SEXO del paciente"
                        Exit Function
                    End If
                End If
               
                If ListaDiagnosticos!CodLAB <> "" Then
                    Set oRcs_DetalleProductosHis = mr_ReglasHIS.ObtenerListaCodigosLABporCodigoyNombre(ListaDiagnosticos!CodLAB, "")
                    If oRcs_DetalleProductosHis.RecordCount = 0 Then
                        oConexion.Close
                        Set oConexion = Nothing
                        ValidaConsistenciaDiagnosticosHis = "El código Lab '" & ListaDiagnosticos!CodLAB & "' es incorrecto, por favor corrija"
                        Exit Function
                    End If
                End If
            End If
          End If
          ListaDiagnosticos.MoveNext
        Loop
    End If
    
    'Validar diagnosticos repetidos
    Dim oRcsCloneTemp As New Recordset
    Dim oRcs_Temp2 As New Recordset
    If ListaDiagnosticos.RecordCount > 0 Then
        ListaDiagnosticos.MoveFirst
        Do While Not ListaDiagnosticos.EOF
            Set oRcsCloneTemp = ListaDiagnosticos.Clone(adLockReadOnly)
            oRcsCloneTemp.Filter = "IdCIE=" & ListaDiagnosticos!IdCIE
            oRcsCloneTemp.MoveFirst
            If oRcsCloneTemp.RecordCount > 1 Then
                Set oRcs_Temp2 = mr_ReglasHIS.DevuelveCodigoDiagnosticosHis(ListaDiagnosticos!IdCIE)
                If oRcs_Temp2.RecordCount > 0 Then
                    If oRcs_Temp2!MasDeUnDiagnosticos = 0 Then
                        oRcsCloneTemp.Close
                        Set oRcsCloneTemp = Nothing
                        ValidaConsistenciaDiagnosticosHis = "El diagnóstico  " + ListaDiagnosticos!descripcion_cie + " no puede ser duplicado"
                        Exit Function
                    End If
                End If
            End If
            ListaDiagnosticos.MoveNext
        Loop
    End If
    
          
    'cargamos la tabla temporal
    'Actuallizado Yamill
    Dim ors_temp2 As Recordset
    With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = oConexion
         .CommandTimeout = 150
         .CommandText = "HIS_TEMPORALAgregar"
         
         lnNumeroDiagnostico = 1
         ListaDiagnosticos.MoveFirst
         Do While Not ListaDiagnosticos.EOF
            Set ors_temp2 = DevuelveCodigoDiagnosticosHis(ListaDiagnosticos!IdCIE)
            Set oParameter = .CreateParameter("@Codigo" + Str(lnNumeroDiagnostico), adVarChar, adParamInput, 6, sighentidades.DevuelveCodigoDxSinPUNTO(Trim(ors_temp2.Fields!codigodiagcpt))): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@LabConf" + Str(lnNumeroDiagnostico), adVarChar, adParamInput, 4, IIf(IsNull(ListaDiagnosticos!CodLAB), "", Trim(ListaDiagnosticos!CodLAB))): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Diagnost" + Str(lnNumeroDiagnostico), adVarChar, adParamInput, 1, DevuelveTipoDiagnosticoHis(ListaDiagnosticos!IdSubclasificacionDx)): .Parameters.Append oParameter
            lnNumeroDiagnostico = lnNumeroDiagnostico + 1
            ListaDiagnosticos.MoveNext
         Loop
         While lnNumeroDiagnostico <= 6
            Set oParameter = .CreateParameter("@Codigo" + Str(lnNumeroDiagnostico), adVarChar, adParamInput, 6, ""): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@LabConf" + Str(lnNumeroDiagnostico), adVarChar, adParamInput, 3, ""): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@Diagnost" + Str(lnNumeroDiagnostico), adVarChar, adParamInput, 1, ""): .Parameters.Append oParameter
            lnNumeroDiagnostico = lnNumeroDiagnostico + 1
         Wend
         Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, Edad): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@TipoEdad", adVarChar, adParamInput, 1, TipoEdadHis): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@Sexo", adVarChar, adParamInput, 1, SexoHis): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@Peso", adCurrency, adParamInput, 8, PesoHis): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@FichaFam", adVarChar, adParamInput, 10, Mid(FichaFamHis, 1, 10)): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@Establecimiento", adVarChar, adParamInput, 1, Establecimiento): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@Servicio", adVarChar, adParamInput, 1, Servicio): .Parameters.Append oParameter
         Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1): .Parameters.Append oParameter
         
         .Execute
    End With
    Set oCommand = Nothing
    Set oParameter = Nothing
   
   'Consultamos la lista de validaciones
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "His_ValidacionesConsultarTodos"
        Set oRsValidaciones = .Execute
        Set oRsValidaciones.ActiveConnection = Nothing
    End With
    Set oCommand = Nothing
   
    'Recorremos cada validacion para encontrar si los diagnosticos tienen inconsistencia
    oRsValidaciones.MoveFirst
    Do While Not oRsValidaciones.EOF
        With oCommand
                .CommandType = adCmdStoredProc
                Set .ActiveConnection = oConexion
                .CommandTimeout = 150
                .CommandText = "HIS_TEMPORALBuscarInconsistencia"
                Set oParameter = .CreateParameter("@TxtValidacion", adVarChar, adParamInput, 1000, oRsValidaciones.Fields!txtvalidacion): .Parameters.Append oParameter
                Set oRsBuscarInconsistencia = .Execute
                Set oRsBuscarInconsistencia.ActiveConnection = Nothing
        End With
        Set oCommand = Nothing
        Set oParameter = Nothing
    
        If oRsBuscarInconsistencia.RecordCount > 0 Then
            'Si encuentra inconsistencia, guarda el registro del lote y nro de pagina
            With oCommand
                 .CommandType = adCmdStoredProc
                 Set .ActiveConnection = oConexion
                 .CommandTimeout = 150
                 .CommandText = "HIS_DxOmitidosAgregar"
                 
                    Set oParameter = .CreateParameter("@CODVALIDACION", adInteger, adParamInput, 0, oRsValidaciones.Fields!CODVALIDACION): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdLote", adVarChar, adParamInput, 4, IdLote): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@DiaAtencion", adVarChar, adParamInput, 4, DiaAtencion): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@NroPagina", adInteger, adParamInput, 0, NroPagina): .Parameters.Append oParameter
        
                    Set oParameter = .CreateParameter("@Codigo1", adVarChar, adParamInput, 6, oRsBuscarInconsistencia.Fields!CODIGO1): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@LabConf1", adVarChar, adParamInput, 4, oRsBuscarInconsistencia.Fields!LabConf1): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Diagnost1", adVarChar, adParamInput, 1, oRsBuscarInconsistencia.Fields!diagnost1): .Parameters.Append oParameter
                    
                    Set oParameter = .CreateParameter("@Codigo2", adVarChar, adParamInput, 6, oRsBuscarInconsistencia.Fields!Codigo2): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@LabConf2", adVarChar, adParamInput, 4, oRsBuscarInconsistencia.Fields!LabConf2): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Diagnost2", adVarChar, adParamInput, 1, oRsBuscarInconsistencia.Fields!Diagnost2): .Parameters.Append oParameter
                    
                    Set oParameter = .CreateParameter("@Codigo3", adVarChar, adParamInput, 6, oRsBuscarInconsistencia.Fields!Codigo3): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@LabConf3", adVarChar, adParamInput, 4, oRsBuscarInconsistencia.Fields!LabConf3): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Diagnost3", adVarChar, adParamInput, 1, oRsBuscarInconsistencia.Fields!Diagnost3): .Parameters.Append oParameter
                    
                    Set oParameter = .CreateParameter("@Codigo4", adVarChar, adParamInput, 6, oRsBuscarInconsistencia.Fields!Codigo4): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@LabConf4", adVarChar, adParamInput, 4, oRsBuscarInconsistencia.Fields!LabConf4): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Diagnost4", adVarChar, adParamInput, 1, oRsBuscarInconsistencia.Fields!Diagnost4): .Parameters.Append oParameter
                    
                    Set oParameter = .CreateParameter("@Codigo5", adVarChar, adParamInput, 6, oRsBuscarInconsistencia.Fields!Codigo5): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@LabConf5", adVarChar, adParamInput, 4, oRsBuscarInconsistencia.Fields!LabConf5): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Diagnost5", adVarChar, adParamInput, 1, oRsBuscarInconsistencia.Fields!Diagnost5): .Parameters.Append oParameter
                    
                    Set oParameter = .CreateParameter("@Codigo6", adVarChar, adParamInput, 6, oRsBuscarInconsistencia.Fields!Codigo6): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@LabConf6", adVarChar, adParamInput, 4, oRsBuscarInconsistencia.Fields!LabConf6): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Diagnost6", adVarChar, adParamInput, 1, oRsBuscarInconsistencia.Fields!Diagnost6): .Parameters.Append oParameter
                    
                    Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, oRsBuscarInconsistencia.Fields!Edad): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@TipoEdad", adVarChar, adParamInput, 1, oRsBuscarInconsistencia.Fields!Tip_Edad): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Sexo", adVarChar, adParamInput, 1, oRsBuscarInconsistencia.Fields!Sexo): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Peso", adCurrency, adParamInput, 8, oRsBuscarInconsistencia.Fields!Peso): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@FichaFam", adVarChar, adParamInput, 10, oRsBuscarInconsistencia.Fields!FichaFam): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Establecimiento", adVarChar, adParamInput, 1, oRsBuscarInconsistencia.Fields!Establecimiento): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@Servicio", adVarChar, adParamInput, 1, oRsBuscarInconsistencia.Fields!Servicio): .Parameters.Append oParameter
                    Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 0): .Parameters.Append oParameter
                 
                 .Execute
            End With
            Set oCommand = Nothing
            Set oParameter = Nothing
        
            lcMensajeValidacionDiag = lcMensajeValidacionDiag & vbCrLf & "Codigo de validación Nº " & oRsValidaciones.Fields!CODVALIDACION & ": " & vbCrLf & oRsValidaciones.Fields!txtDescripcion
        End If
        
        oRsValidaciones.MoveNext
    Loop
    
    'Elimina datos de la tabla temporal
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HIS_TEMPORALEliminar"
        .Execute
    End With
    Set oCommand = Nothing
    
    oConexion.Close
    Set oConexion = Nothing
    
    ValidaConsistenciaDiagnosticosHis = lcMensajeValidacionDiag
End Function

'INGRESADO POR YEPE 22 NOVIEMBRE
Public Function DevuelveCodigoDiagnosticosHis(IdDiagnostico As Long) As Recordset
' ================================== VALIDAR CONSISTENCIA DE REGISTRO DE DIAGNOSTICOS ==============================
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRs As New ADODB.Recordset
    
    Set DevuelveCodigoDiagnosticosHis = Nothing
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HIS_FACTCATALOGOSERVICIOSSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdDiagCpt", adInteger, adParamInput, 0, IdDiagnostico): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    Set DevuelveCodigoDiagnosticosHis = oRs
'     DevuelveCodigoDiagnosticosHis = sighentidades.DevuelveCodigoDxSinPUNTO(Trim(oRs.Fields!codigodiagcpt))
'     DevuelveCodigoDiagnosticosHis = Trim(oRs.Fields!CodigoCIE10)
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    
End Function

'INGRESADO POR YEPE 22 NOVIEMBRE
Public Function DevuelveTipoEdadHis(IdTipoEdadHis As Integer) As String
' ================================== VALIDAR CONSISTENCIA DE REGISTRO DE DIAGNOSTICOS ==============================
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRs As New ADODB.Recordset
    
    DevuelveTipoEdadHis = ""
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion


    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HIS_TipoEdadSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdHisTipoEdad", adInteger, adParamInput, 0, IdTipoEdadHis): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
     DevuelveTipoEdadHis = oRs.Fields!CodigoEdad
     
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    
End Function

'INGRESADO POR EYPE 22 NOVIEMBRE
Public Function DevuelveTipoDiagnosticoHis(IdTipoDiagnosticoHis As Integer) As String
' ================================== VALIDAR CONSISTENCIA DE REGISTRO DE DIAGNOSTICOS ==============================
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRs As New ADODB.Recordset
    
    DevuelveTipoDiagnosticoHis = ""
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion


    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "SubclasificacionDiagnosticosSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdSubclasificacionDx", adInteger, adParamInput, 0, IdTipoDiagnosticoHis): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    DevuelveTipoDiagnosticoHis = oRs.Fields!Codigo
    
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    
End Function

'Yamill 07/03/2014
Function HIS_DatosEstablecimientoXidEstablecimiento(ml_idEstablecimiento As Long) As ADODB.Recordset
Dim oRecordset As New Recordset
Dim oConexion As New Connection
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter

    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HIS_DatosEstablecimientoXidEstablecimiento"
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, ml_idEstablecimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set HIS_DatosEstablecimientoXidEstablecimiento = oRecordset
   
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    
End Function

Public Function HisObtenerListaEstablecimientosMRPorEspecialidad(ByVal ml_IdEspecialidad As Long) As Recordset
Dim oRecordset As New Recordset
Dim oConexion As New ADODB.Connection
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sWhere As String
Dim mo_TablaHISPaciente As New DOHIS_Paciente

    Set HisObtenerListaEstablecimientosMRPorEspecialidad = Nothing
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HisObtenerListaEstablecimientosMRPorEspecialidad"
        Set oParameter = .CreateParameter("@IdEspecialidad", adInteger, adParamInput, 0, ml_IdEspecialidad): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set HisObtenerListaEstablecimientosMRPorEspecialidad = oRecordset
 
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
End Function


'''''''''' FUNCIONES PADRON NOMINAL --- YEPE

Function PadronNominalCabeceraAgregar(oDoPadronNominal_Cabecera As DoPadronNominal_Cabecera) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPadronNominal_Cabecera As New PadronNominal_Cabecera
    
    PadronNominalCabeceraAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oPadronNominal_Cabecera.Conexion = oConexion
    If oPadronNominal_Cabecera.Insertar(oDoPadronNominal_Cabecera) Then
        PadronNominalCabeceraAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oPadronNominal_Cabecera.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPadronNominal_Cabecera = Nothing
End Function

Function PadronNominalDetalleAgregar(oDoPadronNominal_Detalle As DoPadronNominal_Detalle) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPadronNominal_Detalle As New PadronNominal_Detalle
    
    PadronNominalDetalleAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oPadronNominal_Detalle.Conexion = oConexion
    If oPadronNominal_Detalle.Insertar(oDoPadronNominal_Detalle) Then
        PadronNominalDetalleAgregar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oPadronNominal_Detalle.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPadronNominal_Detalle = Nothing
End Function

Function PadronNominalDetalleModificar(oDoPadronNominal_Detalle As DoPadronNominal_Detalle) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPadronNominal_Detalle As New PadronNominal_Detalle
    
    PadronNominalDetalleModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oPadronNominal_Detalle.Conexion = oConexion
    If oPadronNominal_Detalle.modificar(oDoPadronNominal_Detalle) Then
        PadronNominalDetalleModificar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oPadronNominal_Detalle.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPadronNominal_Detalle = Nothing
End Function

Function PadronNominalDetalleEliminar(oDoPadronNominal_Detalle As DoPadronNominal_Detalle) As Boolean
Dim oConexion As New ADODB.Connection
Dim oPadronNominal_Detalle As New PadronNominal_Detalle
    
    PadronNominalDetalleEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.BeginTrans
    Set oPadronNominal_Detalle.Conexion = oConexion
    If oPadronNominal_Detalle.Eliminar(oDoPadronNominal_Detalle) Then
        PadronNominalDetalleEliminar = True
        oConexion.CommitTrans
    Else
        ms_MensajeError = oPadronNominal_Detalle.MensajeError
        oConexion.RollbackTrans
    End If
    oConexion.Close
    Set oConexion = Nothing
    Set oPadronNominal_Detalle = Nothing
End Function

Public Function ListaDiagNutricional() As Recordset
    Dim oConexion As New ADODB.Connection
    Dim oCommand As New ADODB.Command
    Dim oRs As New ADODB.Recordset
    On Error GoTo ListaDiagNutricional_Error
    
    Set ListaDiagNutricional = Nothing
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PadNominalListaDiagNutricional"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    Set ListaDiagNutricional = oRs
    oConexion.Close
    Set oCommand = Nothing
    Set oConexion = Nothing
    
    On Error GoTo 0
    Exit Function
ListaDiagNutricional_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListaDiagNutricional of Módulo de clase ReglasHISGalenos"
End Function

'Modificado por Frank 09/04/2014
Public Function PadronNominalFiltrarNroHisClinica(oPadronNominal_Detalle As DoPadronNominal_Detalle) As Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRs As New ADODB.Recordset
    
    Set PadronNominalFiltrarNroHisClinica = Nothing
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PadronNominalFiltrarNroHisClinica"
        Set oParameter = .CreateParameter("@numdocumento", adVarChar, adParamInput, 20, IIf(oPadronNominal_Detalle.NumDocumento = 0, "", CStr(oPadronNominal_Detalle.NumDocumento))): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@apellidopaterno", adVarChar, adParamInput, 50, oPadronNominal_Detalle.ApellidoPaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@apellidomaterno", adVarChar, adParamInput, 50, oPadronNominal_Detalle.ApellidoMaterno): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nombres", adVarChar, adParamInput, 50, oPadronNominal_Detalle.Nombres): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@histclinica", adVarChar, adParamInput, 20, IIf(oPadronNominal_Detalle.HistClinica = "", "", CStr(oPadronNominal_Detalle.HistClinica))): .Parameters.Append oParameter
'        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oPadronNominal_Detalle.IdEstablecimiento = 0, 0, oPadronNominal_Detalle.IdEstablecimiento)): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    Set PadronNominalFiltrarNroHisClinica = oRs
    
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    
End Function

Function PadronNominal_DetalleSeleccionarPorId(ByVal oPadronNominal_Detalle As DoPadronNominal_Detalle) As Boolean
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_PadronNominal_Detalle As New PadronNominal_Detalle
    
    PadronNominal_DetalleSeleccionarPorId = False
    
    On Error GoTo PadronNominal_Detalle_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_PadronNominal_Detalle.Conexion = oConexion
    
    PadronNominal_DetalleSeleccionarPorId = mo_PadronNominal_Detalle.SeleccionarPorId(oPadronNominal_Detalle)
    
    On Error GoTo 0
    Exit Function
PadronNominal_Detalle_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PadronNominal_DetalleSeleccionarPorId of Módulo de clase ReglasHISGalenos"
End Function

Public Function PadNominalBuscarDatosxNroHC(lcNroHistoriaClinica As String) As Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRs As New ADODB.Recordset
    
    Set PadNominalBuscarDatosxNroHC = Nothing
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PadNominalBuscarDatosxNroHC"
        Set oParameter = .CreateParameter("@histclinica", adVarChar, adParamInput, 10, lcNroHistoriaClinica): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    Set PadNominalBuscarDatosxNroHC = oRs
    
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    
End Function

Public Function PadNominalSeleccionarDxNutricionalPorRangoZ(lcRefCodLab As String, ldZ As Double) As Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRs As New ADODB.Recordset
    
    Set PadNominalSeleccionarDxNutricionalPorRangoZ = Nothing
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PadNominalSeleccionarDxNutricionalPorRangoZ"
        Set oParameter = .CreateParameter("@RefCodLab", adVarChar, adParamInput, 10, lcRefCodLab): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Z", adCurrency, adParamInput, 0, ldZ): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    Set PadNominalSeleccionarDxNutricionalPorRangoZ = oRs
    
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    
End Function
Public Function HisLotesXEstablecimientos(IdEstablecimiento As Long) As Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRs As New ADODB.Recordset
    
    Set HisLotesXEstablecimientos = Nothing
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HisLotesXEstablecimientos"
        Set oParameter = .CreateParameter("@idestablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    Set HisLotesXEstablecimientos = oRs
    
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    
End Function

Public Function HisLotesXEstablecimientoyLote(IdEstablecimiento As Integer, idhislote As Integer) As Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRs As New ADODB.Recordset
    
    Set HisLotesXEstablecimientoyLote = Nothing
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HisLotesXEstablecimientoyLote"
        Set oParameter = .CreateParameter("@idestablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idhislote", adInteger, adParamInput, 0, idhislote): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    Set HisLotesXEstablecimientoyLote = oRs
    
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    
End Function


Public Function HIS_CONSULTATOTALREGISTROS(IdEstablecimiento As Integer, idhislote As Integer) As Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRs As New ADODB.Recordset
    
    Set HIS_CONSULTATOTALREGISTROS = Nothing
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HIS_CONSULTATOTALREGISTROS"
        Set oParameter = .CreateParameter("@idestablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idhislote", adInteger, adParamInput, 0, idhislote): .Parameters.Append oParameter
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    
    Set HIS_CONSULTATOTALREGISTROS = oRs
    
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
    
End Function

'Nuevo frannklin 14/04/2014
Public Function ConsultarEstablecimiento(lnIdEstablecimiento As Long) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim mo_His_Establecimientos As New SIGHDatos.His_Establecimientos
    Dim oDOHis_Establecimientos As New SIGHcomun.DOHis_Establecimientos
    
    On Error GoTo ConsultarEstablecimiento_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set mo_His_Establecimientos.Conexion = oConexion
    oDOHis_Establecimientos.IdEstablecimiento = lnIdEstablecimiento
    ConsultarEstablecimiento = mo_His_Establecimientos.SeleccionarPorId(oDOHis_Establecimientos)
    
    On Error GoTo 0
    Exit Function
ConsultarEstablecimiento_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListarServiciosEstablecimiento of Módulo de clase ReglasHISGalenos"
End Function

Public Function HIS_EstadosLoteSeleccionarTodos() As Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New Connection
    Dim oRs As New ADODB.Recordset
    Set HIS_EstadosLoteSeleccionarTodos = Nothing
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "HIS_EstadosLoteSeleccionarTodos"
        Set oRs = .Execute
        Set oRs.ActiveConnection = Nothing
    End With
    Set HIS_EstadosLoteSeleccionarTodos = oRs
    oConexion.Close
    Set oCommand = Nothing
    Set oParameter = Nothing
    Set oConexion = Nothing
End Function

Function His_LotesConsultarFiltro(IdEstablecimiento As Long, IdEstado As Integer) As Recordset
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_HIS_Lotes As New HIS_Lotes
    
    On Error GoTo His_LotesConsultarFiltro_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_HIS_Lotes.Conexion = oConexion
        
    Set His_LotesConsultarFiltro = mo_HIS_Lotes.His_LotesConsultarFiltro(IdEstablecimiento, IdEstado)
    
    On Error GoTo 0
    Exit Function
His_LotesConsultarFiltro_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ConsultarRegistroFiltroLotes of Módulo de clase ReglasHISGalenos"
End Function

Function HIS_BuscaResponsableFiltro(ml_idEstablecimiento As Long, ml_IdServicio As Long, mc_Anio As String, mc_Mes As String, mc_Nombre As String) As Recordset
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_HIS_ProgMedEstMR As New HIS_ProgMedEstMR
    
    On Error GoTo HIS_BuscaResponsableFiltro_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_HIS_ProgMedEstMR.Conexion = oConexion
    Set HIS_BuscaResponsableFiltro = mo_HIS_ProgMedEstMR.HIS_BuscaResponsableFiltro(ml_idEstablecimiento, ml_IdServicio, mc_Anio, mc_Mes, mc_Nombre)
    
    On Error GoTo 0
    Exit Function
    
HIS_BuscaResponsableFiltro_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HIS_BuscaResponsableFiltro of Módulo de clase ReglasHISGalenos"
End Function

Function HISPacientesFiltraPorNroDocumentoYtipo(ml_nrodocumento As String, ml_idtipodoc As Integer) As Recordset
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_HIS_Paciente As New HIS_Paciente
    
    On Error GoTo HISPacientesFiltraPorNroDocumentoYtipo_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_HIS_Paciente.Conexion = oConexion
    Set HISPacientesFiltraPorNroDocumentoYtipo = mo_HIS_Paciente.HISPacientesFiltraPorNroDocumentoYtipo(ml_nrodocumento, ml_idtipodoc)
    
    On Error GoTo 0
    Exit Function
    
HISPacientesFiltraPorNroDocumentoYtipo_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HISPacientesFiltraPorNroDocumentoYtipo of Módulo de clase ReglasHISGalenos"
End Function

Function HisActualizarNroRegistroHisDetalle(ml_IdHisDetalle As Long, ml_NumReg As Long) As Boolean
Dim oRcs_Result As New Recordset
Dim oConexion As New ADODB.Connection
Dim mo_HIS_Lotes As New HIS_Lotes

On Error GoTo HisActualizarNroRegistroHisDetalle_Error
oConexion.Open sighentidades.CadenaConexion
oConexion.CursorLocation = adUseClient
Set mo_HIS_Lotes.Conexion = oConexion

HisActualizarNroRegistroHisDetalle = mo_HIS_Lotes.HisActualizarNroRegistroHisDetalle(ml_IdHisDetalle, ml_NumReg)

On Error GoTo 0
Exit Function
HisActualizarNroRegistroHisDetalle_Error:
MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HisActualizarNroRegistroHisDetalle of Módulo de clase ReglasHISGalenos"
End Function

'
Function HIS_ConsultarRegistrosTotalesLotes(IdLote As Long) As Recordset
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_HIS_Lotes As New HIS_Lotes
    
    On Error GoTo HIS_ConsultarRegistrosTotalesLotes_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_HIS_Lotes.Conexion = oConexion
        
    Set HIS_ConsultarRegistrosTotalesLotes = mo_HIS_Lotes.HIS_ConsultarRegistrosTotalesLotes(IdLote)
    
    On Error GoTo 0
    Exit Function
HIS_ConsultarRegistrosTotalesLotes_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HIS_ConsultarRegistrosTotalesLotes of Módulo de clase ReglasHISGalenos"
End Function



'DOBLE DIGITACION
Function IngresarHISVerificado(ListaAtenciones As ADODB.Recordset) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oDOHIS_Detalle_Verifica As New DOHIS_Detalle_Verifica
    Dim oHIS_Detalle_Verifica As New HIS_Detalle_Verifica
    
    Dim bProcesoOK As Boolean
    Dim sSql As String
    
    On Error GoTo ManejadorDeError
    ms_MensajeError = ""
    bProcesoOK = True
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oHIS_Detalle_Verifica.Conexion = oConexion
    oConexion.BeginTrans
    ' ================================== REGISTRO DE ATENCION ==============================
    ListaAtenciones.MoveFirst
    Do While Not ListaAtenciones.EOF
        oHIS_Detalle_Verifica.SetDefaults oDOHIS_Detalle_Verifica
        'Datos comunes
        oDOHIS_Detalle_Verifica.IdHisDetalle = ListaAtenciones!IdHisDetalle
        oDOHIS_Detalle_Verifica.IdHisCabecera = ListaAtenciones!IdHisCabecera
        oDOHIS_Detalle_Verifica.NroRegistroHoja = ListaAtenciones!NroRegistroHoja
        oDOHIS_Detalle_Verifica.NroRegistroLote = ListaAtenciones!NroRegistroLote
        oDOHIS_Detalle_Verifica.Registrado = 0
        oDOHIS_Detalle_Verifica.Coincide = 0
        If Not oHIS_Detalle_Verifica.Insertar(oDOHIS_Detalle_Verifica) Then
            bProcesoOK = False
            GoTo ManejadorDeError:
        End If
        ListaAtenciones.MoveNext
    Loop
    oConexion.CommitTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oDOHIS_Detalle_Verifica = Nothing
    Set oHIS_Detalle_Verifica = Nothing
    IngresarHISVerificado = True
    Exit Function
ManejadorDeError:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    IngresarHISVerificado = False
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

'============================================== FORMULARIO PRINCIPAL - CARGA DE DATOS DE CABECERA ==========================================
Public Function ObtenerDatosHisDetalleVerif(IdRegHISDetalleVer As Long) As DOHIS_Detalle_Verifica
    Dim oConexion As New ADODB.Connection
    Dim oDOHIS_Detalle_Verifica As New DOHIS_Detalle_Verifica
    Dim oHIS_Detalle_Verifica As New HIS_Detalle_Verifica
    
    On Error GoTo ObtenerDatosHisDetalleVerif_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oHIS_Detalle_Verifica.Conexion = oConexion
    
    oDOHIS_Detalle_Verifica.IdHisDetalle = IdRegHISDetalleVer
    oHIS_Detalle_Verifica.SeleccionarPorId oDOHIS_Detalle_Verifica
    Set ObtenerDatosHisDetalleVerif = oDOHIS_Detalle_Verifica
    
    On Error GoTo 0
    Exit Function
ObtenerDatosHisDetalleVerif_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ObtenerDatosHisDetalleVerif of Módulo de clase ReglasHISGalenos"
End Function

Function IngresarHISDobleDigitacion(oDOHIS_Detalle_Verifica As DOHIS_Detalle_Verifica, ListaDiagnosticos As ADODB.Recordset) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oHIS_Detalle_Verifica As New HIS_Detalle_Verifica
    Dim oHIS_DetalleDx_Verifica As New HIS_DetalleDx_Verifica
    Dim oDOHIS_DetalleDx_Verifica As New DOHIS_DetalleDx_Verifica
    
    Dim bProcesoOK As Boolean
    Dim sSql As String
    
    On Error GoTo ManejadorDeError
    ms_MensajeError = ""
    bProcesoOK = True
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oHIS_Detalle_Verifica.Conexion = oConexion
    Set oHIS_DetalleDx_Verifica.Conexion = oConexion
    oConexion.BeginTrans

    If Not oHIS_Detalle_Verifica.modificar(oDOHIS_Detalle_Verifica) Then
        bProcesoOK = False
        GoTo ManejadorDeError:
    End If
    ' ================================== ACTUALIZACION DE DIAGNOSTICOS ==============================
    'Eliminamos todos los DX
'    If Not oHIS_DetalleDx_Verifica.Eliminar(oDOHIS_Detalle_Verifica.IdHisDetalle) Then
'        bProcesoOK = False
'    End If
    
    If ListaDiagnosticos.RecordCount > 0 Then
         ListaDiagnosticos.MoveFirst
         Do While Not ListaDiagnosticos.EOF
             oHIS_DetalleDx_Verifica.SetDefaults oDOHIS_DetalleDx_Verifica
             oDOHIS_DetalleDx_Verifica.IdHisDetalle = oDOHIS_Detalle_Verifica.IdHisDetalle 'contiene el id existente en la BD despues de ingresar
             oDOHIS_DetalleDx_Verifica.IdCIE = ListaDiagnosticos!IdCIE
             oDOHIS_DetalleDx_Verifica.CodLAB = IIf(IsNull(ListaDiagnosticos!CodLAB), "", ListaDiagnosticos!CodLAB)
             oDOHIS_DetalleDx_Verifica.IdSubclasificacionDx = ListaDiagnosticos!IdSubclasificacionDx
        
             If Not oHIS_DetalleDx_Verifica.Insertar(oDOHIS_DetalleDx_Verifica) Then
                 bProcesoOK = False
             End If
             ListaDiagnosticos.MoveNext
         Loop
    End If
    oConexion.CommitTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oDOHIS_Detalle_Verifica = Nothing
    Set oHIS_Detalle_Verifica = Nothing
    IngresarHISDobleDigitacion = True
    Exit Function
ManejadorDeError:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    IngresarHISDobleDigitacion = False
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function


Function ActualizaHISDobleDigitacion(oDOHIS_Detalle_Verifica As DOHIS_Detalle_Verifica, ListaDiagnosticos As ADODB.Recordset) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oHIS_Detalle_Verifica As New HIS_Detalle_Verifica
    Dim oHIS_DetalleDx_Verifica As New HIS_DetalleDx_Verifica
    Dim oDOHIS_DetalleDx_Verifica As New DOHIS_DetalleDx_Verifica
    
    Dim bProcesoOK As Boolean
    Dim sSql As String
    
    On Error GoTo ManejadorDeError
    ms_MensajeError = ""
    bProcesoOK = True
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oHIS_Detalle_Verifica.Conexion = oConexion
    Set oHIS_DetalleDx_Verifica.Conexion = oConexion
    oConexion.BeginTrans

    If Not oHIS_Detalle_Verifica.modificar(oDOHIS_Detalle_Verifica) Then
        bProcesoOK = False
        GoTo ManejadorDeError:
    End If
    ' ================================== ACTUALIZACION DE DIAGNOSTICOS ==============================
    'Eliminamos todos los DX
    If Not oHIS_DetalleDx_Verifica.Eliminar(oDOHIS_Detalle_Verifica.IdHisDetalle) Then
        bProcesoOK = False
    End If
    
    If ListaDiagnosticos.RecordCount > 0 Then
        ListaDiagnosticos.MoveFirst
        Do While Not ListaDiagnosticos.EOF
            oHIS_DetalleDx_Verifica.SetDefaults oDOHIS_DetalleDx_Verifica
            oDOHIS_DetalleDx_Verifica.IdHisDetalle = oDOHIS_Detalle_Verifica.IdHisDetalle 'contiene el id existente en la BD despues de ingresar
            oDOHIS_DetalleDx_Verifica.IdCIE = ListaDiagnosticos!IdCIE
            oDOHIS_DetalleDx_Verifica.CodLAB = IIf(IsNull(ListaDiagnosticos!CodLAB), "", ListaDiagnosticos!CodLAB)
            oDOHIS_DetalleDx_Verifica.IdSubclasificacionDx = ListaDiagnosticos!IdSubclasificacionDx
        
            If Not oHIS_DetalleDx_Verifica.Insertar(oDOHIS_DetalleDx_Verifica) Then
                bProcesoOK = False
            End If
            ListaDiagnosticos.MoveNext
        Loop
    End If

    oConexion.CommitTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oDOHIS_Detalle_Verifica = Nothing
    Set oHIS_Detalle_Verifica = Nothing
    ActualizaHISDobleDigitacion = True
    Exit Function
ManejadorDeError:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    ActualizaHISDobleDigitacion = False
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function

Function HIS_ConsultarRegistroDetalleHis(mlIdLote As Long) As Recordset
    Dim oRcs_Result As New Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_HIS_Detalle_Verifica As New HIS_Detalle_Verifica
    On Error GoTo HIS_ConsultarRegistroDetalleHis_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_HIS_Detalle_Verifica.Conexion = oConexion
    Set HIS_ConsultarRegistroDetalleHis = mo_HIS_Detalle_Verifica.HIS_ConsultarRegistroDetalleHis(mlIdLote)
    On Error GoTo 0
    Exit Function
HIS_ConsultarRegistroDetalleHis_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HIS_ConsultarRegistroDetalleHis of Módulo de clase ReglasHISGalenos"
End Function

Public Function His_ConsultaDxHisDetalleVerif(ml_IdDetalleHIS As Long) As ADODB.Recordset
    Dim oConexion As New ADODB.Connection
    Dim oRcs_RegistroHISDetalleDiagnostico As New Recordset
    Dim mo_HIS_DetalleDx_Verifica As New HIS_DetalleDx_Verifica
    
    On Error GoTo His_ConsultaDxHisDetalleVerif_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_HIS_DetalleDx_Verifica.Conexion = oConexion
    
    Set oRcs_RegistroHISDetalleDiagnostico = mo_HIS_DetalleDx_Verifica.His_ConsultaDxHisDetalleVerif(ml_IdDetalleHIS)
    If oRcs_RegistroHISDetalleDiagnostico.RecordCount <> 0 Then
        oRcs_RegistroHISDetalleDiagnostico.MoveFirst
    End If
    Set His_ConsultaDxHisDetalleVerif = oRcs_RegistroHISDetalleDiagnostico
    
    On Error GoTo 0
    Exit Function
His_ConsultaDxHisDetalleVerif_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure His_ConsultaDxHisDetalleVerif of Módulo de clase ReglasHISGalenos"
End Function

Function HIS_ConsultarEstadosLote() As Recordset
    Dim oConexion As New ADODB.Connection
    Dim mo_HIS_Lotes As New HIS_Lotes
    On Error GoTo HIS_ConsultarEstadosLote_Error
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set mo_HIS_Lotes.Conexion = oConexion
    Set HIS_ConsultarEstadosLote = mo_HIS_Lotes.HIS_ConsultarEstadosLote
    On Error GoTo 0
    Exit Function
HIS_ConsultarEstadosLote_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HIS_ConsultarEstadosLote of Módulo de clase ReglasHISGalenos"
End Function

Function EliminaHISDobleDigitacion(ml_IdHisLote As Long) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oHIS_Detalle_Verifica As New HIS_Detalle_Verifica
    
    Dim bProcesoOK As Boolean
    Dim sSql As String
    
    On Error GoTo ManejadorDeError
    ms_MensajeError = ""
    bProcesoOK = True
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oHIS_Detalle_Verifica.Conexion = oConexion
    oConexion.BeginTrans

    If Not oHIS_Detalle_Verifica.Eliminar(ml_IdHisLote) Then
        bProcesoOK = False
        GoTo ManejadorDeError:
    End If

    oConexion.CommitTrans
    oConexion.Close
    Set oConexion = Nothing
    Set oHIS_Detalle_Verifica = Nothing
    EliminaHISDobleDigitacion = True
    Exit Function
ManejadorDeError:
    oConexion.RollbackTrans
    oConexion.Close
    Set oConexion = Nothing
    EliminaHISDobleDigitacion = False
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de negocios"
End Function


Function PadronNominal_DetalleSeleccionarTodos() As ADODB.Recordset
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oConexion As New Connection
    Dim oParameter As ADODB.Parameter
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "PadronNominal_DetalleSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
    Set PadronNominal_DetalleSeleccionarTodos = oRecordset
    Set oRecordset = Nothing
    Set oCommand = Nothing
    Set oParameter = Nothing
    oConexion.Close
    Set oConexion = Nothing
    
ManejadorDeError:
  'ms_MensajeError = Err.Description
End Function
