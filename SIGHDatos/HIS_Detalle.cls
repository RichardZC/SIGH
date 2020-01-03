VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_Detalle"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_Detalle
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla HIS_Detalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_Detalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DetalleAgregar"
   Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdHisCabecera = 0, Null, oTabla.IdHisCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAtencion = 0, Null, oTabla.IdTipoAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DiaAtencion", adInteger, adParamInput, 0, IIf(oTabla.DiaAtencion = 0, Null, oTabla.DiaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdHisPaciente = 0, Null, oTabla.IdHisPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoActividad", adVarChar, adParamInput, 10, IIf(oTabla.CodigoActividad = "", Null, oTabla.CodigoActividad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, IIf(oTabla.IdDistrito = 0, Null, oTabla.IdDistrito))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoEdad", adInteger, adParamInput, 0, IIf(oTabla.IdTipoEdad = 0, Null, oTabla.IdTipoEdad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 0, IIf(oTabla.Edad = 0, Null, oTabla.Edad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Talla", adInteger, adParamInput, 0, IIf(oTabla.Talla = 0, 0, oTabla.Talla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Peso", adChar, adParamInput, 50, IIf(oTabla.Peso = 0, 0, oTabla.Peso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoaEstablec", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoaEstablec = 0, Null, oTabla.IdEstadoaEstablec))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoaServicio", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoaServicio = 0, Null, oTabla.IdEstadoaServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroRegistroLote", adInteger, adParamInput, 0, IIf(oTabla.NroRegistroLote = 0, Null, oTabla.NroRegistroLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroRegistroHoja", adInteger, adParamInput, 0, IIf(oTabla.NroRegistroHoja = 0, Null, oTabla.NroRegistroHoja))
   .Parameters.Append oParameter
   
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdHisDetalle = .Parameters("@IdHisDetalle")
End With
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla HIS_Detalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_Detalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DetalleModificar"
   Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 4, IIf(oTabla.IdHisDetalle = 0, Null, oTabla.IdHisDetalle))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisCabecera", adInteger, adParamInput, 4, IIf(oTabla.IdHisCabecera = 0, Null, oTabla.IdHisCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoAtencion", adInteger, adParamInput, 4, IIf(oTabla.IdTipoAtencion = 0, Null, oTabla.IdTipoAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DiaAtencion", adInteger, adParamInput, 4, IIf(oTabla.DiaAtencion = 0, Null, oTabla.DiaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisPaciente", adInteger, adParamInput, 4, IIf(oTabla.IdHisPaciente = 0, Null, oTabla.IdHisPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoActividad", adVarChar, adParamInput, 10, IIf(oTabla.CodigoActividad = "", Null, oTabla.CodigoActividad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 4, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdDistrito", adInteger, adParamInput, 0, IIf(oTabla.IdDistrito = 0, Null, oTabla.IdDistrito))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoEdad", adInteger, adParamInput, 4, IIf(oTabla.IdTipoEdad = 0, Null, oTabla.IdTipoEdad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Edad", adInteger, adParamInput, 4, IIf(oTabla.Edad = 0, Null, oTabla.Edad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Talla", adInteger, adParamInput, 4, IIf(oTabla.Talla = 0, Null, oTabla.Talla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Peso", adVarChar, adParamInput, 50, IIf(oTabla.Peso = "", Null, oTabla.Peso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoaEstablec", adInteger, adParamInput, 4, IIf(oTabla.IdEstadoaEstablec = 0, Null, oTabla.IdEstadoaEstablec))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoaServicio", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoaServicio = 0, Null, oTabla.IdEstadoaServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroRegistroLote", adInteger, adParamInput, 0, IIf(oTabla.NroRegistroLote = 0, Null, oTabla.NroRegistroLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroRegistroHoja", adInteger, adParamInput, 0, IIf(oTabla.NroRegistroHoja = 0, Null, oTabla.NroRegistroHoja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Modificar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla HIS_Detalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHIS_Detalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_DetalleEliminar"
   Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, oTabla.IdHisDetalle)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Eliminar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla HIS_Detalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_Detalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_DetalleSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, oTabla.IdHisDetalle)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdHisDetalle = IIf(IsNull(oRecordset!IdHisDetalle), 0, oRecordset!IdHisDetalle)
       oTabla.IdHisCabecera = IIf(IsNull(oRecordset!IdHisCabecera), 0, oRecordset!IdHisCabecera)
       oTabla.IdTipoAtencion = IIf(IsNull(oRecordset!IdTipoAtencion), 0, oRecordset!IdTipoAtencion)
       oTabla.DiaAtencion = IIf(IsNull(oRecordset!DiaAtencion), 0, oRecordset!DiaAtencion)
       oTabla.IdHisPaciente = IIf(IsNull(oRecordset!IdHisPaciente), 0, oRecordset!IdHisPaciente)
       oTabla.CodigoActividad = IIf(IsNull(oRecordset!CodigoActividad), "", oRecordset!CodigoActividad)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.IdDistrito = IIf(IsNull(oRecordset!IdDistrito), 0, oRecordset!IdDistrito)
       oTabla.IdTipoEdad = IIf(IsNull(oRecordset!IdTipoEdad), 0, oRecordset!IdTipoEdad)
       oTabla.Edad = IIf(IsNull(oRecordset!Edad), 0, oRecordset!Edad)
       oTabla.Talla = IIf(IsNull(oRecordset!Talla), 0, oRecordset!Talla)
       oTabla.Peso = IIf(IsNull(oRecordset!Peso), "", oRecordset!Peso)
       oTabla.IdEstadoaEstablec = IIf(IsNull(oRecordset!IdEstadoaEstablec), 0, oRecordset!IdEstadoaEstablec)
       oTabla.IdEstadoaServicio = IIf(IsNull(oRecordset!IdEstadoaServicio), 0, oRecordset!IdEstadoaServicio)
       oTabla.NroRegistroHoja = IIf(IsNull(oRecordset!NroRegistroHoja), 0, oRecordset!NroRegistroHoja)
       oTabla.NroRegistroLote = IIf(IsNull(oRecordset!NroRegistroLote), 0, oRecordset!NroRegistroLote)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_Detalle)
   oTabla.IdHisDetalle = 0
   oTabla.IdHisCabecera = 0
   oTabla.IdTipoAtencion = 0
   oTabla.DiaAtencion = 0
   oTabla.IdHisPaciente = 0
   oTabla.CodigoActividad = ""
   oTabla.idTipoFinanciamiento = 0
   oTabla.IdDistrito = 0
   oTabla.IdTipoEdad = 0
   oTabla.Edad = 0
   oTabla.Talla = 0
   oTabla.Peso = ""
   oTabla.IdEstadoaEstablec = 0
   oTabla.IdEstadoaServicio = 0
   oTabla.NroRegistroHoja = 0
   oTabla.NroRegistroLote = 0
End Sub

Function ListaIdsAtencionesPorIdCabecera(RegCabeceraHIS As DOHIS_Cabecera) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "his_detallePorIdCabecera"
        Set oParameter = .CreateParameter("@IdHisCabecera", adInteger, adParamInput, 0, RegCabeceraHIS.IdHisCabecera): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListaIdsAtencionesPorIdCabecera = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerDatosDetalleAtencion(ml_IdCabeceraHIS As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_detalleObtenerDatosDetalleAtencion"
        Set oParameter = .CreateParameter("@ml_IdCabeceraHIS", adInteger, adParamInput, 0, ml_IdCabeceraHIS): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerDatosDetalleAtencion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Public Function ExportacionHIS_Atenciones(IdUsuario As Long, ml_Mes As Long, mi_anio As Long) As Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String

sSql = ""

With oCommand
    .CommandType = adCmdText
    Set .ActiveConnection = mo_Conexion
    .CommandText = sSql
    Set oRecordset = .Execute
    Set oRecordset.ActiveConnection = Nothing
End With

Set ExportacionHIS_Atenciones = oRecordset

ms_MensajeError = ""
Exit Function
ManejadorDeError:
       ms_MensajeError = Err.Number & " " + Err.Description
End Function



