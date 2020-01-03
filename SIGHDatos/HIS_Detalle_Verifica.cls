VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_Detalle_Verifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_Detalle_Verifica
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------

Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla HIS_Detalle_Verifica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_Detalle_Verifica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_Detalle_VerificaAgregar"
   Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdHisDetalle = 0, Null, oTabla.IdHisDetalle))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdHisCabecera = 0, Null, oTabla.IdHisCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAtencion = 0, Null, oTabla.IdTipoAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DiaAtencion", adInteger, adParamInput, 0, IIf(oTabla.DiaAtencion = 0, Null, oTabla.DiaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Sexo", adInteger, adParamInput, 0, IIf(oTabla.Sexo = 0, Null, oTabla.Sexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdNacionalidad", adInteger, adParamInput, 0, IIf(oTabla.IdNacionalidad = 0, Null, oTabla.IdNacionalidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocIdentidad", adVarChar, adParamInput, 12, IIf(oTabla.NroDocIdentidad = "", Null, oTabla.NroDocIdentidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHijo", adChar, adParamInput, 2, IIf(oTabla.NroHijo = "", Null, oTabla.NroHijo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEtnia", adChar, adParamInput, 2, IIf(oTabla.IdEtnia = "", Null, oTabla.IdEtnia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoDocumento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoDocumento = 0, Null, oTabla.IdTipoDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHC_FF", adVarChar, adParamInput, 20, IIf(oTabla.NroHC_FF = "", Null, oTabla.NroHC_FF))
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
   Set oParameter = .CreateParameter("@Talla", adInteger, adParamInput, 0, IIf(oTabla.Talla = 0, Null, oTabla.Talla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Peso", adChar, adParamInput, 50, IIf(oTabla.Peso = "", Null, oTabla.Peso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoaEstablec", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoaEstablec = 0, Null, oTabla.IdEstadoaEstablec))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoaServicio", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoaServicio = 0, Null, oTabla.IdEstadoaServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroRegistroLote", adInteger, adParamInput, 0, IIf(oTabla.NroRegistroLote = 0, Null, oTabla.NroRegistroLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroRegistroHoja", adInteger, adParamInput, 0, IIf(oTabla.NroRegistroHoja = 0, Null, oTabla.NroRegistroHoja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Registrado", adInteger, adParamInput, 0, oTabla.Registrado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Coincide", adInteger, adParamInput, 0, oTabla.Coincide)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
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
'   Descripción:    modifica un registro a la tabla HIS_Detalle_Verifica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_Detalle_Verifica) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_Detalle_VerificaModificar"
   Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, IIf(oTabla.IdHisDetalle = 0, Null, oTabla.IdHisDetalle))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdHisCabecera = 0, Null, oTabla.IdHisCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdTipoAtencion = 0, Null, oTabla.IdTipoAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DiaAtencion", adInteger, adParamInput, 0, IIf(oTabla.DiaAtencion = 0, Null, oTabla.DiaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Sexo", adInteger, adParamInput, 0, IIf(oTabla.Sexo = 0, Null, oTabla.Sexo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdNacionalidad", adInteger, adParamInput, 0, IIf(oTabla.IdNacionalidad = 0, Null, oTabla.IdNacionalidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroDocIdentidad", adVarChar, adParamInput, 12, IIf(oTabla.NroDocIdentidad = "", Null, oTabla.NroDocIdentidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHijo", adChar, adParamInput, 2, IIf(oTabla.NroHijo = "", Null, oTabla.NroHijo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEtnia", adChar, adParamInput, 2, IIf(oTabla.IdEtnia = "", Null, oTabla.IdEtnia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoDocumento", adInteger, adParamInput, 0, IIf(oTabla.IdTipoDocumento = 0, Null, oTabla.IdTipoDocumento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHC_FF", adVarChar, adParamInput, 20, IIf(oTabla.NroHC_FF = "", Null, oTabla.NroHC_FF))
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
   Set oParameter = .CreateParameter("@Talla", adInteger, adParamInput, 0, IIf(oTabla.Talla = 0, Null, oTabla.Talla))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Peso", adChar, adParamInput, 50, IIf(oTabla.Peso = "", Null, oTabla.Peso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoaEstablec", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoaEstablec = 0, Null, oTabla.IdEstadoaEstablec))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoaServicio", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoaServicio = 0, Null, oTabla.IdEstadoaServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroRegistroLote", adInteger, adParamInput, 0, IIf(oTabla.NroRegistroLote = 0, Null, oTabla.NroRegistroLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroRegistroHoja", adInteger, adParamInput, 0, IIf(oTabla.NroRegistroHoja = 0, Null, oTabla.NroRegistroHoja))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Registrado", adInteger, adParamInput, 0, oTabla.Registrado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Coincide", adInteger, adParamInput, 0, oTabla.Coincide)
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
'   Descripción:    elimia o borra un registro a la tabla HIS_Detalle_Verifica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal ml_IdHisLote As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_Detalle_VerificaEliminar"
   Set oParameter = .CreateParameter("@IdHisLote", adInteger, adParamInput, 0, ml_IdHisLote)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, 1)
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
'   Descripción:    Selecciona un unico registro de la tabla HIS_Detalle_Verifica
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_Detalle_Verifica) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_Detalle_VerificaSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, oTabla.IdHisDetalle)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdHisDetalle = IIf(IsNull(oRecordset!IdHisDetalle), 0, oRecordset!IdHisDetalle)
       oTabla.IdHisCabecera = IIf(IsNull(oRecordset!IdHisCabecera), 0, oRecordset!IdHisCabecera)
       oTabla.IdTipoAtencion = IIf(IsNull(oRecordset!IdTipoAtencion), 0, oRecordset!IdTipoAtencion)
       oTabla.DiaAtencion = IIf(IsNull(oRecordset!DiaAtencion), 0, oRecordset!DiaAtencion)
       oTabla.Sexo = IIf(IsNull(oRecordset!Sexo), 0, oRecordset!Sexo)
       oTabla.IdNacionalidad = IIf(IsNull(oRecordset!IdNacionalidad), 0, oRecordset!IdNacionalidad)
       oTabla.NroDocIdentidad = IIf(IsNull(oRecordset!NroDocIdentidad), "", oRecordset!NroDocIdentidad)
       oTabla.NroHijo = IIf(IsNull(oRecordset!NroHijo), "", oRecordset!NroHijo)
       oTabla.IdEtnia = IIf(IsNull(oRecordset!IdEtnia), "", oRecordset!IdEtnia)
       oTabla.IdTipoDocumento = IIf(IsNull(oRecordset!IdTipoDocumento), 0, oRecordset!IdTipoDocumento)
       oTabla.NroHC_FF = IIf(IsNull(oRecordset!NroHC_FF), "", oRecordset!NroHC_FF)
       oTabla.CodigoActividad = IIf(IsNull(oRecordset!CodigoActividad), "", oRecordset!CodigoActividad)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.IdDistrito = IIf(IsNull(oRecordset!IdDistrito), 0, oRecordset!IdDistrito)
       oTabla.IdTipoEdad = IIf(IsNull(oRecordset!IdTipoEdad), 0, oRecordset!IdTipoEdad)
       oTabla.Edad = IIf(IsNull(oRecordset!Edad), 0, oRecordset!Edad)
       oTabla.Talla = IIf(IsNull(oRecordset!Talla), 0, oRecordset!Talla)
       oTabla.Peso = IIf(IsNull(oRecordset!Peso), "", oRecordset!Peso)
       oTabla.IdEstadoaEstablec = IIf(IsNull(oRecordset!IdEstadoaEstablec), 0, oRecordset!IdEstadoaEstablec)
       oTabla.IdEstadoaServicio = IIf(IsNull(oRecordset!IdEstadoaServicio), 0, oRecordset!IdEstadoaServicio)
       oTabla.NroRegistroLote = IIf(IsNull(oRecordset!NroRegistroLote), 0, oRecordset!NroRegistroLote)
       oTabla.NroRegistroHoja = IIf(IsNull(oRecordset!NroRegistroHoja), 0, oRecordset!NroRegistroHoja)
       oTabla.Registrado = IIf(IsNull(oRecordset!Registrado), 0, oRecordset!Registrado)
       oTabla.Coincide = IIf(IsNull(oRecordset!Coincide), 0, oRecordset!Coincide)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_Detalle_Verifica)
   oTabla.IdHisDetalle = 0
   oTabla.IdHisCabecera = 0
   oTabla.IdTipoAtencion = 0
   oTabla.DiaAtencion = 0
   oTabla.Sexo = 0
   oTabla.IdNacionalidad = 0
   oTabla.NroDocIdentidad = ""
   oTabla.NroHijo = ""
   oTabla.IdEtnia = ""
   oTabla.IdTipoDocumento = 0
   oTabla.NroHC_FF = ""
   oTabla.CodigoActividad = ""
   oTabla.idTipoFinanciamiento = 0
   oTabla.IdDistrito = 0
   oTabla.IdTipoEdad = 0
   oTabla.Edad = 0
   oTabla.Talla = 0
   oTabla.Peso = ""
   oTabla.IdEstadoaEstablec = 0
   oTabla.IdEstadoaServicio = 0
   oTabla.NroRegistroLote = 0
   oTabla.NroRegistroHoja = 0
   oTabla.Registrado = 0
   oTabla.Coincide = 0
End Sub

Function HIS_ConsultarRegistroDetalleHis(lnIdHisDetalle As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
     Set HIS_ConsultarRegistroDetalleHis = Nothing
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandType = adCmdStoredProc
        .CommandText = "HIS_ConsultarRegistroDetalleHis"
        Set oParameter = .CreateParameter("@IdHisDetalle", adInteger, adParamInput, 0, lnIdHisDetalle): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
    Set HIS_ConsultarRegistroDetalleHis = oRecordset
    ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

