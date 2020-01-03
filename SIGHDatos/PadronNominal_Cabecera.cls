VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PadronNominal_Cabecera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla PadronNominal_Cabecera
'        Programado por: Palomino Y
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla PadronNominal_Cabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOPadronNominal_Cabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PadronNominal_CabeceraAgregar"
   Set oParameter = .CreateParameter("@IdPaNomCabecera", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdResponsableAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdResponsableAtencion = 0, Null, oTabla.IdResponsableAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCodigoRenaes", adInteger, adParamInput, 0, IIf(oTabla.IdCodigoRenaes = 0, Null, oTabla.IdCodigoRenaes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, IIf(oTabla.Mes = 0, Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Año", adChar, adParamInput, 4, IIf(oTabla.Año = "", Null, oTabla.Año))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdPaNomCabecera = .Parameters("@IdPaNomCabecera")
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
'   Descripción:    modifica un registro a la tabla PadronNominal_Cabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOPadronNominal_Cabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PadronNominal_CabeceraModificar"
   Set oParameter = .CreateParameter("@IdPaNomCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdPaNomCabecera = 0, Null, oTabla.IdPaNomCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdResponsableAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdResponsableAtencion = 0, Null, oTabla.IdResponsableAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdCodigoRenaes", adInteger, adParamInput, 0, IIf(oTabla.IdCodigoRenaes = 0, Null, oTabla.IdCodigoRenaes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, IIf(oTabla.Mes = 0, Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Año", adChar, adParamInput, 4, IIf(oTabla.Año = "", Null, oTabla.Año))
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
'   Descripción:    elimia o borra un registro a la tabla PadronNominal_Cabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOPadronNominal_Cabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PadronNominal_CabeceraEliminar"
   Set oParameter = .CreateParameter("@IdPaNomCabecera", adInteger, adParamInput, 0, oTabla.IdPaNomCabecera)
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
'   Descripción:    Selecciona un unico registro de la tabla PadronNominal_Cabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOPadronNominal_Cabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PadronNominal_CabeceraSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdPaNomCabecera", adInteger, adParamInput, 0, oTabla.IdPaNomCabecera)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdPaNomCabecera = IIf(IsNull(oRecordset!IdPaNomCabecera), 0, oRecordset!IdPaNomCabecera)
       oTabla.IdResponsableAtencion = IIf(IsNull(oRecordset!IdResponsableAtencion), 0, oRecordset!IdResponsableAtencion)
       oTabla.IdCodigoRenaes = IIf(IsNull(oRecordset!IdCodigoRenaes), 0, oRecordset!IdCodigoRenaes)
       oTabla.Mes = IIf(IsNull(oRecordset!Mes), 0, oRecordset!Mes)
       oTabla.Año = IIf(IsNull(oRecordset!Año), "", oRecordset!Año)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOPadronNominal_Cabecera)
   oTabla.IdPaNomCabecera = 0
   oTabla.IdResponsableAtencion = 0
   oTabla.IdCodigoRenaes = 0
   oTabla.Mes = 0
   oTabla.Año = ""
End Sub


