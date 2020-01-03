VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_presentaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: a_presentaciones
'        Autor: Daniel Enrique Barrantes
'        Fecha: 06/12/2013 12:52:27 p.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

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
'   Descripción:    Inserta un registro a la tabla a_presentaciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As Doa_presentaciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "a_presentacionesAgregar"
   Set oParameter = .CreateParameter("@tpre_IdPresentacion", adVarChar, adParamInput, 3, IIf(oTabla.tpre_IdPresentacion = "", Null, oTabla.tpre_IdPresentacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_Descripcion", adVarChar, adParamInput, 250, IIf(oTabla.tpre_Descripcion = "", Null, oTabla.tpre_Descripcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_Abreviatura", adVarChar, adParamInput, 15, IIf(oTabla.tpre_Abreviatura = "", Null, oTabla.tpre_Abreviatura))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_TopeMinimo", adInteger, adParamInput, 0, IIf(oTabla.tpre_TopeMinimo = 0, Null, oTabla.tpre_TopeMinimo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_TopeNoHosp", adInteger, adParamInput, 0, IIf(oTabla.tpre_TopeNoHosp = 0, Null, oTabla.tpre_TopeNoHosp))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_TopeHosp", adInteger, adParamInput, 0, IIf(oTabla.tpre_TopeHosp = 0, Null, oTabla.tpre_TopeHosp))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_IdEstado", adVarWChar, adParamInput, 1, oTabla.tpre_IdEstado)
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
'   Descripción:    modifica un registro a la tabla a_presentaciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As Doa_presentaciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "a_presentacionesModificar"
   Set oParameter = .CreateParameter("@tpre_IdPresentacion", adVarChar, adParamInput, 3, IIf(oTabla.tpre_IdPresentacion = "", Null, oTabla.tpre_IdPresentacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_Descripcion", adVarChar, adParamInput, 250, IIf(oTabla.tpre_Descripcion = "", Null, oTabla.tpre_Descripcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_Abreviatura", adVarChar, adParamInput, 15, IIf(oTabla.tpre_Abreviatura = "", Null, oTabla.tpre_Abreviatura))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_TopeMinimo", adInteger, adParamInput, 0, IIf(oTabla.tpre_TopeMinimo = 0, Null, oTabla.tpre_TopeMinimo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_TopeNoHosp", adInteger, adParamInput, 0, IIf(oTabla.tpre_TopeNoHosp = 0, Null, oTabla.tpre_TopeNoHosp))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_TopeHosp", adInteger, adParamInput, 0, IIf(oTabla.tpre_TopeHosp = 0, Null, oTabla.tpre_TopeHosp))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@tpre_IdEstado", adVarWChar, adParamInput, 1, oTabla.tpre_IdEstado)
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
'   Descripción:    elimia o borra un registro a la tabla a_presentaciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As Doa_presentaciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "a_presentacionesEliminar"
   Set oParameter = .CreateParameter("@tpre_IdPresentacion", adVarChar, adParamInput, 3, oTabla.tpre_IdPresentacion)
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
'   Descripción:    Selecciona un unico registro de la tabla a_presentaciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As Doa_presentaciones) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "a_presentacionesSeleccionarPorId"
     Set oParameter = .CreateParameter("@tpre_IdPresentacion", adVarChar, adParamInput, 3, oTabla.tpre_IdPresentacion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.tpre_IdPresentacion = IIf(IsNull(oRecordset!tpre_IdPresentacion), "", oRecordset!tpre_IdPresentacion)
       oTabla.tpre_Descripcion = IIf(IsNull(oRecordset!tpre_Descripcion), "", oRecordset!tpre_Descripcion)
       oTabla.tpre_Abreviatura = IIf(IsNull(oRecordset!tpre_Abreviatura), "", oRecordset!tpre_Abreviatura)
       oTabla.tpre_TopeMinimo = IIf(IsNull(oRecordset!tpre_TopeMinimo), 0, oRecordset!tpre_TopeMinimo)
       oTabla.tpre_TopeNoHosp = IIf(IsNull(oRecordset!tpre_TopeNoHosp), 0, oRecordset!tpre_TopeNoHosp)
       oTabla.tpre_TopeHosp = IIf(IsNull(oRecordset!tpre_TopeHosp), 0, oRecordset!tpre_TopeHosp)
       oTabla.tpre_IdEstado = oRecordset!tpre_IdEstado
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As Doa_presentaciones)
   oTabla.tpre_IdPresentacion = ""
   oTabla.tpre_Descripcion = ""
   oTabla.tpre_Abreviatura = ""
   oTabla.tpre_TopeMinimo = 0
   oTabla.tpre_TopeNoHosp = 0
   oTabla.tpre_TopeHosp = 0
   oTabla.tpre_IdEstado = ""
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: a_presentaciones
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: a_presentaciones
'        Autor:   DAniel Barrantes B.
'        Fecha: 06/12/2013 12:52:27 p.m.
'------------------------------------------------------------------------------------
 

