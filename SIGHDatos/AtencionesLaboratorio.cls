VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesLaboratorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de c�digo autogenerado para la clase: AtencionesLabSinDxSinCpt
'        Autor: Daniel Enrique Barrantes
'        Fecha: 01/06/2018 10:17:15 a.m.
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
'   Funci�n:        Insertar
'   Descripci�n:    Inserta un registro a la tabla AtencionesLabSinDxSinCpt
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOAtencionesLaboratorio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesLabSinDxSinCptAgregar"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Grupo", adInteger, adParamInput, 0, IIf(oTabla.Grupo = 0, Null, oTabla.Grupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SubGrupo", adInteger, adParamInput, 0, IIf(oTabla.SubGrupo = 0, Null, oTabla.SubGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@lab", adVarChar, adParamInput, 3, IIf(oTabla.lab = "", Null, oTabla.lab))
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
'   Funci�n:        Modificar
'   Descripci�n:    modifica un registro a la tabla AtencionesLabSinDxSinCpt
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionesLaboratorio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesLabSinDxSinCptModificar"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Grupo", adInteger, adParamInput, 0, IIf(oTabla.Grupo = 0, Null, oTabla.Grupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SubGrupo", adInteger, adParamInput, 0, IIf(oTabla.SubGrupo = 0, Null, oTabla.SubGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@lab", adVarChar, adParamInput, 3, IIf(oTabla.lab = "", Null, oTabla.lab))
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
'   Funci�n:        Eliminar
'   Descripci�n:    elimia o borra un registro a la tabla AtencionesLabSinDxSinCpt
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionesLaboratorio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "AtencionesLabSinDxSinCptEliminar"
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
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
'   Funci�n:        SeleccionarPorId
'   Descripci�n:    Selecciona un unico registro de la tabla AtencionesLabSinDxSinCpt
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionesLaboratorio) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesLabSinDxSinCptSeleccionarPorId"
     Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, oTabla.idAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.Grupo = IIf(IsNull(oRecordset!Grupo), 0, oRecordset!Grupo)
       oTabla.SubGrupo = IIf(IsNull(oRecordset!SubGrupo), 0, oRecordset!SubGrupo)
       oTabla.lab = IIf(IsNull(oRecordset!lab), "", oRecordset!lab)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOAtencionesLaboratorio)
   oTabla.idAtencion = 0
   oTabla.Grupo = 0
   oTabla.SubGrupo = 0
   oTabla.lab = ""
End Sub
'------------------------------------------------------------------------------------
'        Fin de c�digo autogenerado para la clase: AtencionesLabSinDxSinCpt
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de c�digo manual para la clase: AtencionesLabSinDxSinCpt
'        Autor:   DAniel Barrantes B.
'        Fecha: 01/06/2018 10:17:15 a.m.
'------------------------------------------------------------------------------------
 

