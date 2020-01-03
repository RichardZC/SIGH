VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "m_eess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: m_eess
'        Autor: Daniel Enrique Barrantes
'        Fecha: 06/12/2013 12:51:59 p.m.
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
'   Descripción:    Inserta un registro a la tabla m_eess
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As Dom_eess) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "m_eessAgregar"
   Set oParameter = .CreateParameter("@pre_IdEESS", adVarChar, adParamInput, 10, IIf(oTabla.pre_IdEESS = "", Null, oTabla.pre_IdEESS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_Nombre", adVarChar, adParamInput, 150, IIf(oTabla.pre_Nombre = "", Null, oTabla.pre_Nombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_Afilia", adVarChar, adParamInput, 1, IIf(oTabla.pre_Afilia = "", Null, oTabla.pre_Afilia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_UCI", adVarChar, adParamInput, 1, IIf(oTabla.pre_UCI = "", Null, oTabla.pre_UCI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_IdCategoriaEESS", adVarChar, adParamInput, 2, IIf(oTabla.pre_IdCategoriaEESS = "", Null, oTabla.pre_IdCategoriaEESS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_IdDisa", adVarChar, adParamInput, 3, IIf(oTabla.pre_IdDisa = "", Null, oTabla.pre_IdDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_IdOdsis", adVarChar, adParamInput, 3, IIf(oTabla.pre_IdOdsis = "", Null, oTabla.pre_IdOdsis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_IdUbigeo", adVarChar, adParamInput, 6, IIf(oTabla.pre_IdUbigeo = "", Null, oTabla.pre_IdUbigeo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_CodEjeAdm", adVarChar, adParamInput, 4, IIf(oTabla.pre_CodEjeAdm = "", Null, oTabla.pre_CodEjeAdm))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_Vrae", adVarChar, adParamInput, 1, IIf(oTabla.pre_Vrae = "", Null, oTabla.pre_Vrae))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_Umbral", adVarChar, adParamInput, 1, IIf(oTabla.pre_Umbral = "", Null, oTabla.pre_Umbral))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_Aisped", adVarChar, adParamInput, 1, IIf(oTabla.pre_Aisped = "", Null, oTabla.pre_Aisped))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_esmn", adVarChar, adParamInput, 50, IIf(oTabla.pre_esmn = "", Null, oTabla.pre_esmn))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_IdEstado", adVarChar, adParamInput, 50, IIf(oTabla.pre_IdEstado = "", Null, oTabla.pre_IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_CodigoRENAES", adVarChar, adParamInput, 50, IIf(oTabla.pre_CodigoRENAES = "", Null, oTabla.pre_CodigoRENAES))
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
   If Err.Number = -2147217873 Then
      ms_MensajeError = ""
   End If
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla m_eess
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As Dom_eess) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "m_eessModificar"
   Set oParameter = .CreateParameter("@pre_IdEESS", adVarChar, adParamInput, 10, IIf(oTabla.pre_IdEESS = "", Null, oTabla.pre_IdEESS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_Nombre", adVarChar, adParamInput, 150, IIf(oTabla.pre_Nombre = "", Null, oTabla.pre_Nombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_Afilia", adVarChar, adParamInput, 1, IIf(oTabla.pre_Afilia = "", Null, oTabla.pre_Afilia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_UCI", adVarChar, adParamInput, 1, IIf(oTabla.pre_UCI = "", Null, oTabla.pre_UCI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_IdCategoriaEESS", adVarChar, adParamInput, 2, IIf(oTabla.pre_IdCategoriaEESS = "", Null, oTabla.pre_IdCategoriaEESS))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_IdDisa", adVarChar, adParamInput, 3, IIf(oTabla.pre_IdDisa = "", Null, oTabla.pre_IdDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_IdOdsis", adVarChar, adParamInput, 3, IIf(oTabla.pre_IdOdsis = "", Null, oTabla.pre_IdOdsis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_IdUbigeo", adVarChar, adParamInput, 6, IIf(oTabla.pre_IdUbigeo = "", Null, oTabla.pre_IdUbigeo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_CodEjeAdm", adVarChar, adParamInput, 4, IIf(oTabla.pre_CodEjeAdm = "", Null, oTabla.pre_CodEjeAdm))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_Vrae", adVarChar, adParamInput, 1, IIf(oTabla.pre_Vrae = "", Null, oTabla.pre_Vrae))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_Umbral", adVarChar, adParamInput, 1, IIf(oTabla.pre_Umbral = "", Null, oTabla.pre_Umbral))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_Aisped", adVarChar, adParamInput, 1, IIf(oTabla.pre_Aisped = "", Null, oTabla.pre_Aisped))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_esmn", adVarChar, adParamInput, 50, IIf(oTabla.pre_esmn = "", Null, oTabla.pre_esmn))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_IdEstado", adVarChar, adParamInput, 50, IIf(oTabla.pre_IdEstado = "", Null, oTabla.pre_IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@pre_CodigoRENAES", adVarChar, adParamInput, 50, IIf(oTabla.pre_CodigoRENAES = "", Null, oTabla.pre_CodigoRENAES))
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
'   Descripción:    elimia o borra un registro a la tabla m_eess
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As Dom_eess) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "m_eessEliminar"
   Set oParameter = .CreateParameter("@pre_IdEESS", adVarChar, adParamInput, 10, oTabla.pre_IdEESS)
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
'   Descripción:    Selecciona un unico registro de la tabla m_eess
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As Dom_eess) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "m_eessSeleccionarPorId"
     Set oParameter = .CreateParameter("@pre_IdEESS", adVarChar, adParamInput, 10, oTabla.pre_IdEESS)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.pre_IdEESS = IIf(IsNull(oRecordset!pre_IdEESS), "", oRecordset!pre_IdEESS)
       oTabla.pre_Nombre = IIf(IsNull(oRecordset!pre_Nombre), "", oRecordset!pre_Nombre)
       oTabla.pre_Afilia = IIf(IsNull(oRecordset!pre_Afilia), "", oRecordset!pre_Afilia)
       oTabla.pre_UCI = IIf(IsNull(oRecordset!pre_UCI), "", oRecordset!pre_UCI)
       oTabla.pre_IdCategoriaEESS = IIf(IsNull(oRecordset!pre_IdCategoriaEESS), "", oRecordset!pre_IdCategoriaEESS)
       oTabla.pre_IdDisa = IIf(IsNull(oRecordset!pre_IdDisa), "", oRecordset!pre_IdDisa)
       oTabla.pre_IdOdsis = IIf(IsNull(oRecordset!pre_IdOdsis), "", oRecordset!pre_IdOdsis)
       oTabla.pre_IdUbigeo = IIf(IsNull(oRecordset!pre_IdUbigeo), "", oRecordset!pre_IdUbigeo)
       oTabla.pre_CodEjeAdm = IIf(IsNull(oRecordset!pre_CodEjeAdm), "", oRecordset!pre_CodEjeAdm)
       oTabla.pre_Vrae = IIf(IsNull(oRecordset!pre_Vrae), "", oRecordset!pre_Vrae)
       oTabla.pre_Umbral = IIf(IsNull(oRecordset!pre_Umbral), "", oRecordset!pre_Umbral)
       oTabla.pre_Aisped = IIf(IsNull(oRecordset!pre_Aisped), "", oRecordset!pre_Aisped)
       oTabla.pre_esmn = IIf(IsNull(oRecordset!pre_esmn), "", oRecordset!pre_esmn)
       oTabla.pre_IdEstado = IIf(IsNull(oRecordset!pre_IdEstado), "", oRecordset!pre_IdEstado)
       oTabla.pre_CodigoRENAES = IIf(IsNull(oRecordset!pre_CodigoRENAES), "", oRecordset!pre_CodigoRENAES)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As Dom_eess)
   oTabla.pre_IdEESS = ""
   oTabla.pre_Nombre = ""
   oTabla.pre_Afilia = ""
   oTabla.pre_UCI = ""
   oTabla.pre_IdCategoriaEESS = ""
   oTabla.pre_IdDisa = ""
   oTabla.pre_IdOdsis = ""
   oTabla.pre_IdUbigeo = ""
   oTabla.pre_CodEjeAdm = ""
   oTabla.pre_Vrae = ""
   oTabla.pre_Umbral = ""
   oTabla.pre_Aisped = ""
   oTabla.pre_esmn = ""
   oTabla.pre_IdEstado = ""
   oTabla.pre_CodigoRENAES = ""
End Sub
'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: m_eess
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: m_eess
'        Autor:   DAniel Barrantes B.
'        Fecha: 06/12/2013 12:51:59 p.m.
'------------------------------------------------------------------------------------
 

