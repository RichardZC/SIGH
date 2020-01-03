VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SisFuaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla SisFuaUsuario
'        Programado por: Barrantes D
'        Fecha: Marzo 2014
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
'   Descripción:    Inserta un registro a la tabla SisFuaUsuario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOSisFuaUsuario) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaUsuarioAgregar"
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 8, IIf(oTabla.DNI = "", Null, Trim(oTabla.DNI)))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoDoc", adVarChar, adParamInput, 2, IIf(oTabla.TipoDoc = "", Null, oTabla.TipoDoc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoPat", adVarChar, adParamInput, 20, IIf(oTabla.ApellidoPat = "", Null, oTabla.ApellidoPat))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoMat", adVarChar, adParamInput, 20, IIf(oTabla.ApellidoMat = "", Null, oTabla.ApellidoMat))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 20, IIf(oTabla.PrimerNombre = "", Null, oTabla.PrimerNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 20, IIf(oTabla.SegundoNombre = "", Null, oTabla.SegundoNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroEnvio", adInteger, adParamInput, 0, IIf(oTabla.NroEnvio = 0, Null, oTabla.NroEnvio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Periodo", adVarChar, adParamInput, 4, IIf(oTabla.Periodo = "", Null, oTabla.Periodo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adVarChar, adParamInput, 2, IIf(oTabla.Mes = "", Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoEstablecimiento", adVarChar, adParamInput, 10, IIf(oTabla.CodigoEstablecimiento = "", Null, oTabla.CodigoEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdUsuario = .Parameters("@idUsuario")
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
'   Descripción:    modifica un registro a la tabla SisFuaUsuario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOSisFuaUsuario) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaUsuarioModificar"
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 8, IIf(oTabla.DNI = "", Null, oTabla.DNI))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoDoc", adVarChar, adParamInput, 2, IIf(oTabla.TipoDoc = "", Null, oTabla.TipoDoc))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoPat", adVarChar, adParamInput, 20, IIf(oTabla.ApellidoPat = "", Null, oTabla.ApellidoPat))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ApellidoMat", adVarChar, adParamInput, 20, IIf(oTabla.ApellidoMat = "", Null, oTabla.ApellidoMat))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrimerNombre", adVarChar, adParamInput, 20, IIf(oTabla.PrimerNombre = "", Null, oTabla.PrimerNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SegundoNombre", adVarChar, adParamInput, 20, IIf(oTabla.SegundoNombre = "", Null, oTabla.SegundoNombre))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroEnvio", adInteger, adParamInput, 0, IIf(oTabla.NroEnvio = 0, Null, oTabla.NroEnvio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Periodo", adVarChar, adParamInput, 4, IIf(oTabla.Periodo = "", Null, oTabla.Periodo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adVarChar, adParamInput, 2, IIf(oTabla.Mes = "", Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CodigoEstablecimiento", adVarChar, adParamInput, 10, IIf(oTabla.CodigoEstablecimiento = "", Null, oTabla.CodigoEstablecimiento))
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
'   Descripción:    elimia o borra un registro a la tabla SisFuaUsuario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOSisFuaUsuario) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaUsuarioEliminar"
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, oTabla.IdUsuario)
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
'   Descripción:    Selecciona un unico registro de la tabla SisFuaUsuario
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOSisFuaUsuario) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SisFuaUsuarioSeleccionarPorId"
     Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, oTabla.IdUsuario)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.DNI = IIf(IsNull(oRecordset!DNI), "", oRecordset!DNI)
       oTabla.TipoDoc = IIf(IsNull(oRecordset!TipoDoc), "", oRecordset!TipoDoc)
       oTabla.ApellidoPat = IIf(IsNull(oRecordset!ApellidoPat), "", oRecordset!ApellidoPat)
       oTabla.ApellidoMat = IIf(IsNull(oRecordset!ApellidoMat), "", oRecordset!ApellidoMat)
       oTabla.PrimerNombre = IIf(IsNull(oRecordset!PrimerNombre), "", oRecordset!PrimerNombre)
       oTabla.SegundoNombre = IIf(IsNull(oRecordset!SegundoNombre), "", oRecordset!SegundoNombre)
       oTabla.NroEnvio = IIf(IsNull(oRecordset!NroEnvio), 0, oRecordset!NroEnvio)
       oTabla.Periodo = IIf(IsNull(oRecordset!Periodo), "", oRecordset!Periodo)
       oTabla.Mes = IIf(IsNull(oRecordset!Mes), "", oRecordset!Mes)
       oTabla.CodigoEstablecimiento = IIf(IsNull(oRecordset!CodigoEstablecimiento), "", oRecordset!CodigoEstablecimiento)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Function SeleccionarPorDNI(oDni As String) As ADODB.Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorDNI = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SisFuaUsuarioSeleccionarPorDNI"
     Set oParameter = .CreateParameter("@DNI", adVarChar, adParamInput, 10, Trim(oDni))
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   
    Set SeleccionarPorDNI = oRecordset
   
   ms_MensajeError = ""
'   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DOSisFuaUsuario)
   oTabla.IdUsuario = 0
   oTabla.DNI = ""
   oTabla.TipoDoc = ""
   oTabla.ApellidoPat = ""
   oTabla.ApellidoMat = ""
   oTabla.PrimerNombre = ""
   oTabla.SegundoNombre = ""
   oTabla.NroEnvio = 0
   oTabla.Periodo = ""
   oTabla.Mes = ""
   oTabla.CodigoEstablecimiento = ""
End Sub
