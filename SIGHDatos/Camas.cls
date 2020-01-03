VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Camas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla Camas
'        Programado por: Castro W
'        Fecha: Agosto 2006
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Conexion As ADODB.Connection
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
'   Descripción:    Inserta un registro a la tabla Camas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCama) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CamasAgregar"
           Set oParameter = .CreateParameter("@Y", adInteger, adParamInput, 0, IIf(oTabla.Y = 0, Null, oTabla.Y)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@X", adInteger, adParamInput, 0, IIf(oTabla.X = 0, Null, oTabla.X)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioUbicacionActual", adInteger, adParamInput, 0, IIf(oTabla.IdServicioUbicacionActual = 0, Null, oTabla.IdServicioUbicacionActual)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 5, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoCama", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoCama = 0, Null, oTabla.IdEstadoCama)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionOcupacion", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionOcupacion = 0, Null, oTabla.IdCondicionOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTiposCama", adInteger, adParamInput, 0, IIf(oTabla.IdTiposCama = 0, Null, oTabla.IdTiposCama)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioPropietario", adInteger, adParamInput, 0, IIf(oTabla.IdServicioPropietario = 0, Null, oTabla.IdServicioPropietario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCama", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdCama = .Parameters("@IdCama")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla Camas
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCama) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CamasModificar"
           Set oParameter = .CreateParameter("@Y", adInteger, adParamInput, 0, IIf(oTabla.Y = 0, Null, oTabla.Y)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@X", adInteger, adParamInput, 0, IIf(oTabla.X = 0, Null, oTabla.X)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, IIf(oTabla.IdPaciente = 0, Null, oTabla.IdPaciente)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioUbicacionActual", adInteger, adParamInput, 0, IIf(oTabla.IdServicioUbicacionActual = 0, Null, oTabla.IdServicioUbicacionActual)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adChar, adParamInput, 5, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstadoCama", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoCama = 0, Null, oTabla.IdEstadoCama)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionOcupacion", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionOcupacion = 0, Null, oTabla.IdCondicionOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTiposCama", adInteger, adParamInput, 0, IIf(oTabla.IdTiposCama = 0, Null, oTabla.IdTiposCama)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioPropietario", adInteger, adParamInput, 0, IIf(oTabla.IdServicioPropietario = 0, Null, oTabla.IdServicioPropietario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, IIf(oTabla.IdCama = 0, Null, oTabla.IdCama)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Modificar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla Camas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCama) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CamasEliminar"
           Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, IIf(oTabla.IdCama = 0, Null, oTabla.IdCama)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Eliminar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla Camas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCama) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CamasSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, oTabla.IdCama): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
   
       SeleccionarPorId = True
       cargaTabla oTabla, oRecordset
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
    
    Set oRecordset = Nothing
    Set oCommand = Nothing
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Sub cargaTabla(ByRef oTabla As DOCama, oRecordset As Recordset)
       oTabla.Y = IIf(IsNull(oRecordset!Y), 0, oRecordset!Y)
       oTabla.X = IIf(IsNull(oRecordset!X), 0, oRecordset!X)
       oTabla.IdPaciente = IIf(IsNull(oRecordset!IdPaciente), 0, oRecordset!IdPaciente)
       oTabla.IdServicioUbicacionActual = IIf(IsNull(oRecordset!IdServicioUbicacionActual), 0, oRecordset!IdServicioUbicacionActual)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.IdEstadoCama = IIf(IsNull(oRecordset!IdEstadoCama), 0, oRecordset!IdEstadoCama)
       oTabla.IdCondicionOcupacion = IIf(IsNull(oRecordset!IdCondicionOcupacion), 0, oRecordset!IdCondicionOcupacion)
       oTabla.IdTiposCama = IIf(IsNull(oRecordset!IdTiposCama), 0, oRecordset!IdTiposCama)
       oTabla.IdServicioPropietario = IIf(IsNull(oRecordset!IdServicioPropietario), 0, oRecordset!IdServicioPropietario)
       oTabla.IdCama = IIf(IsNull(oRecordset!IdCama), 0, oRecordset!IdCama)

End Sub

Function SeleccionarPorServicioUbicacionActual(ByVal lIdServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorServicioUbicacionActual = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "CamasSeleccionarPorServicioUbicacionActual"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lIdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorServicioUbicacionActual = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarDisponibilidadPorServicioUbicacionActual(ByVal lIdServicio As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarDisponibilidadPorServicioUbicacionActual = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "CamasSeleccionarDisponibilidadPorServicioUbicacionActual"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lIdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarDisponibilidadPorServicioUbicacionActual = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ActualizarIdPaciente(lIdPaciente As Long, lIdCama As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CamasLimpiaIdPaciente"
        Set oParameter = .CreateParameter("@lIdPaciente", adInteger, adParamInput, 0, lIdPaciente): .Parameters.Append oParameter
        .Execute
   End With
   'Atualiza la cama con el Id de Paciente
   If lIdCama <> 0 Then
         Set oCommand = Nothing
         Set oParameter = Nothing
         With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = mo_Conexion
             .CommandTimeout = 150
             .CommandText = "CamasActualizaIdPaciente"
             Set oParameter = .CreateParameter("@lIdPaciente", adInteger, adParamInput, 0, lIdPaciente): .Parameters.Append oParameter
             Set oParameter = .CreateParameter("@lIdCama", adInteger, adParamInput, 0, lIdCama): .Parameters.Append oParameter
             .Execute
        End With
   End If
   ActualizarIdPaciente = True
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function SeleccionarPorCodigo(ByVal oTabla As DOCama) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CamasSeleccionarXcodigo"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 5, oTabla.codigo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   SeleccionarPorCodigo = False
   If oRecordset.RecordCount > 0 Then
      cargaTabla oTabla, oRecordset
      SeleccionarPorCodigo = True
   End If
   Set oCommand = Nothing
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerConElMismoCodigo(ByVal oTabla As DOCama) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CamasObtenerConElMismoCodigo"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 5, oTabla.codigo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, oTabla.IdCama): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerConElMismoCodigo = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
Exit Function
End Function

