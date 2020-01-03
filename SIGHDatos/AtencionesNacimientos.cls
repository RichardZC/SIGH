VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesNacimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtencionesNacimientos
'        Programado por: Barrantes D
'        Fecha: Octubre 2013
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
'   Descripción:    Inserta un registro a la tabla AtencionesNacimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOAtencionNacimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesNacimientosAgregar"
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionRN", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionRN = 0, Null, oTabla.IdCondicionRN)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Peso", adDouble, adParamInput, 0, IIf(oTabla.Peso = 0, Null, oTabla.Peso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Talla", adDouble, adParamInput, 0, IIf(oTabla.Talla = 0, Null, oTabla.Talla)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EdadSemanas", adInteger, adParamInput, 0, oTabla.EdadSemanas): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdNacimiento", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@apgar_1", adInteger, adParamInput, 0, IIf(oTabla.Apgar_1 = 0, Null, oTabla.Apgar_1)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@apgar_5", adInteger, adParamInput, 0, IIf(oTabla.Apgar_5 = 0, Null, oTabla.Apgar_5)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ClamplajeFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ClamplajeFecha = 0, Null, oTabla.ClamplajeFecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroOrdenHijoEnParto", adInteger, adParamInput, 0, IIf(oTabla.NroOrdenHijoEnParto = 0, Null, oTabla.NroOrdenHijoEnParto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroOrdenHijo", adInteger, adParamInput, 0, IIf(oTabla.NroOrdenHijo = 0, Null, oTabla.NroOrdenHijo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idPacienteNacido", adInteger, adParamInput, 0, IIf(oTabla.idPacienteNacido = 0, Null, oTabla.idPacienteNacido)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@docIdentidad", adVarChar, adParamInput, 20, IIf(oTabla.DocIdentidad = "", Null, oTabla.DocIdentidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, IIf(oTabla.IdDocIdentidad = 0, Null, oTabla.IdDocIdentidad)): .Parameters.Append oParameter
       
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdNacimiento = .Parameters("@IdNacimiento")
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
'   Descripción:    modifica un registro a la tabla AtencionesNacimientos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOAtencionNacimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesNacimientosModificar"
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCondicionRN", adInteger, adParamInput, 0, IIf(oTabla.IdCondicionRN = 0, Null, oTabla.IdCondicionRN)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoSexo", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSexo = 0, Null, oTabla.IdTipoSexo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Peso", adDouble, adParamInput, 0, IIf(oTabla.Peso = 0, Null, oTabla.Peso)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Talla", adDouble, adParamInput, 0, IIf(oTabla.Talla = 0, Null, oTabla.Talla)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EdadSemanas", adInteger, adParamInput, 0, oTabla.EdadSemanas): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaNacimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaNacimiento = 0, Null, oTabla.FechaNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdNacimiento = 0, Null, oTabla.IdNacimiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@apgar_1", adInteger, adParamInput, 0, IIf(oTabla.Apgar_1 = 0, Null, oTabla.Apgar_1)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@apgar_5", adInteger, adParamInput, 0, IIf(oTabla.Apgar_5 = 0, Null, oTabla.Apgar_5)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ClamplajeFecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.ClamplajeFecha = 0, Null, oTabla.ClamplajeFecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroOrdenHijoEnParto", adInteger, adParamInput, 0, IIf(oTabla.NroOrdenHijoEnParto = 0, Null, oTabla.NroOrdenHijoEnParto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroOrdenHijo", adInteger, adParamInput, 0, IIf(oTabla.NroOrdenHijo = 0, Null, oTabla.NroOrdenHijo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idPacienteNacido", adInteger, adParamInput, 0, IIf(oTabla.idPacienteNacido = 0, Null, oTabla.idPacienteNacido)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@docIdentidad", adVarChar, adParamInput, 20, IIf(oTabla.DocIdentidad = "", Null, oTabla.DocIdentidad)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDocIdentidad", adInteger, adParamInput, 0, IIf(oTabla.IdDocIdentidad = 0, Null, oTabla.IdDocIdentidad)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtencionesNacimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOAtencionNacimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesNacimientosEliminar"
           Set oParameter = .CreateParameter("@IdNacimiento", adInteger, adParamInput, 0, IIf(oTabla.IdNacimiento = 0, Null, oTabla.IdNacimiento)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesNacimientos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOAtencionNacimiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesNacimientosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdNacimiento", adInteger, adParamInput, 0, oTabla.IdNacimiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.IdCondicionRN = IIf(IsNull(oRecordset!IdCondicionRN), 0, oRecordset!IdCondicionRN)
       oTabla.IdTipoSexo = IIf(IsNull(oRecordset!IdTipoSexo), 0, oRecordset!IdTipoSexo)
       oTabla.Peso = IIf(IsNull(oRecordset!Peso), 0, oRecordset!Peso)
       oTabla.Talla = IIf(IsNull(oRecordset!Talla), 0, oRecordset!Talla)
       oTabla.EdadSemanas = IIf(IsNull(oRecordset!EdadSemanas), 0, oRecordset!EdadSemanas)
       oTabla.FechaNacimiento = IIf(IsNull(oRecordset!FechaNacimiento), 0, oRecordset!FechaNacimiento)
       oTabla.IdNacimiento = IIf(IsNull(oRecordset!IdNacimiento), 0, oRecordset!IdNacimiento)
       oTabla.Apgar_1 = IIf(IsNull(oRecordset!Apgar_1), 0, oRecordset!Apgar_1)
       oTabla.Apgar_5 = IIf(IsNull(oRecordset!Apgar_5), 0, oRecordset!Apgar_5)
       oTabla.ClamplajeFecha = IIf(IsNull(oRecordset!ClamplajeFecha), 0, oRecordset!ClamplajeFecha)
       oTabla.NroOrdenHijoEnParto = IIf(IsNull(oRecordset!NroOrdenHijoEnParto), 0, oRecordset!NroOrdenHijoEnParto)
       oTabla.NroOrdenHijo = IIf(IsNull(oRecordset!NroOrdenHijo), 1, oRecordset!NroOrdenHijo)
       oTabla.idPacienteNacido = IIf(IsNull(oRecordset!idPacienteNacido), 1, oRecordset!idPacienteNacido)
       oTabla.IdDocIdentidad = IIf(IsNull(oRecordset!IdDocIdentidad), 0, oRecordset!IdDocIdentidad)
       oTabla.DocIdentidad = IIf(IsNull(oRecordset!DocIdentidad), "", oRecordset!DocIdentidad)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorAtencion(lIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorAtencion = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesNacimientosSeleccionarPorAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorAtencion = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function ActualizarNacimientosAtencion(oNacimientos As Collection, lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oDOAtencionNacimiento As New DOAtencionNacimiento
 
   ActualizarNacimientosAtencion = False
   
   'elimina todos los dx menos los de interconsultas

   With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = mo_Conexion
         .CommandTimeout = 150
         .CommandText = "AtencionesNacimientosEliminarXIdAtencion"
         Set oParameter = .CreateParameter("@lIdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
         .Execute
   End With
   
   
   If Not oNacimientos Is Nothing Then
        For Each oDOAtencionNacimiento In oNacimientos
            oDOAtencionNacimiento.idAtencion = lIdAtencion
            If Not Insertar(oDOAtencionNacimiento) Then
                Exit Function
            End If
         Next
    End If
    
   ActualizarNacimientosAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
      MsgBox Err.Description
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarNacimientosPorAtencion(lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
 
   EliminarNacimientosPorAtencion = False
   
   'elimina los diagnosticos

   With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = mo_Conexion
         .CommandTimeout = 150
         .CommandText = "AtencionesNacimientosEliminarXIdAtencion"
         Set oParameter = .CreateParameter("@lIdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
         .Execute
   End With
   
   EliminarNacimientosPorAtencion = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function AsignaNULLidPacienteNacido(lnIdPaciente As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "AtencionesNacimientosAsignaNULLidPacienteNacido"
        Set oParameter = .CreateParameter("@idPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        .Execute
   End With
   AsignaNULLidPacienteNacido = True
   Set oRecordset = Nothing
   Set oCommand = Nothing
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

