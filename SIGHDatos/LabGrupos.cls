VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla LabGrupos
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------

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
'   Descripción:    Inserta un registro a la tabla labGrupos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOlabGrupos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "labGruposAgregar"
   Set oParameter = .CreateParameter("@idGrupo", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NombreGrupo", adChar, adParamInput, 20, IIf(oTabla.NombreGrupo = "", Null, oTabla.NombreGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SiglasGrupo", adChar, adParamInput, 3, IIf(oTabla.SiglasGrupo = "", Null, oTabla.SiglasGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCargo", adInteger, adParamInput, 0, IIf(oTabla.idCargo = 0, Null, oTabla.idCargo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdGrupo = .Parameters("@idGrupo")
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
'   Descripción:    modifica un registro a la tabla labGrupos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOlabGrupos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "labGruposModificar"
   Set oParameter = .CreateParameter("@idGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdGrupo = 0, Null, oTabla.IdGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NombreGrupo", adChar, adParamInput, 20, IIf(oTabla.NombreGrupo = "", Null, oTabla.NombreGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SiglasGrupo", adChar, adParamInput, 3, IIf(oTabla.SiglasGrupo = "", Null, oTabla.SiglasGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCargo", adInteger, adParamInput, 0, IIf(oTabla.idCargo = 0, Null, oTabla.idCargo))
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
'   Descripción:    elimia o borra un registro a la tabla labGrupos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOlabGrupos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "labGruposEliminar"
   Set oParameter = .CreateParameter("@idGrupo", adInteger, adParamInput, 0, oTabla.IdGrupo)
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
'   Descripción:    Selecciona un unico registro de la tabla labGrupos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOlabGrupos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "labGruposSeleccionarPorId"
     Set oParameter = .CreateParameter("@idGrupo", adInteger, adParamInput, 0, oTabla.IdGrupo)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdGrupo = IIf(IsNull(oRecordset!IdGrupo), 0, oRecordset!IdGrupo)
       oTabla.NombreGrupo = IIf(IsNull(oRecordset!NombreGrupo), "", oRecordset!NombreGrupo)
       oTabla.SiglasGrupo = IIf(IsNull(oRecordset!SiglasGrupo), "", oRecordset!SiglasGrupo)
       oTabla.idCargo = IIf(IsNull(oRecordset!idCargo), 0, oRecordset!idCargo)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOlabGrupos)
   oTabla.IdGrupo = 0
   oTabla.NombreGrupo = ""
   oTabla.SiglasGrupo = ""
   oTabla.idCargo = 0
End Sub

Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    
    Set SeleccionarTodos = Nothing
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = mo_Conexion
      .CommandText = "LabGruposSeleccionarTodos"
      Set oRecordset = .Execute
      Set oRecordset.ActiveConnection = Nothing
    End With
   
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
   Set oCommand = Nothing
   Set oRecordset = Nothing
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

