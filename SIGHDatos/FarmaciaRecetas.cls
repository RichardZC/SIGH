VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FarmaciaRecetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: FarmaciaRecetas
'        Autor: William Castro Grijalva
'        Fecha: 05/08/2005 08:46:43
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
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
'   Descripción:    Inserta un registro a la tabla FarmaciaRecetas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFarmaciaRecetas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FarmaciaRecetasAgregar"
           Set oParameter = .CreateParameter("@IdMedicoOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdena = 0, Null, oTabla.IdMedicoOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrdena = 0, Null, oTabla.IdServicioOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaReceta", adChar, adParamInput, 10, IIf(oTabla.FechaReceta = "", Null, oTabla.FechaReceta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroReceta", adChar, adParamInput, 10, IIf(oTabla.NroReceta = "", Null, oTabla.NroReceta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdReceta = .Parameters("@IdReceta")
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
'   Descripción:    modifica un registro a la tabla FarmaciaRecetas
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFarmaciaRecetas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FarmaciaRecetasModificar"
           Set oParameter = .CreateParameter("@IdMedicoOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdena = 0, Null, oTabla.IdMedicoOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdServicioOrdena = 0, Null, oTabla.IdServicioOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaReceta", adChar, adParamInput, 10, IIf(oTabla.FechaReceta = "", Null, oTabla.FechaReceta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroReceta", adChar, adParamInput, 10, IIf(oTabla.NroReceta = "", Null, oTabla.NroReceta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla FarmaciaRecetas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFarmaciaRecetas) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FarmaciaRecetasEliminar"
           Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FarmaciaRecetas
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFarmaciaRecetas) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FarmaciaRecetasSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdReceta", adInteger, adParamInput, 0, oTabla.IdReceta): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdMedicoOrdena = IIf(IsNull(oRecordset!IdMedicoOrdena), "", oRecordset!IdMedicoOrdena)
       oTabla.IdServicioOrdena = IIf(IsNull(oRecordset!IdServicioOrdena), "", oRecordset!IdServicioOrdena)
       oTabla.FechaReceta = IIf(IsNull(oRecordset!FechaReceta), "", oRecordset!FechaReceta)
       oTabla.NroReceta = IIf(IsNull(oRecordset!NroReceta), "", oRecordset!NroReceta)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), "", oRecordset!IdCuentaAtencion)
       oTabla.IdReceta = IIf(IsNull(oRecordset!IdReceta), "", oRecordset!IdReceta)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function Filtrar(ByVal oTabla As DOFarmaciaRecetas, oDOPaciente As doPaciente) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

     Set Filtrar = Nothing
     
        sWhere = ""
        
        sSQL = "select FarmaciaRecetas.IdCuentaAtencion, FarmaciaRecetas.NroReceta," + _
        "FarmaciaRecetas.FechaReceta, " + _
        "Pacientes.NroHistoriaClinica , TiposNumeracionHistoria.Descripcion as TipoHistoria, FarmaciaRecetas.IdReceta " + _
        "from ((FarmaciaRecetas " + _
        "left join CuentasAtencion on FarmaciaRecetas.IdCuentaAtencion = CuentasAtencion.IdCuentaAtencion) " + _
        "left join Pacientes on CuentasAtencion.IdPaciente = Pacientes.IdPaciente)" + _
        "left join TiposNumeracionHistoria on Pacientes.IdTipoNumeracion = TiposNumeracionHistoria.IdTipoNumeracion"

       If oDOPaciente.NroHistoriaClinica <> 0 Then
            sWhere = sWhere + " Pacientes.NroHistoriaClinica = " & oDOPaciente.NroHistoriaClinica & " and "
       End If
       If oTabla.IdCuentaAtencion <> 0 Then
            sWhere = sWhere + " FarmaciaRecetas.IdCuentaAtencion = " & oTabla.IdCuentaAtencion & " and "
       End If
       If oTabla.NroReceta <> "" Then
            sWhere = sWhere + " FarmaciaRecetas.NroReceta = '" + oTabla.NroReceta + "' and "
       End If

       If sWhere <> "" Then
            sSQL = sSQL + " where " + Left(sWhere, Len(sWhere) - 4)
       End If

        sSQL = sSQL + " order by Pacientes.NroHistoriaClinica, FarmaciaRecetas.IdCuentaAtencion "
        
        With oCommand
           .CommandType = adCmdText
           Set .ActiveConnection = mo_Conexion
           .CommandText = sSQL
           Set oRecordset = .Execute
           Set oRecordset.ActiveConnection = Nothing
        End With
    
    Set Filtrar = oRecordset
    
    ms_MensajeError = ""
   
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

