VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionesEstanciaHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla AtencionesEstanciaHospitalaria
'        Programado por: Castro W
'        Fecha: Noviembre 2006
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
'   Descripción:    Inserta un registro a la tabla AtencionesEstanciaHospitalaria
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOEstanciaHospitalaria) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   oTabla.idProducto = DevuelveIdProductoSegunIdServicio(oTabla.IdServicio)
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesEstanciaHospitalariaAgregar"
           Set oParameter = .CreateParameter("@DiasEstancia", adDecimal, adParamInput, 5, IIf(oTabla.DiasEstancia = 0, Null, oTabla.DiasEstancia)):
           oParameter.Precision = 8
           oParameter.NumericScale = 2
           .Parameters.Append oParameter
           
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, Null, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdena = 0, Null, oTabla.IdMedicoOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, IIf(oTabla.IdCama = 0, Null, oTabla.IdCama)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraDesocupacion", adChar, adParamInput, 5, IIf(oTabla.HoraDesocupacion = "", Null, oTabla.HoraDesocupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaDesocupacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaDesocupacion = 0, Null, oTabla.FechaDesocupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraOcupacion", adChar, adParamInput, 5, IIf(oTabla.HoraOcupacion = "", Null, oTabla.HoraOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaOcupacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaOcupacion = 0, Null, oTabla.FechaOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Secuencia", adInteger, adParamInput, 0, IIf(oTabla.Secuencia = 0, Null, oTabla.Secuencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstanciaHospitalaria", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LlegoAlServicio", adInteger, adParamInput, 0, oTabla.LlegoAlServicio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           'actualizado 20142009
           Set oParameter = .CreateParameter("@IdMedicoOrdenaOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdenaOrigen = 0, Null, oTabla.IdMedicoOrdenaOrigen)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdEstanciaHospitalaria = .Parameters("@IdEstanciaHospitalaria")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "Insertar"
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla AtencionesEstanciaHospitalaria
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOEstanciaHospitalaria) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   oTabla.idProducto = DevuelveIdProductoSegunIdServicio(oTabla.IdServicio)
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesEstanciaHospitalariaModificar"
           Set oParameter = .CreateParameter("@DiasEstancia", adDecimal, adParamInput, 0, IIf(oTabla.DiasEstancia = 0, Null, oTabla.DiasEstancia)):
           oParameter.Precision = 8
           oParameter.NumericScale = 2
           .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFacturacionServicio", adInteger, adParamInput, 0, IIf(oTabla.IdFacturacionServicio = 0, 0, oTabla.IdFacturacionServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdMedicoOrdena", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdena = 0, Null, oTabla.IdMedicoOrdena)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCama", adInteger, adParamInput, 0, IIf(oTabla.IdCama = 0, Null, oTabla.IdCama)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraDesocupacion", adChar, adParamInput, 5, IIf(oTabla.HoraDesocupacion = "", Null, oTabla.HoraDesocupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaDesocupacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaDesocupacion = 0, Null, oTabla.FechaDesocupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@HoraOcupacion", adChar, adParamInput, 5, IIf(oTabla.HoraOcupacion = "", Null, oTabla.HoraOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaOcupacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaOcupacion = 0, Null, oTabla.FechaOcupacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Secuencia", adInteger, adParamInput, 0, IIf(oTabla.Secuencia = 0, Null, oTabla.Secuencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdEstanciaHospitalaria", adInteger, adParamInput, 0, oTabla.IdEstanciaHospitalaria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LlegoAlServicio", adInteger, adParamInput, 0, oTabla.LlegoAlServicio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           'actualizado 20142009
           Set oParameter = .CreateParameter("@IdMedicoOrdenaOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdMedicoOrdenaOrigen = 0, Null, oTabla.IdMedicoOrdenaOrigen)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla AtencionesEstanciaHospitalaria
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOEstanciaHospitalaria) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "AtencionesEstanciaHospitalariaEliminar"
           Set oParameter = .CreateParameter("@IdEstanciaHospitalaria", adInteger, adParamInput, 0, IIf(oTabla.IdEstanciaHospitalaria = 0, Null, oTabla.IdEstanciaHospitalaria)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla AtencionesEstanciaHospitalaria
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOEstanciaHospitalaria) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "AtencionesEstanciaHospitalariaSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdEstanciaHospitalaria", adInteger, adParamInput, 0, oTabla.IdEstanciaHospitalaria): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.DiasEstancia = IIf(IsNull(oRecordset!DiasEstancia), 0, oRecordset!DiasEstancia)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.IdFacturacionServicio = IIf(IsNull(oRecordset!IdFacturacionServicio), 0, oRecordset!IdFacturacionServicio)
       oTabla.IdMedicoOrdena = IIf(IsNull(oRecordset!IdMedicoOrdena), 0, oRecordset!IdMedicoOrdena)
       oTabla.IdCama = IIf(IsNull(oRecordset!IdCama), 0, oRecordset!IdCama)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
       oTabla.HoraDesocupacion = IIf(IsNull(oRecordset!HoraDesocupacion), "", oRecordset!HoraDesocupacion)
       oTabla.FechaDesocupacion = IIf(IsNull(oRecordset!FechaDesocupacion), 0, oRecordset!FechaDesocupacion)
       oTabla.HoraOcupacion = IIf(IsNull(oRecordset!HoraOcupacion), "", oRecordset!HoraOcupacion)
       oTabla.FechaOcupacion = IIf(IsNull(oRecordset!FechaOcupacion), 0, oRecordset!FechaOcupacion)
       oTabla.Secuencia = IIf(IsNull(oRecordset!Secuencia), 0, oRecordset!Secuencia)
       oTabla.IdEstanciaHospitalaria = IIf(IsNull(oRecordset!IdEstanciaHospitalaria), 0, oRecordset!IdEstanciaHospitalaria)
       oTabla.LlegoAlServicio = IIf(IsNull(oRecordset!LlegoAlServicio), 0, oRecordset!LlegoAlServicio)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       'actualizado 20142009
       oTabla.IdMedicoOrdenaOrigen = IIf(IsNull(oRecordset!IdMedicoOrdenaOrigen), 0, oRecordset!IdMedicoOrdenaOrigen)
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



Function SeleccionarPorAtencion(lIdAtencion As Long, lnSecuenciaMayorA As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorAtencion = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "EstanciaHospitalariaSeleccionarPorAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@SecuenciaMayorA", adInteger, adParamInput, 0, lnSecuenciaMayorA): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorAtencion = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "SeleccionarPorAtencion"
Exit Function
End Function

Function SeleccionarTodosPorCuentaAtencion(lIdCuentaAtencion As Long) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "AtencionesEstanciaHospitalariaSeleccionarTodosPorCuentaAtencion"
        Set oParameter = .CreateParameter("@lIdCuentaAtencion", adInteger, adParamInput, 0, lIdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTodosPorCuentaAtencion = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ActualizarEstanciaHospitalaria(oEstanciaHospitalaria As Collection, oDOAtencion As DOAtencion) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
Dim oDOEstanciaHospitalaria As New DOEstanciaHospitalaria
 
   ActualizarEstanciaHospitalaria = False
   
    'Revisamos si las estancias ya estan completas para generar sus servicios
    Dim bModificarEstancias As Boolean
    Dim bRegistrarServicio As Boolean
    bModificarEstancias = True
    bRegistrarServicio = True
    
    If Not oEstanciaHospitalaria Is Nothing Then
        For Each oDOEstanciaHospitalaria In oEstanciaHospitalaria
            If oDOEstanciaHospitalaria.IdFacturacionServicio <> 0 Then
                bModificarEstancias = False
                Exit For
            End If
            If oDOEstanciaHospitalaria.FechaOcupacion = 0 Or oDOEstanciaHospitalaria.FechaDesocupacion = 0 Then
                bRegistrarServicio = False
            End If
        Next
    Else
        bModificarEstancias = False
    End If
    
    'Si es que aun se puede modificar las estancias
    If bModificarEstancias Then

        With oCommand
             .CommandType = adCmdStoredProc
             Set .ActiveConnection = mo_Conexion
             .CommandTimeout = 150
             .CommandText = "AtencionesEstanciaHospitalariaEliminaXidAtencion"
             Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, oDOAtencion.idAtencion): .Parameters.Append oParameter
             .Execute
        End With
        
        
            
        'Si es que ya se completaron todos las estancias, (el paciente ha sido dado de alta)
        If bRegistrarServicio Then
            Dim oFactOrdenesServicio As New FactOrdenesServicio
            Set oFactOrdenesServicio.Conexion = mo_Conexion
            
             'Crear orden de servicio
             Dim oDoFactordenServicio As New DOFactOrdenServicio
             With oDoFactordenServicio
                 .FechaCreacion = Now
                 .FechaModificacion = 0
                 .FechaOrden = Now
                 .idAtencion = oDOAtencion.idAtencion
                 .IdOrden = 0
                 .IdPuntoCarga = 8
                 .IdUsuarioAuditoria = oDOAtencion.IdUsuarioAuditoria
                 .IdUsuarioCrea = oDOAtencion.IdUsuarioAuditoria
             End With
            

            'Inserta en la tabla de estancias y en FacturacionServicios
            For Each oDOEstanciaHospitalaria In oEstanciaHospitalaria
                If Not AgregarServicioyEstanciaHospitalaria(oDOEstanciaHospitalaria, oDOAtencion, oDoFactordenServicio) Then
                    Exit Function
                End If
             Next
        Else
            'Insert solo en la tabla de estancias
            For Each oDOEstanciaHospitalaria In oEstanciaHospitalaria
                oDOEstanciaHospitalaria.idAtencion = oDOAtencion.idAtencion
                If Not Insertar(oDOEstanciaHospitalaria) Then
                    Exit Function
                End If
             Next
        End If
    End If
    ActualizarEstanciaHospitalaria = True
    ms_MensajeError = ""
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "ActualizarEstanciaHospitalaria"
Exit Function
End Function

Function AgregarServicioyEstanciaHospitalaria(oDOEstanciaHospitalaria As DOEstanciaHospitalaria, oDOAtencion As DOAtencion, oDoFactordenServicio As DOFactOrdenServicio) As Boolean
Dim oDOFacturacionServicio As New DOFacturacionServicios
Dim oDoCatalogoServicio As New DOCatalogoServicio
Dim oDOServicios As New DOServicio

Dim oServicios As New Servicios
Dim oCatalogoServicios As New CatalogoServicios
Dim oFacturacionServicios As New FacturacionServicios

    AgregarServicioyEstanciaHospitalaria = False
    
    Set oServicios.Conexion = mo_Conexion
    Set oFacturacionServicios.Conexion = mo_Conexion
    Set oCatalogoServicios.Conexion = mo_Conexion
    
    Dim dFHIngreso As Date
    Dim dFHEgreso As Date
    dFHIngreso = 0
    dFHEgreso = 0
    If oDOEstanciaHospitalaria.FechaOcupacion <> 0 And oDOEstanciaHospitalaria.HoraOcupacion <> "" Then
        dFHIngreso = CDate(Format(oDOEstanciaHospitalaria.FechaOcupacion, SIGHENTIDADES.DevuelveFechaSoloFormato_DMY) + " " + Format(oDOEstanciaHospitalaria.HoraOcupacion, "hh:nn"))
    End If
    If oDOEstanciaHospitalaria.FechaDesocupacion <> 0 And oDOEstanciaHospitalaria.HoraDesocupacion <> "" Then
        dFHEgreso = CDate(Format(oDOEstanciaHospitalaria.FechaDesocupacion, SIGHENTIDADES.DevuelveFechaSoloFormato_DMY) + " " + Format(oDOEstanciaHospitalaria.HoraDesocupacion, "hh:nn"))
    End If
    
    If dFHIngreso <> 0 And dFHEgreso <> 0 Then
        oDOServicios.IdServicio = oDOEstanciaHospitalaria.IdServicio
        If Not oServicios.SeleccionarPorId(oDOServicios) Then Exit Function
            If oDOServicios.idProducto <> 0 Then
                '<<<<<<<<<<<
                With oDOFacturacionServicio
                    .IdOrden = oDoFactordenServicio.IdOrden
                    .idAtencion = oDOAtencion.idAtencion
                    .IdFacturacionServicio = 0
                    .idFuenteFinanciamiento = sghFuenteFinanciamiento.sghFFPaciente
                    .idTipoFinanciamiento = sghPacienteNormal
                    .Cantidad = oDOEstanciaHospitalaria.DiasEstancia * 24
                    .idProducto = oDOServicios.idProducto
                    .IdUsuarioAuditoria = oDOEstanciaHospitalaria.IdUsuarioAuditoria
                    .IdEstadoFacturacion = sghEstadoFacturacion.sghRegistrado
                    .FechaAutorizaPendiente = 0
                    .FechaAutorizaSeguro = 0
                    .IdCentroCosto = 0
                    .IdUsuarioAutorizaPendiente = 0
                    .IdUsuarioAutorizaSeguro = 0
                    .PrecioUnitario = 0
                    .TotalPorPagar = 0
                    .IdServicioInternamiento = oDOEstanciaHospitalaria.IdServicio
                    .IdPuntoCarga = 8
                    
                End With
            

                Dim sResultado As String
                If oDOServicios.idProducto <> 4238 Then

                End If
                oDOEstanciaHospitalaria.idAtencion = oDOAtencion.idAtencion
                If Insertar(oDOEstanciaHospitalaria) Then
                Else
                    Trace sResultado, "AgregarServicioyEstanciaHospitalaria"
                    Exit Function
                End If
            End If
        End If
    AgregarServicioyEstanciaHospitalaria = True

End Function

Function EliminarEstanciaHospitalaria(lIdAtencion As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
 
   EliminarEstanciaHospitalaria = False
   
   'elimina los diagnosticos
   With oCommand
         .CommandType = adCmdStoredProc
         Set .ActiveConnection = mo_Conexion
         .CommandTimeout = 150
         .CommandText = "AtencionesEstanciaHospitalariaEliminaXidAtencion"
         Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lIdAtencion): .Parameters.Append oParameter
         .Execute
   End With
   
   
   EliminarEstanciaHospitalaria = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   Trace Err.Number & " " + Err.Description, "EliminarEstanciaHospitalaria"
Exit Function
End Function


Sub Trace(sMensaje As String, NombreDeMetodo As String)
    MsgBox "Mensaje: " & sMensaje + Chr(13) + "Metodo:" & NombreDeMetodo + Chr(13) + "Clase de Datos: AtencionesEstanciaHosp", vbInformation, "Atenciones Estancia Hospitalaria"
End Sub



Function DevuelveIdProductoSegunIdServicio(lnIdServicio As Long) As Long

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHENTIDADES.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "ServiciosSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   DevuelveIdProductoSegunIdServicio = 0
   If oRecordset.RecordCount > 0 Then
       DevuelveIdProductoSegunIdServicio = IIf(IsNull(oRecordset.Fields!idProducto), 0, oRecordset.Fields!idProducto)
   End If
   oConexion.Close
   Set oCommand = Nothing
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
