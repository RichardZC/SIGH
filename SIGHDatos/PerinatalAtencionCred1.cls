VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PerinatalAtencionCred1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla PerinatalAtencionCred1
'        Programado por: Barrantes D
'        Fecha: Agosto 2011
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
'   Descripción:    Inserta un registro a la tabla PerinatalAtencionCred1
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOPerinatalAtencionCred1) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionCred1Agregar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, IIf(oTabla.idPerinatalAtencion = 0, Null, oTabla.idPerinatalAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idModulo", adInteger, adParamInput, 0, IIf(oTabla.idModulo = 0, Null, oTabla.idModulo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstimulacionTemprana", adBoolean, adParamInput, 0, IIf(oTabla.EstimulacionTemprana = 0, Null, oTabla.EstimulacionTemprana))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AlimentacionComplementaria", adBoolean, adParamInput, 0, IIf(oTabla.AlimentacionComplementaria = 0, Null, oTabla.AlimentacionComplementaria))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LactanciaMaterna", adBoolean, adParamInput, 0, IIf(oTabla.LactanciaMaterna = 0, Null, oTabla.LactanciaMaterna))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PersonalSalud", adBoolean, adParamInput, 0, IIf(oTabla.PersonalSalud = 0, Null, oTabla.PersonalSalud))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DemandaIndividual", adBoolean, adParamInput, 0, IIf(oTabla.DemandaIndividual = 0, Null, oTabla.DemandaIndividual))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MujerEdadReproductiva", adBoolean, adParamInput, 0, IIf(oTabla.MujerEdadReproductiva = 0, Null, oTabla.MujerEdadReproductiva))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MujerGestante", adBoolean, adParamInput, 0, IIf(oTabla.MujerGestante = 0, Null, oTabla.MujerGestante))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LactanciaMaternaComp", adBoolean, adParamInput, 0, IIf(oTabla.LactanciaMaternaComp = 0, Null, oTabla.LactanciaMaternaComp))
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
'   Descripción:    modifica un registro a la tabla PerinatalAtencionCred1
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOPerinatalAtencionCred1) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionCred1Modificar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, IIf(oTabla.idPerinatalAtencion = 0, Null, oTabla.idPerinatalAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idModulo", adInteger, adParamInput, 0, IIf(oTabla.idModulo = 0, Null, oTabla.idModulo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstimulacionTemprana", adBoolean, adParamInput, 0, IIf(oTabla.EstimulacionTemprana = 0, Null, oTabla.EstimulacionTemprana))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@AlimentacionComplementaria", adBoolean, adParamInput, 0, IIf(oTabla.AlimentacionComplementaria = 0, Null, oTabla.AlimentacionComplementaria))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LactanciaMaterna", adBoolean, adParamInput, 0, IIf(oTabla.LactanciaMaterna = 0, Null, oTabla.LactanciaMaterna))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PersonalSalud", adBoolean, adParamInput, 0, IIf(oTabla.PersonalSalud = 0, Null, oTabla.PersonalSalud))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DemandaIndividual", adBoolean, adParamInput, 0, IIf(oTabla.DemandaIndividual = 0, Null, oTabla.DemandaIndividual))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MujerEdadReproductiva", adBoolean, adParamInput, 0, IIf(oTabla.MujerEdadReproductiva = 0, Null, oTabla.MujerEdadReproductiva))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MujerGestante", adBoolean, adParamInput, 0, IIf(oTabla.MujerGestante = 0, Null, oTabla.MujerGestante))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAtencion", adInteger, adParamInput, 0, IIf(oTabla.idAtencion = 0, Null, oTabla.idAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LactanciaMaternaComp", adBoolean, adParamInput, 0, IIf(oTabla.LactanciaMaternaComp = 0, Null, oTabla.LactanciaMaternaComp))
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
'   Descripción:    elimia o borra un registro a la tabla PerinatalAtencionCred1
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOPerinatalAtencionCred1) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "PerinatalAtencionCred1Eliminar"
   Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, oTabla.idPerinatalAtencion)
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
'   Descripción:    Selecciona un unico registro de la tabla PerinatalAtencionCred1
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOPerinatalAtencionCred1) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PerinatalAtencionCred1SeleccionarPorId"
     Set oParameter = .CreateParameter("@idPerinatalAtencion", adInteger, adParamInput, 0, oTabla.idPerinatalAtencion)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idPerinatalAtencion = IIf(IsNull(oRecordset!idPerinatalAtencion), 0, oRecordset!idPerinatalAtencion)
       oTabla.idModulo = IIf(IsNull(oRecordset!idModulo), 0, oRecordset!idModulo)
       oTabla.EstimulacionTemprana = IIf(IsNull(oRecordset!EstimulacionTemprana), 0, oRecordset!EstimulacionTemprana)
       oTabla.AlimentacionComplementaria = IIf(IsNull(oRecordset!AlimentacionComplementaria), 0, oRecordset!AlimentacionComplementaria)
       oTabla.LactanciaMaterna = IIf(IsNull(oRecordset!LactanciaMaterna), 0, oRecordset!LactanciaMaterna)
       oTabla.PersonalSalud = IIf(IsNull(oRecordset!PersonalSalud), 0, oRecordset!PersonalSalud)
       oTabla.DemandaIndividual = IIf(IsNull(oRecordset!DemandaIndividual), 0, oRecordset!DemandaIndividual)
       oTabla.MujerEdadReproductiva = IIf(IsNull(oRecordset!MujerEdadReproductiva), 0, oRecordset!MujerEdadReproductiva)
       oTabla.MujerGestante = IIf(IsNull(oRecordset!MujerGestante), 0, oRecordset!MujerGestante)
       oTabla.idAtencion = IIf(IsNull(oRecordset!idAtencion), 0, oRecordset!idAtencion)
       oTabla.LactanciaMaternaComp = IIf(IsNull(oRecordset!LactanciaMaternaComp), 0, oRecordset!LactanciaMaternaComp)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOPerinatalAtencionCred1)
   oTabla.idPerinatalAtencion = 0
   oTabla.idModulo = 0
   oTabla.EstimulacionTemprana = False
   oTabla.AlimentacionComplementaria = False
   oTabla.LactanciaMaterna = False
   oTabla.PersonalSalud = False
   oTabla.DemandaIndividual = False
   oTabla.MujerEdadReproductiva = False
   oTabla.MujerGestante = False
   oTabla.idAtencion = 0
End Sub
 Function PerinatalAtencionCred1SeleccionarTodoPorIdAtencion(lnIdAtencion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sWhere As String
Dim oParameter As ADODB.Parameter
     
     With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandType = adCmdStoredProc
       .CommandText = "PerinatalAtencionCred1SeleccionarTodoPorIdAtencion"
        Set oParameter = .CreateParameter("@IdAtencion", adInteger, adParamInput, 0, lnIdAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
     End With
     Set PerinatalAtencionCred1SeleccionarTodoPorIdAtencion = oRecordset
     ms_MensajeError = ""
     Set oCommand = Nothing
Exit Function
ManejadorDeError:
   MsgBox Err.Description
Exit Function
End Function


