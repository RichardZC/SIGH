VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasAtencionIntegral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para atención integral
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String

Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property

Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Public Function SeleccionarPlanIntegralPorPacienteYGrupo(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As DOAtenIntePlanIntePaciente
On Error GoTo miError
'    Dim oPregunta As New AtenIntePregunta
    Dim oAtenIntePlanIntegralPaciente As New AtenIntePlanIntegralPaciente
    
    
    Dim oConexion As New ADODB.Connection
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set SeleccionarPlanIntegralPorPacienteYGrupo = Nothing
    
    Set oAtenIntePlanIntegralPaciente.Conexion = oConexion
    
    If oAtenIntePlanIntegralPaciente.SeleccionarPorPacienteYGrupo(oDOAtenIntePlanIntePaciente) Then
        Set SeleccionarPlanIntegralPorPacienteYGrupo = oDOAtenIntePlanIntePaciente
    End If
    ms_MensajeError = oAtenIntePlanIntegralPaciente.MensajeError
    Set oAtenIntePlanIntegralPaciente = Nothing
    Set oDOAtenIntePlanIntePaciente = Nothing
    oConexion.Close
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
End Function

Private Function ListarPlanProcedimientoPaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenIntePlanProcedPaciente
    Dim rsPlanProcedimiento As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set rsPlanProcedimiento = oTabla.ListarPlanProcedimientosPaciente(oDOAtenIntePlanIntePaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set ListarPlanProcedimientoPaciente = rsPlanProcedimiento
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Private Function ListarPlanProcedimientosPacientePendientes(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenIntePlanProcedPaciente
    Dim rsPlanProcedimiento As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    'oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set rsPlanProcedimiento = oTabla.ListarPlanProcedimientosPacientePendientes(oDOAtenIntePlanIntePaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    
    Set ListarPlanProcedimientosPacientePendientes = rsPlanProcedimiento
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Public Function ListarPlanInmunizacionPaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
    
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Inmunizacion
    Set ListarPlanInmunizacionPaciente = ListarPlanProcedimientoPaciente(oDOAtenIntePlanIntePaciente)
    
End Function

Public Function ListarPlanInmunizacionPacientePendientes(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Inmunizacion
    Set ListarPlanInmunizacionPacientePendientes = ListarPlanProcedimientosPacientePendientes(oDOAtenIntePlanIntePaciente)
End Function

Public Function ListarPlanTamizajePaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Tamizaje
    Set ListarPlanTamizajePaciente = ListarPlanProcedimientoPaciente(oDOAtenIntePlanIntePaciente)
End Function

Public Function ListarPlanTamizajePacientePendientes(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Tamizaje
    Set ListarPlanTamizajePacientePendientes = ListarPlanProcedimientosPacientePendientes(oDOAtenIntePlanIntePaciente)
End Function

Private Function GenerarPlanProcedimientosPaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As Boolean
On Error GoTo miError

    Dim ErrorInTransaction As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oAtenIntePlanIntegralPaciente As New AtenIntePlanIntegralPaciente
    ErrorInTransaction = False
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    
    Set oAtenIntePlanIntegralPaciente.Conexion = oConexion
    
    If oAtenIntePlanIntegralPaciente.verificarInsertar(oDOAtenIntePlanIntePaciente) = True Then
        If oAtenIntePlanIntegralPaciente.GenerarPlanIntegralProcedimiento(oDOAtenIntePlanIntePaciente) = True Then
            Call registrarItemPlanElaborado(oDOAtenIntePlanIntePaciente, _
                                oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan, oConexion)
        Else
            ms_MensajeError = oAtenIntePlanIntegralPaciente.MensajeError
            ErrorInTransaction = True
            GoTo EndTransaction
        End If
    Else
        ms_MensajeError = oAtenIntePlanIntegralPaciente.MensajeError
        ErrorInTransaction = True
        GoTo EndTransaction
    End If
EndTransaction:
    Dim response As Boolean
    response = False
    If ErrorInTransaction = True Then
        oConexion.RollbackTrans
    Else
        oConexion.CommitTrans
        response = True
    End If
    oConexion.Close
    GenerarPlanProcedimientosPaciente = response
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
        If oConexion.State = 1 Then
            oConexion.Close
        End If
    End If
    Set oConexion = Nothing
End Function

Public Function GenerarPlanInmunizacionPaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As Boolean
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Inmunizacion
    GenerarPlanInmunizacionPaciente = GenerarPlanProcedimientosPaciente(oDOAtenIntePlanIntePaciente)
End Function

Public Function GenerarPlanTamizajePaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As Boolean
    GenerarPlanTamizajePaciente = True
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Tamizaje
    GenerarPlanTamizajePaciente = GenerarPlanProcedimientosPaciente(oDOAtenIntePlanIntePaciente)
End Function

Private Function registrarItemPlanElaborado(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente, _
            IdItemPlanIntegral As sighItemPlanIntegral, oConexion As ADODB.Connection) As Boolean
    registrarItemPlanElaborado = True
End Function

Public Function GenerarPlanCrecimientoPaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As Boolean
On Error GoTo miError

    Dim ErrorInTransaction As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oAtenIntePlanIntegralPaciente As New AtenIntePlanIntegralPaciente
    ErrorInTransaction = False
    
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Crecimiento
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    
    Set oAtenIntePlanIntegralPaciente.Conexion = oConexion
    
    If oAtenIntePlanIntegralPaciente.verificarInsertar(oDOAtenIntePlanIntePaciente) = True Then
        If oAtenIntePlanIntegralPaciente.GenerarPlanIntegralCrecimiento(oDOAtenIntePlanIntePaciente) = True Then
            Call registrarItemPlanElaborado(oDOAtenIntePlanIntePaciente, oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan, oConexion)
        Else
            ms_MensajeError = oAtenIntePlanIntegralPaciente.MensajeError
            ErrorInTransaction = True
            GoTo EndTransaction
        End If
    Else
        ms_MensajeError = oAtenIntePlanIntegralPaciente.MensajeError
        ErrorInTransaction = True
        GoTo EndTransaction
    End If
EndTransaction:
    Dim response As Boolean
    response = False
    If ErrorInTransaction = True Then
        oConexion.RollbackTrans
    Else
        oConexion.CommitTrans
        response = True
    End If
    oConexion.Close
    GenerarPlanCrecimientoPaciente = response
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
        If oConexion.State = 1 Then
            oConexion.Close
        End If
    End If
    Set oConexion = Nothing
End Function


Public Function GenerarPlanDesarrolloPaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As Boolean
On Error GoTo miError

    Dim ErrorInTransaction As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oAtenIntePlanIntegralPaciente As New AtenIntePlanIntegralPaciente
    ErrorInTransaction = False
    
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Desarrollo
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    
    Set oAtenIntePlanIntegralPaciente.Conexion = oConexion
    
    If oAtenIntePlanIntegralPaciente.verificarInsertar(oDOAtenIntePlanIntePaciente) = True Then
        If oAtenIntePlanIntegralPaciente.GenerarPlanIntegralDesarrollo(oDOAtenIntePlanIntePaciente) = True Then
            Call registrarItemPlanElaborado(oDOAtenIntePlanIntePaciente, Desarrollo, oConexion)
        Else
            ms_MensajeError = oAtenIntePlanIntegralPaciente.MensajeError
            ErrorInTransaction = True
            GoTo EndTransaction
        End If
    Else
        ms_MensajeError = oAtenIntePlanIntegralPaciente.MensajeError
        ErrorInTransaction = True
        GoTo EndTransaction
    End If
EndTransaction:
    Dim response As Boolean
    response = False
    If ErrorInTransaction = True Then
        oConexion.RollbackTrans
    Else
        oConexion.CommitTrans
        response = True
    End If
    oConexion.Close
    GenerarPlanDesarrolloPaciente = response
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
        If oConexion.State = 1 Then
            oConexion.Close
        End If
    End If
    Set oConexion = Nothing
End Function

Public Function ListarPlanDesarrolloPaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenIntePlanDesPaciente
    Dim rsPlanProcedimiento As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Desarrollo
    
    Set oTabla.Conexion = oConexion
    Set rsPlanProcedimiento = oTabla.ListarPlanDesarrolloPaciente(oDOAtenIntePlanIntePaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set ListarPlanDesarrolloPaciente = rsPlanProcedimiento
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Public Function ListarPlanDesarrolloPacientePendientes(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenIntePlanDesPaciente
    Dim rsPlanProcedimiento As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    'oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Desarrollo
    
    Set oTabla.Conexion = oConexion
    Set rsPlanProcedimiento = oTabla.ListarPlanDesarrolloPacientePendientes(oDOAtenIntePlanIntePaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set ListarPlanDesarrolloPacientePendientes = rsPlanProcedimiento
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Public Function GetItemDesarrolloPacientePendiente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenIntePlanDesPaciente
    Dim rsPlanProcedimiento As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    'oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set rsPlanProcedimiento = oTabla.ItemDesarrolloPacientePendiente(oDOAtenIntePlanIntePaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set GetItemDesarrolloPacientePendiente = rsPlanProcedimiento
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Public Function GenerarPlanSuplementoPaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As Boolean
On Error GoTo miError

    Dim ErrorInTransaction As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oAtenIntePlanIntegralPaciente As New AtenIntePlanIntegralPaciente
    ErrorInTransaction = False
    
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.SuplementoNutricional
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    oConexion.BeginTrans
    
    Set oAtenIntePlanIntegralPaciente.Conexion = oConexion
    
    If oAtenIntePlanIntegralPaciente.verificarInsertar(oDOAtenIntePlanIntePaciente) = True Then
        If oAtenIntePlanIntegralPaciente.GenerarPlanIntegralSuplemento(oDOAtenIntePlanIntePaciente) = True Then
            Call registrarItemPlanElaborado(oDOAtenIntePlanIntePaciente, _
                                oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan, oConexion)
        Else
            ms_MensajeError = oAtenIntePlanIntegralPaciente.MensajeError
            ErrorInTransaction = True
            GoTo EndTransaction
        End If
    Else
        ms_MensajeError = oAtenIntePlanIntegralPaciente.MensajeError
        ErrorInTransaction = True
        GoTo EndTransaction
    End If
EndTransaction:
    Dim response As Boolean
    response = False
    If ErrorInTransaction = True Then
        oConexion.RollbackTrans
    Else
        oConexion.CommitTrans
        response = True
    End If
    oConexion.Close
    GenerarPlanSuplementoPaciente = response
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
        If oConexion.State = 1 Then
            oConexion.Close
        End If
    End If
    Set oConexion = Nothing
End Function

Public Function ListarPlanSuplementoPaciente(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenIntePlanSuplPaciente
    Dim rsPlanProcedimiento As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.SuplementoNutricional
    
    Set oTabla.Conexion = oConexion
    Set rsPlanProcedimiento = oTabla.ListarPlanSuplementoPaciente(oDOAtenIntePlanIntePaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set ListarPlanSuplementoPaciente = rsPlanProcedimiento
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Public Function ListarPlanSuplementoPacientePendientes(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenIntePlanSuplPaciente
    Dim rsPlanProcedimiento As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    'oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.SuplementoNutricional
    
    Set oTabla.Conexion = oConexion
    Set rsPlanProcedimiento = oTabla.ListarPlanSuplementoPacientePendientes(oDOAtenIntePlanIntePaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set ListarPlanSuplementoPacientePendientes = rsPlanProcedimiento
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Public Function ListarPlanDesarrolloPacientePendientesParaImpresion(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenIntePlanDesPaciente
    Dim rsPlanProcedimiento As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    'oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Desarrollo
    
    Set oTabla.Conexion = oConexion
    Set rsPlanProcedimiento = oTabla.ListarPlanDesarrolloPacientePendientesParaImpresion(oDOAtenIntePlanIntePaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set ListarPlanDesarrolloPacientePendientesParaImpresion = rsPlanProcedimiento
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'mgaray20141012
Public Function ListarPlanDesarrolloPacienteVencidos(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenIntePlanDesPaciente
    Dim rsPlanProcedimiento As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    oDOAtenIntePlanIntePaciente.IdAtenInteItemPlan = sighItemPlanIntegral.Desarrollo
    
    Set oTabla.Conexion = oConexion
    Set rsPlanProcedimiento = oTabla.ListarPlanDesarrolloPacienteVencidos(oDOAtenIntePlanIntePaciente)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set ListarPlanDesarrolloPacienteVencidos = rsPlanProcedimiento
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function PlanDesarrolloPacienteSeleccionarPorId(lIdPlanIntegralPaciente As Long, _
            lIdPlanDesarrolloPaciente As Long) As DOAtenIntePlanDesPaciente
            
    Dim oConexion As New ADODB.Connection
    Dim oAtenIntePlanDesPaciente As New AtenIntePlanDesPaciente
    Dim oDOAtenIntePlanDesPaciente As New DOAtenIntePlanDesPaciente
    
    Set PlanDesarrolloPacienteSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oAtenIntePlanDesPaciente.Conexion = oConexion
    oDOAtenIntePlanDesPaciente.IdPlanDesarrolloPaciente = lIdPlanDesarrolloPaciente
    oDOAtenIntePlanDesPaciente.IdPlanIntegralPaciente = lIdPlanIntegralPaciente
    
    If oAtenIntePlanDesPaciente.SeleccionarPorId(oDOAtenIntePlanDesPaciente) Then
        Set PlanDesarrolloPacienteSeleccionarPorId = oDOAtenIntePlanDesPaciente
    End If
    oConexion.Close
    Set oConexion = Nothing
    ms_MensajeError = oAtenIntePlanDesPaciente.MensajeError
    Set oDOAtenIntePlanDesPaciente = Nothing
End Function

'Public Function SeleccionarPlanDesarrolloPacientePorId(oDOAtenIntePlanIntePaciente As DOAtenIntePlanIntePaciente) As ADODB.Recordset
'On Error GoTo miError
'    Dim oConexion As New ADODB.Connection
'    Dim oTabla As New SIGHDatos.AtenIntePlanDesPaciente
'    Dim rsPlanProcedimiento As New ADODB.Recordset
'
'    oConexion.Open sighentidades.CadenaConexion
'    'oConexion.Open sighentidades.CadenaConexionShape
'    oConexion.CursorLocation = adUseClient
'
'    Set oTabla.Conexion = oConexion
'    Set rsPlanProcedimiento = oTabla.ItemDesarrolloPacientePendiente(oDOAtenIntePlanIntePaciente)
'    oConexion.Close
'    ms_MensajeError = oTabla.MensajeError
'    Set GetItemDesarrolloPacientePendiente = rsPlanProcedimiento
'miError:
'    If Err Then
'        ms_MensajeError = Err.Number & " : " & Err.Description
'    End If
'    Set oConexion = Nothing
'    Set oTabla = Nothing
'End Function

Public Function ListarPlanDesarrolloPacienteDetallePorId(oDOAtenIntePlanDesPacienteDet As DOAtenIntePlanDesPacienteDet) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenIntePlanDesPacienteDet
    Dim rsPlanProcedimiento As New ADODB.Recordset
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set rsPlanProcedimiento = oTabla.ListarPlanDesarrolloPacienteDetallePorId(oDOAtenIntePlanDesPacienteDet)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set ListarPlanDesarrolloPacienteDetallePorId = rsPlanProcedimiento
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

'mgaray201411e
Public Function DevuelveValorPesoPorSexoDesviacion(lIdTipoSexo As Long, lNroDesviacion As Long) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenInteValorPeso
    Dim oDOValorPeso As New DOAtenInteValorPeso
    Dim oRsValorPeso As New ADODB.Recordset
    
    oDOValorPeso.IdTipoSexo = lIdTipoSexo
    oDOValorPeso.NroDesviacion = lNroDesviacion
    
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set oRsValorPeso = oTabla.SeleccionarPorSexoYDesviacion(oDOValorPeso)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    
    Set DevuelveValorPesoPorSexoDesviacion = oRsValorPeso
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Public Function DevuelveValorTallaPorSexoDesviacion(lIdTipoSexo As Long, lNroDesviacion As Long) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtenInteValorTalla
    Dim oDOValorTalla As New DOAtenInteValorTalla
    Dim oRsValorTalla As New ADODB.Recordset
    
    oDOValorTalla.IdTipoSexo = lIdTipoSexo
    oDOValorTalla.NroDesviacion = lNroDesviacion
    
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set oRsValorTalla = oTabla.SeleccionarPorSexoYDesviacion(oDOValorTalla)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    
    Set DevuelveValorTallaPorSexoDesviacion = oRsValorTalla
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function


Public Function AtencionesCeListaTriaje(lIdAtencion As Long, lNroHistoria As Long) As ADODB.Recordset
On Error GoTo miError
    Dim oConexion As New ADODB.Connection
    Dim oTabla As New SIGHDatos.AtencionesCE
    Dim oDOAtencionesCE As New DOAtencionesCE
    Dim oRsListTriaje As New ADODB.Recordset
    
    oDOAtencionesCE.idAtencion = lIdAtencion
    oDOAtencionesCE.NroHistoriaClinica = lNroHistoria
    
    
'    oConexion.Open sighentidades.ConsultaExterna
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghJamo)
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set oRsListTriaje = oTabla.AtencionesCeListaTriaje(oDOAtencionesCE)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    
    Set AtencionesCeListaTriaje = oRsListTriaje
miError:
    If Err Then
        ms_MensajeError = Err.Number & " : " & Err.Description
    End If
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function
