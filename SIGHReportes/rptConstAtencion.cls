VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rptConstAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim ml_idUsuario As Long
Dim ml_IdMovimiento As Long
Dim ml_Historia As Long
Dim mi_Opcion As sghOpciones
Dim ml_idAtencion As Long
Dim ml_idTipoConstancia As Long
Dim ml_Recibo As String
Dim ml_Observaciones As String
Dim ml_idServicio As Long

Property Get Opcion() As sghOpciones
    Opcion = mi_Opcion
End Property

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Let Opcion(lValue As sghOpciones)
   mi_Opcion = lValue
End Property

Property Let IdMovimiento(lValue As Long)
  ml_IdMovimiento = lValue
End Property

Property Let Historia(lValue As Long)
  ml_Historia = lValue
End Property

Property Let idAtencion(lValue As Long)
  ml_idAtencion = lValue
End Property

Property Let idTipoConstancia(lValue As Long)
  ml_idTipoConstancia = lValue
End Property

Property Let Recibo(lValue As String)
  ml_Recibo = lValue
End Property

Property Let Observaciones(lValue As String)
  ml_Observaciones = lValue
End Property

Property Let idServicio(lValue As Long)
  ml_idServicio = lValue
End Property

Sub EjecutaFormulario()
  Dim oFormulario As New frmConstancias
  oFormulario.idUsuario = ml_idUsuario
  oFormulario.Opcion = mi_Opcion
  oFormulario.Historia = ml_Historia
  oFormulario.idAtencion = ml_idAtencion
  oFormulario.idTipoConstancia = ml_idTipoConstancia
  oFormulario.Observaciones = ml_Observaciones
  oFormulario.idServicio = ml_idServicio
  oFormulario.Recibo = ml_Recibo
  oFormulario.Show vbModal
End Sub
