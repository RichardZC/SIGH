VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bioquimica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Resultado de bioquímica
'        Programado por: Bonilla A
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_IdMovimiento As New frmBioquimica
Dim ml_idUsuario As Long
Dim ml_idOrden As Long
Dim ml_idPrueba As String
Dim ml_nombrePrueba As String
Dim ml_idPaciente As Long
Dim ml_idAnalisis As Long
Dim ml_nombreMedico As String
Dim ml_nombrePaciente As String
Dim ml_areaTrabajo As Long
Dim ml_CodigoPruebaSeleccionada As String
Dim ml_DetalleOrden As New ADODB.Recordset
Dim ml_idOrdenLab As Long
Dim ml_FechaNacimiento As Date
Dim ml_idTipoSexo As Long
Dim ml_NoMuestraBotonGrabar As Boolean
Dim mo_lcNombrePc As String
Dim mo_lnIdTablaLISTBARITEMS  As Long
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Let NoMuestraBotonGrabar(lValue As Boolean)
   ml_NoMuestraBotonGrabar = lValue
End Property


Property Let idTipoSexo(lValue As Long)
    ml_idTipoSexo = lValue
End Property
Property Let FechaNacimiento(lValue As Date)
    ml_FechaNacimiento = lValue
End Property


Property Let idOrdenLab(lValue As Long)
   ml_idOrdenLab = lValue
End Property

Property Let AreaTrabajo(lValue As Long)
    ml_areaTrabajo = lValue
End Property

Sub MostrarFormulario()
  mo_IdMovimiento.idUsuario = ml_idUsuario
  mo_IdMovimiento.idOrden = ml_idOrden
  mo_IdMovimiento.idPrueba = ml_idPrueba
  mo_IdMovimiento.nombrePrueba = ml_nombrePrueba
  mo_IdMovimiento.idAnalisis = ml_idAnalisis
  mo_IdMovimiento.idPaciente = ml_idPaciente
  mo_IdMovimiento.nombreMedico = ml_nombreMedico
  mo_IdMovimiento.nombrePaciente = ml_nombrePaciente
  mo_IdMovimiento.AreaTrabajo = ml_areaTrabajo
  mo_IdMovimiento.CodigoPruebaSeleccionada = ml_CodigoPruebaSeleccionada
  mo_IdMovimiento.DetalleOrden = ml_DetalleOrden
  mo_IdMovimiento.idOrdenLab = ml_idOrdenLab
  mo_IdMovimiento.FechaNacimiento = ml_FechaNacimiento
  mo_IdMovimiento.idTipoSexo = ml_idTipoSexo
  mo_IdMovimiento.NoMuestraBotonGrabar = ml_NoMuestraBotonGrabar
  mo_IdMovimiento.lcNombrePc = mo_lcNombrePc
  mo_IdMovimiento.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
  mo_IdMovimiento.Show vbModal
  Set mo_IdMovimiento = Nothing
End Sub

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property

Property Let idPrueba(lValue As String)
   ml_idPrueba = lValue
End Property

Property Let idOrden(lValue As Long)
   ml_idOrden = lValue
End Property

Property Let nombrePrueba(lValue As String)
   ml_nombrePrueba = lValue
End Property

Property Let idAnalisis(lValue As Long)
   ml_idAnalisis = lValue
End Property

Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property

Property Let nombreMedico(lValue As String)
   ml_nombreMedico = lValue
End Property

Property Let nombrePaciente(lValue As String)
   ml_nombrePaciente = lValue
End Property

Property Let CodigoPruebaSeleccionada(lValue As String)
   ml_CodigoPruebaSeleccionada = lValue
End Property

Property Let DetalleOrden(lValue As ADODB.Recordset)
   Set ml_DetalleOrden = lValue
End Property
