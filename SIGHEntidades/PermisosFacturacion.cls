VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PermisosFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Procesos Generales de Permisos
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mb_AgregarServicio As Boolean
Dim mb_AgregarBienesInsumos As Boolean
Dim mb_AgregarExoneraciones As Boolean
Dim mb_AutorizarPendientesDePago As Boolean
Dim mb_AgregarPolizasSeguro As Boolean
Dim mb_AutorizarSIS As Boolean
Dim mb_AutorizarSOAT As Boolean
Dim mb_AutorizarConvenios As Boolean
Dim mb_AutorizarDevoluciones As Boolean
Dim mb_AbrirCuentaAtencion As Boolean
Dim mb_CerrarCuentaAtencion As Boolean
Dim mb_ActualizaFechaDocumentoES As Boolean
Property Let ActualizaFechaDocumentoES(bValue As Boolean)
    mb_ActualizaFechaDocumentoES = bValue
End Property

Property Get ActualizaFechaDocumentoES() As Boolean
    ActualizaFechaDocumentoES = mb_ActualizaFechaDocumentoES
End Property

Property Let AgregarServicios(bValue As Boolean)
    mb_AgregarServicio = bValue
End Property

Property Get AgregarServicios() As Boolean
    AgregarServicios = mb_AgregarServicio
End Property

Property Let AgregarBienesInsumos(bValue As Boolean)
    mb_AgregarBienesInsumos = bValue
End Property

Property Get AgregarBienesInsumos() As Boolean
    AgregarBienesInsumos = mb_AgregarBienesInsumos
End Property

Property Let AgregarExoneraciones(bValue As Boolean)
    mb_AgregarExoneraciones = bValue
End Property

Property Get AgregarExoneraciones() As Boolean
    AgregarExoneraciones = mb_AgregarExoneraciones
End Property

Property Let AutorizarPendientesDePago(bValue As Boolean)
    mb_AutorizarPendientesDePago = bValue
End Property

Property Get AutorizarPendientesDePago() As Boolean
    AutorizarPendientesDePago = mb_AutorizarPendientesDePago
End Property

Property Let AgregarPolizasSeguro(bValue As Boolean)
    mb_AgregarPolizasSeguro = bValue
End Property

Property Get AgregarPolizasSeguro() As Boolean
    AgregarPolizasSeguro = mb_AgregarPolizasSeguro
End Property

Property Let AutorizarSIS(bValue As Boolean)
    mb_AutorizarSIS = bValue
End Property

Property Get AutorizarSIS() As Boolean
    AutorizarSIS = mb_AutorizarSIS
End Property

Property Let AutorizarSOAT(bValue As Boolean)
    mb_AutorizarSOAT = bValue
End Property

Property Get AutorizarSOAT() As Boolean
    AutorizarSOAT = mb_AutorizarSOAT
End Property

Property Let AutorizarConvenios(bValue As Boolean)
    mb_AutorizarConvenios = bValue
End Property

Property Get AutorizarConvenios() As Boolean
    AutorizarConvenios = mb_AutorizarConvenios
End Property

Property Let AutorizarDevoluciones(bValue As Boolean)
    mb_AutorizarDevoluciones = bValue
End Property

Property Get AutorizarDevoluciones() As Boolean
    AutorizarDevoluciones = mb_AutorizarDevoluciones
End Property

Property Let AbrirCuentaAtencion(bValue As Boolean)
    mb_AbrirCuentaAtencion = bValue
End Property

Property Get AbrirCuentaAtencion() As Boolean
    AbrirCuentaAtencion = mb_AbrirCuentaAtencion
End Property

Property Let CerrarCuentaAtencion(bValue As Boolean)
    mb_CerrarCuentaAtencion = bValue
End Property

Property Get CerrarCuentaAtencion() As Boolean
    CerrarCuentaAtencion = mb_CerrarCuentaAtencion
End Property

Private Sub Class_Initialize()

    mb_AgregarServicio = False
    mb_AgregarBienesInsumos = False
    mb_AgregarExoneraciones = False
    mb_AutorizarPendientesDePago = False
    mb_AgregarPolizasSeguro = False
    mb_AutorizarSIS = False
    mb_AutorizarSOAT = False
    mb_AutorizarConvenios = False
    mb_AutorizarDevoluciones = False
    mb_ActualizaFechaDocumentoES = False
End Sub
