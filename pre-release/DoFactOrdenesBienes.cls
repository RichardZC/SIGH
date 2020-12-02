VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFactOrdenesBienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactOrdenesBienes
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_idOrden As Long
Dim ml_idpuntocarga As Long
Dim ml_IdPaciente As Long
Dim ml_IdCuentaAtencion As Long
Dim ml_IdComprobantePago As Long
Dim ms_movNumero As String
Dim ms_movTipo As String
Dim ml_idPreventa As Long
Dim mda_FechaCreacion As Date
Dim ml_IdUsuario As Long
Dim ml_IdEstadoFacturacion As Long
Dim mc_ImporteExonerado As Currency
Dim ml_idUsuarioExonera As Long
Dim ms_Descripcion As String
Dim ms_Paquete As String
Dim ms_DNI As String
Dim ms_NombPaciente As String

Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let Paquete(sValue As String)
   ms_Paquete = sValue
End Property
Property Get Paquete() As String
   Paquete = ms_Paquete
End Property
Property Let DNI(sValue As String)
   ms_DNI = sValue
End Property
Property Get DNI() As String
   DNI = ms_DNI
End Property

Property Let NombPaciente(sValue As String)
   ms_NombPaciente = sValue
End Property
Property Get NombPaciente() As String
   NombPaciente = ms_NombPaciente
End Property

Property Let idUsuarioExonera(lValue As Long)
   ml_idUsuarioExonera = lValue
End Property
Property Get idUsuarioExonera() As Long
   idUsuarioExonera = ml_idUsuarioExonera
End Property


Property Let ImporteExonerado(cValue As Currency)
   mc_ImporteExonerado = cValue
End Property
Property Get ImporteExonerado() As Currency
   ImporteExonerado = mc_ImporteExonerado
End Property



Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idOrden(iValue As Long)
   ml_idOrden = iValue
End Property
Property Get idOrden() As Long
   idOrden = ml_idOrden
End Property
Property Let idpuntocarga(iValue As Long)
   ml_idpuntocarga = iValue
End Property
Property Get idpuntocarga() As Long
   idpuntocarga = ml_idpuntocarga
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let IdCuentaAtencion(iValue As Long)
   ml_IdCuentaAtencion = iValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let IdComprobantePago(iValue As Long)
   ml_IdComprobantePago = iValue
End Property
Property Get IdComprobantePago() As Long
   IdComprobantePago = ml_IdComprobantePago
End Property
Property Let movNumero(sValue As String)
   ms_movNumero = sValue
End Property
Property Get movNumero() As String
   movNumero = ms_movNumero
End Property
Property Let movTipo(sValue As String)
   ms_movTipo = sValue
End Property
Property Get movTipo() As String
   movTipo = ms_movTipo
End Property
Property Let idPreventa(iValue As Long)
   ml_idPreventa = iValue
End Property
Property Get idPreventa() As Long
   idPreventa = ml_idPreventa
End Property

Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property
Property Let IdUsuario(iValue As Long)
   ml_IdUsuario = iValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdEstadoFacturacion(iValue As Long)
   ml_IdEstadoFacturacion = iValue
End Property
Property Get IdEstadoFacturacion() As Long
   IdEstadoFacturacion = ml_IdEstadoFacturacion
End Property


