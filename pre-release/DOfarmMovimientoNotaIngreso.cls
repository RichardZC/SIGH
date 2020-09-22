VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOfarmMovimientoNotaIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FarmMovimientoNotaIngreso
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdUsuarioAuditoria As Long
Dim ms_movNumero As String
Dim ms_movTipo As String
Dim mda_DocumentoFechaRecepcion As Date
Dim ml_OrigenIdTipo As Long
Dim ms_OrigenNumero As String
Dim mda_OrigenFecha As Date
Dim ml_idProveedor As Long
Dim ml_idTipoCompra As Long
Dim ml_idTipoProceso As Long
Dim ms_NumeroProceso As String
Dim ml_IdPaciente As Long
Dim mda_FechaModificacion As Date
Dim ml_idUsuarioModifica As Long
Dim ml_idComprobantePago As Long
Dim ml_idCuentaAtencion As Long
Dim ml_idFuenteFinanciamiento As Long


Property Let idFuenteFinanciamiento(lValue As Long)
   ml_idFuenteFinanciamiento = lValue
End Property
Property Get idFuenteFinanciamiento() As Long
   idFuenteFinanciamiento = ml_idFuenteFinanciamiento
End Property


Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property


Property Let idComprobantePago(lValue As Long)
   ml_idComprobantePago = lValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
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
Property Let DocumentoFechaRecepcion(daValue As Date)
   mda_DocumentoFechaRecepcion = daValue
End Property
Property Get DocumentoFechaRecepcion() As Date
   DocumentoFechaRecepcion = mda_DocumentoFechaRecepcion
End Property
Property Let OrigenIdTipo(iValue As Long)
   ml_OrigenIdTipo = iValue
End Property
Property Get OrigenIdTipo() As Long
   OrigenIdTipo = ml_OrigenIdTipo
End Property
Property Let OrigenNumero(sValue As String)
   ms_OrigenNumero = sValue
End Property
Property Get OrigenNumero() As String
   OrigenNumero = ms_OrigenNumero
End Property
Property Let OrigenFecha(daValue As Date)
   mda_OrigenFecha = daValue
End Property
Property Get OrigenFecha() As Date
   OrigenFecha = mda_OrigenFecha
End Property
Property Let idProveedor(iValue As Long)
   ml_idProveedor = iValue
End Property
Property Get idProveedor() As Long
   idProveedor = ml_idProveedor
End Property

Property Let idTipoCompra(iValue As Long)
   ml_idTipoCompra = iValue
End Property
Property Get idTipoCompra() As Long
   idTipoCompra = ml_idTipoCompra
End Property
Property Let idTipoProceso(iValue As Long)
   ml_idTipoProceso = iValue
End Property
Property Get idTipoProceso() As Long
   idTipoProceso = ml_idTipoProceso
End Property
Property Let NumeroProceso(sValue As String)
   ms_NumeroProceso = sValue
End Property
Property Get NumeroProceso() As String
   NumeroProceso = ms_NumeroProceso
End Property
Property Let IdPaciente(iValue As Long)
   ml_IdPaciente = iValue
End Property
Property Get IdPaciente() As Long
   IdPaciente = ml_IdPaciente
End Property
Property Let FechaModificacion(daValue As Date)
   mda_FechaModificacion = daValue
End Property
Property Get FechaModificacion() As Date
   FechaModificacion = mda_FechaModificacion
End Property
Property Let idUsuarioModifica(iValue As Long)
   ml_idUsuarioModifica = iValue
End Property
Property Get idUsuarioModifica() As Long
   idUsuarioModifica = ml_idUsuarioModifica
End Property

