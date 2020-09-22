VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOTiposFinanciamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla TiposFinanciamiento
'        Programado por: Castro W
'        Fecha: Julio 2004
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim ms_Descripcion As String
Dim ml_IdTipoFinanciamiento As Long
Dim mb_esOficina As Boolean
Dim mb_esSalida As Boolean
Dim mb_SeIngresPrecios As Boolean
Dim mb_EsFarmacia As Boolean
Dim mb_SeImprimeComprobante As Boolean
Dim mb_esFuenteFinanciamiento As Boolean
Dim mb_GeneraPago As Long
Dim ml_idCajaTiposComprobante As Long
Dim ml_idTipoConcepto As Long
Dim ms_tipoVenta As String

Property Let tipoVenta(sValue As String)
   ms_tipoVenta = sValue
End Property
Property Get tipoVenta() As String
   tipoVenta = ms_tipoVenta
End Property

Property Let idTipoConcepto(lValue As Long)
   ml_idTipoConcepto = lValue
End Property
Property Get idTipoConcepto() As Long
   idTipoConcepto = ml_idTipoConcepto
End Property

Property Let idCajaTiposComprobante(lValue As Long)
   ml_idCajaTiposComprobante = lValue
End Property
Property Get idCajaTiposComprobante() As Long
   idCajaTiposComprobante = ml_idCajaTiposComprobante
End Property

Property Let GeneraPago(bValue As Long)
   mb_GeneraPago = bValue
End Property
Property Get GeneraPago() As Long
   GeneraPago = mb_GeneraPago
End Property

Property Let esFuenteFinanciamiento(bValue As Boolean)
   mb_esFuenteFinanciamiento = bValue
End Property
Property Get esFuenteFinanciamiento() As Boolean
   esFuenteFinanciamiento = mb_esFuenteFinanciamiento
End Property

Property Let SeImprimeComprobante(bValue As Boolean)
   mb_SeImprimeComprobante = bValue
End Property
Property Get SeImprimeComprobante() As Boolean
   SeImprimeComprobante = mb_SeImprimeComprobante
End Property

Property Let EsFarmacia(bValue As Boolean)
   mb_EsFarmacia = bValue
End Property
Property Get EsFarmacia() As Boolean
   EsFarmacia = mb_EsFarmacia
End Property

Property Let SeIngresPrecios(bValue As Boolean)
   mb_SeIngresPrecios = bValue
End Property
Property Get SeIngresPrecios() As Boolean
   SeIngresPrecios = mb_SeIngresPrecios
End Property

Property Let esSalida(bValue As Boolean)
   mb_esSalida = bValue
End Property
Property Get esSalida() As Boolean
   esSalida = mb_esSalida
End Property

Property Let esOficina(bValue As Boolean)
   mb_esOficina = bValue
End Property
Property Get esOficina() As Boolean
   esOficina = mb_esOficina
End Property

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property


