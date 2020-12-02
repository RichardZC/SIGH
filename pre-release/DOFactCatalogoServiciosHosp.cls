VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOFactCatalogoServiciosHosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactCatalogoServiciosHosp
'        Programado por: Barrantes D
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdFinanciamientoCatalogo As Long
Dim mc_PrecioUnitario As Currency
Dim ml_IdProducto As Long
Dim ml_IdTipoFinanciamiento As Long
Dim mb_Activo As Boolean
Dim mb_SeUsaSinPrecio As Boolean
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdFinanciamientoCatalogo(iValue As Long)
   ml_IdFinanciamientoCatalogo = iValue
End Property
Property Get IdFinanciamientoCatalogo() As Long
   IdFinanciamientoCatalogo = ml_IdFinanciamientoCatalogo
End Property
Property Let PrecioUnitario(cValue As Currency)
   mc_PrecioUnitario = cValue
End Property
Property Get PrecioUnitario() As Currency
   PrecioUnitario = mc_PrecioUnitario
End Property
Property Let IdProducto(iValue As Long)
   ml_IdProducto = iValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let IdTipoFinanciamiento(iValue As Long)
   ml_IdTipoFinanciamiento = iValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let Activo(bValue As Boolean)
   mb_Activo = bValue
End Property
Property Get Activo() As Boolean
   Activo = mb_Activo
End Property
Property Let SeUsaSinPrecio(bValue As Boolean)
   mb_SeUsaSinPrecio = bValue
End Property
Property Get SeUsaSinPrecio() As Boolean
   SeUsaSinPrecio = mb_SeUsaSinPrecio
End Property

