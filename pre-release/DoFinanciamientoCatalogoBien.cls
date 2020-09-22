VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFinanciamientoCatalogoBien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactCatalogoBienesInsumosHosp
'        Programado por: Barrantes D
'        Fecha: Mayo 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_Auditoria As Long
Dim mb_Activo As Boolean
Dim ml_IdTipoFinanciamiento As Long
Dim ml_IdProducto As Long
Dim mc_PrecioUnitario As Currency
Dim ml_IdPlanCatalogo As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Activo(bValue As Boolean)
   mb_Activo = bValue
End Property
Property Get Activo() As Boolean
   Activo = mb_Activo
End Property
Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let PrecioUnitario(cValue As Currency)
   mc_PrecioUnitario = cValue
End Property
Property Get PrecioUnitario() As Currency
   PrecioUnitario = mc_PrecioUnitario
End Property
Property Let IdPlanCatalogo(lValue As Long)
   ml_IdPlanCatalogo = lValue
End Property
Property Get IdPlanCatalogo() As Long
   IdPlanCatalogo = ml_IdPlanCatalogo
End Property


