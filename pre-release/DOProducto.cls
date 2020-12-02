VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla Productos
'        Programado por: Barrantes D
'        Fecha: Diciembre 2013
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim ml_IdCategoriaProducto As Long
Dim mb_Bloqueado As Boolean
Dim mc_PrecioBase As Currency
Dim ms_Nombre As String
Dim ml_IdProducto As Long
Dim ms_Codigo As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let IdCategoriaProducto(lValue As Long)
   ml_IdCategoriaProducto = lValue
End Property
Property Get IdCategoriaProducto() As Long
   IdCategoriaProducto = ml_IdCategoriaProducto
End Property
Property Let Bloqueado(bValue As Boolean)
   mb_Bloqueado = bValue
End Property
Property Get Bloqueado() As Boolean
   Bloqueado = mb_Bloqueado
End Property
Property Let PrecioBase(cValue As Currency)
   mc_PrecioBase = cValue
End Property
Property Get PrecioBase() As Currency
   PrecioBase = mc_PrecioBase
End Property
Property Let Nombre(sValue As String)
   ms_Nombre = sValue
End Property
Property Get Nombre() As String
   Nombre = ms_Nombre
End Property
Property Let IdProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get IdProducto() As Long
   IdProducto = ml_IdProducto
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property


