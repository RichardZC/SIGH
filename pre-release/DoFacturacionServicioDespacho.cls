VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFacturacionServicioDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FacturacionServicioDespacho
'        Programado por: Barrantes D
'        Fecha: Abril 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_IdOrden As Long
Dim ml_idProducto As Long
Dim ml_Cantidad As Long
Dim mc_Precio As Currency
Dim mc_Total As Currency
'mgaray201411a
Dim ms_LabConfHIS As String
Dim ml_GrupoHIS As Long
Dim ml_SubGrupoHIS As Long

Property Get GrupoHIS() As Long
   GrupoHIS = ml_GrupoHIS
End Property
Property Let GrupoHIS(lValue As Long)
   ml_GrupoHIS = lValue
End Property
Property Get SubGrupoHIS() As Long
   SubGrupoHIS = ml_SubGrupoHIS
End Property
Property Let SubGrupoHIS(lValue As Long)
   ml_SubGrupoHIS = lValue
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let IdOrden(iValue As Long)
   ml_IdOrden = iValue
End Property
Property Get IdOrden() As Long
   IdOrden = ml_IdOrden
End Property
Property Let idProducto(iValue As Long)
   ml_idProducto = iValue
End Property
Property Get idProducto() As Long
   idProducto = ml_idProducto
End Property
Property Let Cantidad(iValue As Long)
   ml_Cantidad = iValue
End Property
Property Get Cantidad() As Long
   Cantidad = ml_Cantidad
End Property
Property Let Precio(cValue As Currency)
   mc_Precio = cValue
End Property
Property Get Precio() As Currency
   Precio = mc_Precio
End Property
Property Let Total(cValue As Currency)
   mc_Total = cValue
End Property
Property Get Total() As Currency
   Total = mc_Total
End Property
'mgaray201411a
Property Let labConfHIS(sValue As String)
   ms_LabConfHIS = sValue
End Property
Property Get labConfHIS() As String
   labConfHIS = ms_LabConfHIS
End Property
