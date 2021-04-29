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
Dim ml_cantidad As Long
Dim mc_precio As Currency
Dim mc_Total As Currency
'mgaray201411a
Dim ms_LabConfHIS As String
Dim ml_GrupoHIS As Long
Dim ml_SubGrupoHIS As Long
'SCCQ 23-04-2021 Cambio 64 Inicio (Reversion Cambio 46)
'RHA 17/11/2020 Cambio 46 Inicio
'Dim ml_IdFinan As Long
'Dim ml_CodPres As String
'Property Get IdFinan() As Long
'   IdFinan = ml_IdFinan
'End Property
'Property Let IdFinan(lValue As Long)
'   ml_IdFinan = lValue
'End Property
'Property Get CodPres() As String
'   CodPres = ml_CodPres
'End Property
'Property Let CodPres(lValue As String)
'   ml_CodPres = lValue
'End Property
'RHA 17/11/2020 Cambio 46 Fin
'SCCQ 23-04-2021 Cambio 64 Fin (Reversion Cambio 46)

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
Property Let cantidad(iValue As Long)
   ml_cantidad = iValue
End Property
Property Get cantidad() As Long
   cantidad = ml_cantidad
End Property
Property Let precio(cValue As Currency)
   mc_precio = cValue
End Property
Property Get precio() As Currency
   precio = mc_precio
End Property
Property Let TOTAL(cValue As Currency)
   mc_Total = cValue
End Property
Property Get TOTAL() As Currency
   TOTAL = mc_Total
End Property
'mgaray201411a
Property Let labConfHIS(sValue As String)
   ms_LabConfHIS = sValue
End Property
Property Get labConfHIS() As String
   labConfHIS = ms_LabConfHIS
End Property
