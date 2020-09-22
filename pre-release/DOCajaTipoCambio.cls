VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla CajaTipoCambio
'        Programado por: Castro W
'        Fecha: Enero 2005
'
'------------------------------------------------------------------------------------

Option Explicit

Dim ml_Auditoria As Long
Dim mc_TipoCambio As Currency
Dim mda_fecha As Date
Dim ml_IdTipoMoneda As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let TipoCambio(cValue As Currency)
   mc_TipoCambio = cValue
End Property
Property Get TipoCambio() As Currency
   TipoCambio = mc_TipoCambio
End Property
Property Let fecha(daValue As Date)
   mda_fecha = daValue
End Property
Property Get fecha() As Date
   fecha = mda_fecha
End Property
Property Let IdTipoMoneda(lValue As Long)
   ml_IdTipoMoneda = lValue
End Property
Property Get IdTipoMoneda() As Long
   IdTipoMoneda = ml_IdTipoMoneda
End Property

