VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFacturacionReembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FacturacionReembolsos
'        Programado por: Barrantes D
'        Fecha: Abril 2010
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ml_IdUsuarioAuditoria As Long
Dim ml_idFactReembolso As Long
Dim ml_idCuentaAtencion As Long
Dim mc_ConsumoPorReembolsar As Currency
Dim mc_ReembolsoPorPagar As Currency
Dim mc_ReembolsoPagadoFarmacia As Currency
Dim mc_ReembolsoPagadoServicio As Currency
Dim ms_IdReembolsosAnteriores As String
Dim ml_IdDiagnostico As Long
Dim ms_NroReferenciaDestino As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idFactReembolso(iValue As Long)
   ml_idFactReembolso = iValue
End Property
Property Get idFactReembolso() As Long
   idFactReembolso = ml_idFactReembolso
End Property
Property Let idCuentaAtencion(iValue As Long)
   ml_idCuentaAtencion = iValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let ConsumoPorReembolsar(cValue As Currency)
   mc_ConsumoPorReembolsar = cValue
End Property
Property Get ConsumoPorReembolsar() As Currency
   ConsumoPorReembolsar = mc_ConsumoPorReembolsar
End Property
Property Let ReembolsoPorPagar(cValue As Currency)
   mc_ReembolsoPorPagar = cValue
End Property
Property Get ReembolsoPorPagar() As Currency
   ReembolsoPorPagar = mc_ReembolsoPorPagar
End Property
Property Let ReembolsoPagadoFarmacia(cValue As Currency)
   mc_ReembolsoPagadoFarmacia = cValue
End Property
Property Get ReembolsoPagadoFarmacia() As Currency
   ReembolsoPagadoFarmacia = mc_ReembolsoPagadoFarmacia
End Property
Property Let ReembolsoPagadoServicio(cValue As Currency)
   mc_ReembolsoPagadoServicio = cValue
End Property
Property Get ReembolsoPagadoServicio() As Currency
   ReembolsoPagadoServicio = mc_ReembolsoPagadoServicio
End Property
Property Let IdReembolsosAnteriores(sValue As String)
   ms_IdReembolsosAnteriores = sValue
End Property
Property Get IdReembolsosAnteriores() As String
   IdReembolsosAnteriores = ms_IdReembolsosAnteriores
End Property
Property Let IdDiagnostico(iValue As Long)
   ml_IdDiagnostico = iValue
End Property
Property Get IdDiagnostico() As Long
   IdDiagnostico = ml_IdDiagnostico
End Property
Property Let NroReferenciaDestino(sValue As String)
   ms_NroReferenciaDestino = sValue
End Property
Property Get NroReferenciaDestino() As String
   NroReferenciaDestino = ms_NroReferenciaDestino
End Property
