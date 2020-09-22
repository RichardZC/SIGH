VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFactReembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla FactReembolsos
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long
Dim ml_idFactReembolso As Long
Dim ml_Anio As Long
Dim ml_Mes As Long
Dim ml_idAreaTramitaSeguro As Long
Dim ml_idFuenteFinanciamiento As Long
Dim ms_Descripcion As String
Dim mc_SaldoInicial As Currency
Dim mc_ConsumoPorReembolsar As Currency
Dim mc_ReembolsoPagado As Currency
Dim mc_ReembolsoPorPagar As Currency
Dim mc_SaldoFinal As Currency
Dim ms_Documentos As String
Dim ml_idEstadoReembolso As Long
Dim ml_idTipoConsumo As Long
Dim ml_IdTipoComprobante As Long
Dim mb_GrabaDefinitivamente As Boolean

Property Let GrabaDefinitivamente(bValue As Boolean)
   mb_GrabaDefinitivamente = bValue
End Property
Property Get GrabaDefinitivamente() As Boolean
   GrabaDefinitivamente = mb_GrabaDefinitivamente
End Property

Property Let IdTipoComprobante(lValue As Long)
   ml_IdTipoComprobante = lValue
End Property
Property Get IdTipoComprobante() As Long
   IdTipoComprobante = ml_IdTipoComprobante
End Property
Property Let idTipoConsumo(lValue As Long)
   ml_idTipoConsumo = lValue
End Property
Property Get idTipoConsumo() As Long
   idTipoConsumo = ml_idTipoConsumo
End Property
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
Property Let Anio(iValue As Long)
   ml_Anio = iValue
End Property
Property Get Anio() As Long
   Anio = ml_Anio
End Property
Property Let Mes(iValue As Long)
   ml_Mes = iValue
End Property
Property Get Mes() As Long
   Mes = ml_Mes
End Property
Property Let idAreaTramitaSeguro(iValue As Long)
   ml_idAreaTramitaSeguro = iValue
End Property
Property Get idAreaTramitaSeguro() As Long
   idAreaTramitaSeguro = ml_idAreaTramitaSeguro
End Property
Property Let idFuenteFinanciamiento(iValue As Long)
   ml_idFuenteFinanciamiento = iValue
End Property
Property Get idFuenteFinanciamiento() As Long
   idFuenteFinanciamiento = ml_idFuenteFinanciamiento
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let SaldoInicial(cValue As Currency)
   mc_SaldoInicial = cValue
End Property
Property Get SaldoInicial() As Currency
   SaldoInicial = mc_SaldoInicial
End Property
Property Let ConsumoPorReembolsar(cValue As Currency)
   mc_ConsumoPorReembolsar = cValue
End Property
Property Get ConsumoPorReembolsar() As Currency
   ConsumoPorReembolsar = mc_ConsumoPorReembolsar
End Property
Property Let ReembolsoPagado(cValue As Currency)
   mc_ReembolsoPagado = cValue
End Property
Property Get ReembolsoPagado() As Currency
   ReembolsoPagado = mc_ReembolsoPagado
End Property
Property Let ReembolsoPorPagar(cValue As Currency)
   mc_ReembolsoPorPagar = cValue
End Property
Property Get ReembolsoPorPagar() As Currency
   ReembolsoPorPagar = mc_ReembolsoPorPagar
End Property
Property Let SaldoFinal(cValue As Currency)
   mc_SaldoFinal = cValue
End Property
Property Get SaldoFinal() As Currency
   SaldoFinal = mc_SaldoFinal
End Property
Property Let Documentos(sValue As String)
   ms_Documentos = sValue
End Property
Property Get Documentos() As String
   Documentos = ms_Documentos
End Property
Property Let idEstadoReembolso(iValue As Long)
   ml_idEstadoReembolso = iValue
End Property
Property Get idEstadoReembolso() As Long
   idEstadoReembolso = ml_idEstadoReembolso
End Property
