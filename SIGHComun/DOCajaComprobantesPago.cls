VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaComprobantesPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla CajaComprobantesPago
'        Programado por: Castro W
'        Fecha: Agosto 2006
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ml_Auditoria As Long
Dim mc_TipoCambio As Currency
Dim ms_NroSerie As String
Dim ms_NroDocumento As String
Dim ms_RazonSocial As String
Dim ms_RUC As String
Dim ms_DNI As String
Dim mc_SubTotal As Currency
Dim mc_IGV As Currency
Dim ml_idComprobantePago As Long
Dim mda_FechaCobranza As Date
Dim ml_IdTipoOrden As Long
Dim ms_Observaciones As String
Dim ml_IdTipoComprobante As Long
Dim ml_idCuentaAtencion As Long
Dim ml_IdEstadoComprobante As Long
Dim ml_IdGestionCaja As Long
Dim ml_IdTipoPago As Long
Dim mc_Total As Currency

'**************daniel barrantes**************
'***************Se registran los Descuentos que haiga en la Boleta, como ADELANTOS DE PAGO
'***************Registro de Boletas con ID del Paciente y la Forma de Pago
Dim mc_Dctos As Currency
Dim ml_idPaciente As Long
Dim ml_IdFormaPago As Long
Dim ml_idFarmacia As Long
Dim ml_IdTurno As Long
Dim ml_idCaja As Long
Dim ml_IdCajero As Long
Dim mc_Exoneraciones As Currency
Dim mc_Adelantos As Currency
Dim ml_IdTipoFinanciamiento As Long
Dim ms_TieneCredito As String
Dim mda_fechaEmision As Date
Dim mc_SunatTotalIgv As Currency
Dim mc_SunatOpeInafectas As Currency

Property Let SunatTotalIgv(cValue As Currency)
   mc_SunatTotalIgv = cValue
End Property
Property Get SunatTotalIgv() As Currency
   SunatTotalIgv = mc_SunatTotalIgv
End Property
Property Let SunatOpeInafectas(cValue As Currency)
   mc_SunatOpeInafectas = cValue
End Property
Property Get SunatOpeInafectas() As Currency
    SunatOpeInafectas = mc_SunatOpeInafectas
End Property


Property Let fechaEmision(daValue As Date)
   mda_fechaEmision = daValue
End Property
Property Get fechaEmision() As Date
   fechaEmision = mda_fechaEmision
End Property

Property Let TieneCredito(sValue As String)
   ms_TieneCredito = sValue
End Property
Property Get TieneCredito() As String
   TieneCredito = ms_TieneCredito
End Property

Property Let IdTipoFinanciamiento(lValue As Long)
   ml_IdTipoFinanciamiento = lValue
End Property
Property Get IdTipoFinanciamiento() As Long
   IdTipoFinanciamiento = ml_IdTipoFinanciamiento
End Property


Property Let Adelantos(cValue As Currency)
   mc_Adelantos = cValue
End Property
Property Get Adelantos() As Currency
   Adelantos = mc_Adelantos
End Property
Property Let Exoneraciones(cValue As Currency)
   mc_Exoneraciones = cValue
End Property
Property Get Exoneraciones() As Currency
   Exoneraciones = mc_Exoneraciones
End Property

Property Let IdCajero(lValue As Long)
   ml_IdCajero = lValue
End Property
Property Get IdCajero() As Long
   IdCajero = ml_IdCajero
End Property
Property Let idCaja(lValue As Long)
   ml_idCaja = lValue
End Property
Property Get idCaja() As Long
   idCaja = ml_idCaja
End Property

Property Let IdTurno(lValue As Long)
   ml_IdTurno = lValue
End Property
Property Get IdTurno() As Long
   IdTurno = ml_IdTurno
End Property



Property Let idFarmacia(lValue As Long)
   ml_idFarmacia = lValue
End Property
Property Get idFarmacia() As Long
   idFarmacia = ml_idFarmacia
End Property
Property Let IdFormaPago(lValue As Long)
   ml_IdFormaPago = lValue
End Property
Property Get IdFormaPago() As Long
   IdFormaPago = ml_IdFormaPago
End Property
Property Let idPaciente(lValue As Long)
   ml_idPaciente = lValue
End Property
Property Get idPaciente() As Long
   idPaciente = ml_idPaciente
End Property

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
Property Let NroSerie(sValue As String)
   ms_NroSerie = sValue
End Property
Property Get NroSerie() As String
   NroSerie = ms_NroSerie
End Property
Property Let NroDocumento(sValue As String)
   ms_NroDocumento = sValue
End Property
Property Get NroDocumento() As String
   NroDocumento = ms_NroDocumento
End Property
Property Let RazonSocial(sValue As String)
   ms_RazonSocial = sValue
End Property
Property Get RazonSocial() As String
   RazonSocial = ms_RazonSocial
End Property
Property Let RUC(sValue As String)
   ms_RUC = sValue
End Property
Property Get RUC() As String
   RUC = ms_RUC
End Property
Property Let DNI(sValue As String)
   ms_DNI = sValue
End Property
Property Get DNI() As String
   DNI = ms_DNI
End Property
Property Let SubTotal(cValue As Currency)
   mc_SubTotal = cValue
End Property
Property Get SubTotal() As Currency
   SubTotal = mc_SubTotal
End Property
Property Let IGV(cValue As Currency)
   mc_IGV = cValue
End Property
Property Get IGV() As Currency
   IGV = mc_IGV
End Property
Property Let idComprobantePago(lValue As Long)
   ml_idComprobantePago = lValue
End Property
Property Get idComprobantePago() As Long
   idComprobantePago = ml_idComprobantePago
End Property
Property Let FechaCobranza(daValue As Date)
   mda_FechaCobranza = daValue
End Property
Property Get FechaCobranza() As Date
   FechaCobranza = mda_FechaCobranza
End Property
Property Let IdTipoOrden(lValue As Long)
   ml_IdTipoOrden = lValue
End Property
Property Get IdTipoOrden() As Long
   IdTipoOrden = ml_IdTipoOrden
End Property
Property Let Observaciones(sValue As String)
   ms_Observaciones = sValue
End Property
Property Get Observaciones() As String
   Observaciones = ms_Observaciones
End Property
Property Let IdTipoComprobante(lValue As Long)
   ml_IdTipoComprobante = lValue
End Property
Property Get IdTipoComprobante() As Long
   IdTipoComprobante = ml_IdTipoComprobante
End Property
Property Let idCuentaAtencion(lValue As Long)
   ml_idCuentaAtencion = lValue
End Property
Property Get idCuentaAtencion() As Long
   idCuentaAtencion = ml_idCuentaAtencion
End Property
Property Let IdEstadoComprobante(lValue As Long)
   ml_IdEstadoComprobante = lValue
End Property
Property Get IdEstadoComprobante() As Long
   IdEstadoComprobante = ml_IdEstadoComprobante
End Property
Property Let IdGestionCaja(lValue As Long)
   ml_IdGestionCaja = lValue
End Property
Property Get IdGestionCaja() As Long
   IdGestionCaja = ml_IdGestionCaja
End Property
Property Let IdTipoPago(lValue As Long)
   ml_IdTipoPago = lValue
End Property
Property Get IdTipoPago() As Long
   IdTipoPago = ml_IdTipoPago
End Property
Property Let TOTAL(cValue As Currency)
   mc_Total = cValue
End Property
Property Get TOTAL() As Currency
   TOTAL = mc_Total
End Property
Property Let Dctos(cValue As Currency)
   mc_Dctos = cValue
End Property
Property Get Dctos() As Currency
   Dctos = mc_Dctos
End Property


