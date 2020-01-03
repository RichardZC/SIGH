VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSunatResumenDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: Sunat_ResumenDiario
'        Autor: Daniel Enrique Barrantes
'        Fecha: 28/02/2019 07:04:40 p.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_id As Long
Dim ms_DctoRDI As String
Dim ms_fechaEmision As String
Dim ms_fechaResumen As String
Dim ms_tipoDcto As String
Dim ms_SerieDocumento As String
Dim ms_PacienteTipoDoc As String
Dim ms_PacienteNumeroDoc As String
Dim ms_Moneda As String
Dim ms_OpeGravadas As String
Dim ms_OpeExoneradas As String
Dim ms_OpeInafectas As String
Dim ms_OpeGratuitas As String
Dim ms_OtrosCargos As String
Dim ms_TotalIsc As String
Dim ms_TotalIgv As String
Dim ms_TotalOtros As String
Dim ms_ImporteVenta As String
Dim ms_ModTipoDcto As String
Dim ms_ModSerie As String
Dim ms_ModDocumento As String
Dim ms_PercepRegimen As String
Dim ms_PercepPorcen As String
Dim ms_PercepBaseImp As String
Dim ms_PercepMonto As String
Dim ms_PercepCobrar As String
Dim ms_Estado As String
Dim ms_DctoSunat As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let id(iValue As Long)
   ml_id = iValue
End Property
Property Get id() As Long
   id = ml_id
End Property
Property Let DctoRDI(sValue As String)
   ms_DctoRDI = sValue
End Property
Property Get DctoRDI() As String
   DctoRDI = ms_DctoRDI
End Property
Property Let fechaEmision(sValue As String)
   ms_fechaEmision = sValue
End Property
Property Get fechaEmision() As String
   fechaEmision = ms_fechaEmision
End Property
Property Let fechaResumen(sValue As String)
   ms_fechaResumen = sValue
End Property
Property Get fechaResumen() As String
   fechaResumen = ms_fechaResumen
End Property
Property Let tipoDcto(sValue As String)
   ms_tipoDcto = sValue
End Property
Property Get tipoDcto() As String
   tipoDcto = ms_tipoDcto
End Property
Property Let SerieDocumento(sValue As String)
   ms_SerieDocumento = sValue
End Property
Property Get SerieDocumento() As String
   SerieDocumento = ms_SerieDocumento
End Property
Property Let PacienteTipoDoc(sValue As String)
   ms_PacienteTipoDoc = sValue
End Property
Property Get PacienteTipoDoc() As String
   PacienteTipoDoc = ms_PacienteTipoDoc
End Property
Property Let PacienteNumeroDoc(sValue As String)
   ms_PacienteNumeroDoc = sValue
End Property
Property Get PacienteNumeroDoc() As String
   PacienteNumeroDoc = ms_PacienteNumeroDoc
End Property
Property Let Moneda(sValue As String)
   ms_Moneda = sValue
End Property
Property Get Moneda() As String
   Moneda = ms_Moneda
End Property
Property Let OpeGravadas(sValue As String)
   ms_OpeGravadas = sValue
End Property
Property Get OpeGravadas() As String
   OpeGravadas = ms_OpeGravadas
End Property
Property Let OpeExoneradas(sValue As String)
   ms_OpeExoneradas = sValue
End Property
Property Get OpeExoneradas() As String
   OpeExoneradas = ms_OpeExoneradas
End Property
Property Let OpeInafectas(sValue As String)
   ms_OpeInafectas = sValue
End Property
Property Get OpeInafectas() As String
   OpeInafectas = ms_OpeInafectas
End Property
Property Let OpeGratuitas(sValue As String)
   ms_OpeGratuitas = sValue
End Property
Property Get OpeGratuitas() As String
   OpeGratuitas = ms_OpeGratuitas
End Property
Property Let OtrosCargos(sValue As String)
   ms_OtrosCargos = sValue
End Property
Property Get OtrosCargos() As String
   OtrosCargos = ms_OtrosCargos
End Property
Property Let TotalIsc(sValue As String)
   ms_TotalIsc = sValue
End Property
Property Get TotalIsc() As String
   TotalIsc = ms_TotalIsc
End Property
Property Let TotalIgv(sValue As String)
   ms_TotalIgv = sValue
End Property
Property Get TotalIgv() As String
   TotalIgv = ms_TotalIgv
End Property
Property Let TotalOtros(sValue As String)
   ms_TotalOtros = sValue
End Property
Property Get TotalOtros() As String
   TotalOtros = ms_TotalOtros
End Property
Property Let ImporteVenta(sValue As String)
   ms_ImporteVenta = sValue
End Property
Property Get ImporteVenta() As String
   ImporteVenta = ms_ImporteVenta
End Property
Property Let ModTipoDcto(sValue As String)
   ms_ModTipoDcto = sValue
End Property
Property Get ModTipoDcto() As String
   ModTipoDcto = ms_ModTipoDcto
End Property
Property Let ModSerie(sValue As String)
   ms_ModSerie = sValue
End Property
Property Get ModSerie() As String
   ModSerie = ms_ModSerie
End Property
Property Let ModDocumento(sValue As String)
   ms_ModDocumento = sValue
End Property
Property Get ModDocumento() As String
   ModDocumento = ms_ModDocumento
End Property
Property Let PercepRegimen(sValue As String)
   ms_PercepRegimen = sValue
End Property
Property Get PercepRegimen() As String
   PercepRegimen = ms_PercepRegimen
End Property
Property Let PercepPorcen(sValue As String)
   ms_PercepPorcen = sValue
End Property
Property Get PercepPorcen() As String
   PercepPorcen = ms_PercepPorcen
End Property
Property Let PercepBaseImp(sValue As String)
   ms_PercepBaseImp = sValue
End Property
Property Get PercepBaseImp() As String
   PercepBaseImp = ms_PercepBaseImp
End Property
Property Let PercepMonto(sValue As String)
   ms_PercepMonto = sValue
End Property
Property Get PercepMonto() As String
   PercepMonto = ms_PercepMonto
End Property
Property Let PercepCobrar(sValue As String)
   ms_PercepCobrar = sValue
End Property
Property Get PercepCobrar() As String
   PercepCobrar = ms_PercepCobrar
End Property
Property Let Estado(sValue As String)
   ms_Estado = sValue
End Property
Property Get Estado() As String
   Estado = ms_Estado
End Property
Property Let DctoSunat(sValue As String)
   ms_DctoSunat = sValue
End Property
Property Get DctoSunat() As String
   DctoSunat = ms_DctoSunat
End Property

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: Sunat_ResumenDiario
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: Sunat_ResumenDiario
'        Autor:   DAniel Barrantes B.
'        Fecha: 28/02/2019 07:04:45 p.m.
'------------------------------------------------------------------------------------
 

