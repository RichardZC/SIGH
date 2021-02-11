VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoFarmMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de estructura de la tabla farmMovimiento
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit


Dim ml_IdUsuarioAuditoria As Long

Dim ms_movNumero As String
Dim ms_movTipo As String
Dim ml_idAlmacenOrigen As Long
Dim ml_idAlmacenDestino As Long
Dim ml_idTipoConcepto As Long
Dim ml_DocumentoIdtipo As Long
Dim ms_DocumentoNumero As String
Dim ms_Observaciones As String
Dim mv_Total As Double
Dim ml_idMotivoAnulacion As Long
Dim mda_fechaAnulacion As Date
Dim ml_idUsuarioAnulacion As Long
Dim mda_FechaCreacion As Date
Dim ml_idUsuario As Long
Dim ml_idEstadoMovimiento As Long

'RHA 12/01/2021 CAMBIO 50 INICIO
Dim ml_docext As String
Property Let docExterno(sValue As String)
   ml_docext = sValue
End Property
Property Get docExterno() As String
   docExterno = ml_docext
End Property
'RHA 12/01/2021 CAMBIO 50 FIN

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property

Property Let movNumero(sValue As String)
   ms_movNumero = sValue
End Property
Property Get movNumero() As String
   movNumero = ms_movNumero
End Property
Property Let movTipo(sValue As String)
   ms_movTipo = sValue
End Property
Property Get movTipo() As String
   movTipo = ms_movTipo
End Property
Property Let idAlmacenOrigen(iValue As Long)
   ml_idAlmacenOrigen = iValue
End Property
Property Get idAlmacenOrigen() As Long
   idAlmacenOrigen = ml_idAlmacenOrigen
End Property
Property Let idAlmacenDestino(iValue As Long)
   ml_idAlmacenDestino = iValue
End Property
Property Get idAlmacenDestino() As Long
   idAlmacenDestino = ml_idAlmacenDestino
End Property
Property Let idTipoConcepto(iValue As Long)
   ml_idTipoConcepto = iValue
End Property
Property Get idTipoConcepto() As Long
   idTipoConcepto = ml_idTipoConcepto
End Property
Property Let DocumentoIdtipo(iValue As Long)
   ml_DocumentoIdtipo = iValue
End Property
Property Get DocumentoIdtipo() As Long
   DocumentoIdtipo = ml_DocumentoIdtipo
End Property
Property Let DocumentoNumero(sValue As String)
   ms_DocumentoNumero = sValue
End Property
Property Get DocumentoNumero() As String
   DocumentoNumero = ms_DocumentoNumero
End Property
Property Let Observaciones(sValue As String)
   ms_Observaciones = sValue
End Property
Property Get Observaciones() As String
   Observaciones = ms_Observaciones
End Property
Property Let TOTAL(vValue As Double)
   mv_Total = vValue
End Property
Property Get TOTAL() As Double
   TOTAL = mv_Total
End Property
Property Let idMotivoAnulacion(iValue As Long)
   ml_idMotivoAnulacion = iValue
End Property
Property Get idMotivoAnulacion() As Long
   idMotivoAnulacion = ml_idMotivoAnulacion
End Property
Property Let fechaAnulacion(daValue As Date)
   mda_fechaAnulacion = daValue
End Property
Property Get fechaAnulacion() As Date
   fechaAnulacion = mda_fechaAnulacion
End Property
Property Let idUsuarioAnulacion(iValue As Long)
   ml_idUsuarioAnulacion = iValue
End Property
Property Get idUsuarioAnulacion() As Long
   idUsuarioAnulacion = ml_idUsuarioAnulacion
End Property
Property Let FechaCreacion(daValue As Date)
   mda_FechaCreacion = daValue
End Property
Property Get FechaCreacion() As Date
   FechaCreacion = mda_FechaCreacion
End Property

Property Let idUsuario(iValue As Long)
   ml_idUsuario = iValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idEstadoMovimiento(iValue As Long)
   ml_idEstadoMovimiento = iValue
End Property
Property Get idEstadoMovimiento() As Long
   idEstadoMovimiento = ml_idEstadoMovimiento
End Property




