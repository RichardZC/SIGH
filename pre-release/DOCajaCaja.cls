VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOCajaCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'MZD Ini 02/06/2005 [Todo el archivo]
'MZD02
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: DOCajaCaja
'        Autor: William Castro Grijalva
'        Fecha: 03/06/2005 10:24:57
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
Dim ml_Auditoria As Long
Dim ms_Descripcion As String
Dim ms_codigo As String
Dim ml_idCaja As Long
Dim ms_LoginPC As String
Dim ms_ImpresoraDefault As String
Dim ms_Impresora2 As String
Dim ms_SerieImpresoraDefault As String
'ypalomino 01102014
Dim ms_SerieImpresora2 As String
Dim ml_IdTipoComprobante As Long
Dim mb_FormatoImpDefaultCinta As Boolean
Dim mb_FormatoImp2Cinta As Boolean
Dim ml_IdTipoComprobante2 As Long
Dim ml_IdPartida As Long
Dim ml_IdCentroCosto As Long
Property Let IdPartida(lValue As Long)
   ml_IdPartida = lValue
End Property
Property Get IdPartida() As Long
   IdPartida = ml_IdPartida
End Property
Property Let IdCentroCosto(lValue As Long)
   ml_IdCentroCosto = lValue
End Property
Property Get IdCentroCosto() As Long
   IdCentroCosto = ml_IdCentroCosto
End Property


Property Let IdTipoComprobante(lValue As Long)
   ml_IdTipoComprobante = lValue
End Property
Property Get IdTipoComprobante() As Long
   IdTipoComprobante = ml_IdTipoComprobante
End Property
Property Let Impresora2(sValue As String)
   ms_Impresora2 = sValue
End Property
Property Get Impresora2() As String
   Impresora2 = ms_Impresora2
End Property
Property Let ImpresoraDefault(sValue As String)
   ms_ImpresoraDefault = sValue
End Property
Property Get ImpresoraDefault() As String
   ImpresoraDefault = ms_ImpresoraDefault
End Property
Property Let LoginPc(sValue As String)
   ms_LoginPC = sValue
End Property
Property Get LoginPc() As String
   LoginPc = ms_LoginPC
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_Auditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_Auditoria
End Property
Property Let Descripcion(sValue As String)
   ms_Descripcion = sValue
End Property
Property Get Descripcion() As String
   Descripcion = ms_Descripcion
End Property
Property Let codigo(sValue As String)
   ms_codigo = sValue
End Property
Property Get codigo() As String
   codigo = ms_codigo
End Property
Property Let idCaja(lValue As Long)
   ml_idCaja = lValue
End Property
Property Get idCaja() As Long
   idCaja = ml_idCaja
End Property
    '**** Programa: se agrego SerieImpresoraDefault y SerieImpresora2
    '**** Programado por:Eder Yamill Palomino Espinoza
    '**** Fecha: 06102014
Property Let SerieImpresoraDefault(sValue As String)
   ms_SerieImpresoraDefault = sValue
End Property
Property Get SerieImpresoraDefault() As String
   SerieImpresoraDefault = ms_SerieImpresoraDefault
End Property
Property Let SerieImpresora2(sValue As String)
   ms_SerieImpresora2 = sValue
End Property
Property Get SerieImpresora2() As String
   SerieImpresora2 = ms_SerieImpresora2
End Property
Property Let FormatoImpDefaultCinta(bValue As Boolean)
   mb_FormatoImpDefaultCinta = bValue
End Property
Property Get FormatoImpDefaultCinta() As Boolean
   FormatoImpDefaultCinta = mb_FormatoImpDefaultCinta
End Property
Property Let FormatoImp2Cinta(bValue As Boolean)
   mb_FormatoImp2Cinta = bValue
End Property
Property Get FormatoImp2Cinta() As Boolean
   FormatoImp2Cinta = mb_FormatoImp2Cinta
End Property
Property Let IdTipoComprobante2(lValue As Long)
   ml_IdTipoComprobante2 = lValue
End Property
Property Get IdTipoComprobante2() As Long
   IdTipoComprobante2 = ml_IdTipoComprobante2
End Property
