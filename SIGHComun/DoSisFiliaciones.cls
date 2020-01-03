VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DoSisFiliaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: SisFiliaciones
'        Autor: Daniel Enrique Barrantes
'        Fecha: 21/11/2013 10:00:07 a.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ml_idSiasis As Long
Dim ms_Codigo As String
Dim ms_AfiliacionDisa As String
Dim ms_AfiliacionTipoFormato As String
Dim ms_AfiliacionNroFormato As String
Dim ms_AfiliacionNroIntegrante As String
Dim ms_DocumentoTipo As String
Dim ms_CodigoEstablAdscripcion As String
Dim mda_AfiliacionFecha As Date
Dim ms_Paterno As String
Dim ms_Materno As String
Dim ms_Pnombre As String
Dim ms_Onombres As String
Dim ms_Genero As String
Dim mda_Fnacimiento As Date
Dim ms_IdDistritoDomicilio As String
Dim ms_Estado As String
Dim mda_Fbaja As Date
Dim ms_DocumentoNumero As String
Dim ms_MotivoBaja As String
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let idSiasis(iValue As Long)
   ml_idSiasis = iValue
End Property
Property Get idSiasis() As Long
   idSiasis = ml_idSiasis
End Property
Property Let Codigo(sValue As String)
   ms_Codigo = sValue
End Property
Property Get Codigo() As String
   Codigo = ms_Codigo
End Property
Property Let AfiliacionDisa(sValue As String)
   ms_AfiliacionDisa = sValue
End Property
Property Get AfiliacionDisa() As String
   AfiliacionDisa = ms_AfiliacionDisa
End Property
Property Let AfiliacionTipoFormato(sValue As String)
   ms_AfiliacionTipoFormato = sValue
End Property
Property Get AfiliacionTipoFormato() As String
   AfiliacionTipoFormato = ms_AfiliacionTipoFormato
End Property
Property Let AfiliacionNroFormato(sValue As String)
   ms_AfiliacionNroFormato = sValue
End Property
Property Get AfiliacionNroFormato() As String
   AfiliacionNroFormato = ms_AfiliacionNroFormato
End Property
Property Let AfiliacionNroIntegrante(sValue As String)
   ms_AfiliacionNroIntegrante = sValue
End Property
Property Get AfiliacionNroIntegrante() As String
   AfiliacionNroIntegrante = ms_AfiliacionNroIntegrante
End Property
Property Let DocumentoTipo(sValue As String)
   ms_DocumentoTipo = sValue
End Property
Property Get DocumentoTipo() As String
   DocumentoTipo = ms_DocumentoTipo
End Property
Property Let CodigoEstablAdscripcion(sValue As String)
   ms_CodigoEstablAdscripcion = sValue
End Property
Property Get CodigoEstablAdscripcion() As String
   CodigoEstablAdscripcion = ms_CodigoEstablAdscripcion
End Property
Property Let AfiliacionFecha(daValue As Date)
   mda_AfiliacionFecha = daValue
End Property
Property Get AfiliacionFecha() As Date
   AfiliacionFecha = mda_AfiliacionFecha
End Property
Property Let Paterno(sValue As String)
   ms_Paterno = sValue
End Property
Property Get Paterno() As String
   Paterno = ms_Paterno
End Property
Property Let Materno(sValue As String)
   ms_Materno = sValue
End Property
Property Get Materno() As String
   Materno = ms_Materno
End Property
Property Let Pnombre(sValue As String)
   ms_Pnombre = sValue
End Property
Property Get Pnombre() As String
   Pnombre = ms_Pnombre
End Property
Property Let Onombres(sValue As String)
   ms_Onombres = sValue
End Property
Property Get Onombres() As String
   Onombres = ms_Onombres
End Property
Property Let Genero(sValue As String)
   ms_Genero = sValue
End Property
Property Get Genero() As String
   Genero = ms_Genero
End Property
Property Let Fnacimiento(daValue As Date)
   mda_Fnacimiento = daValue
End Property
Property Get Fnacimiento() As Date
   Fnacimiento = mda_Fnacimiento
End Property
Property Let IdDistritoDomicilio(sValue As String)
   ms_IdDistritoDomicilio = sValue
End Property
Property Get IdDistritoDomicilio() As String
   IdDistritoDomicilio = ms_IdDistritoDomicilio
End Property
Property Let Estado(sValue As String)
   ms_Estado = sValue
End Property
Property Get Estado() As String
   Estado = ms_Estado
End Property
Property Let Fbaja(daValue As Date)
   mda_Fbaja = daValue
End Property
Property Get Fbaja() As Date
   Fbaja = mda_Fbaja
End Property
Property Let DocumentoNumero(sValue As String)
   ms_DocumentoNumero = sValue
End Property
Property Get DocumentoNumero() As String
   DocumentoNumero = ms_DocumentoNumero
End Property
Property Let MotivoBaja(sValue As String)
   ms_MotivoBaja = sValue
End Property
Property Get MotivoBaja() As String
   MotivoBaja = ms_MotivoBaja
End Property

'------------------------------------------------------------------------------------
'        Fin de código autogenerado para la clase: SisFiliaciones
'        Autor:   Daniel Enrique Barrantes
'------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------
'        Inicio de código manual para la clase: SisFiliaciones
'        Autor:   DAniel Barrantes B.
'        Fecha: 21/11/2013 10:00:07 a.m.
'------------------------------------------------------------------------------------
 

