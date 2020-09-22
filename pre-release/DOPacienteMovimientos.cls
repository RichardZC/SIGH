VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOPacienteMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'------------------------------------------------------------------------------------
'        Organización: Usaid - HFG
'        Aplicativo: SisGalenPlus v.3
'        Programa: código autogenerado para la clase: PacientesMovimientos
'        Programado por: Daniel Enrique Barrantes
'        Fecha: 27/02/2015 04:42:07 p.m.
'------------------------------------------------------------------------------------


Dim ml_IdUsuarioAuditoria As Long

Dim ml_IdCuentaAtencion As Long
Dim mc_Peso As Currency
Dim mc_Talla As Currency
Dim ml_idDxNutricional As Long
Dim ml_GrafXedadEnMeses As Long
Dim ml_GrafYpercentilTE As Long
Dim ml_GrafYpercentilPT As Long
Dim ml_GrafYpercentilPE As Long
Dim mc_ZetaPT As Currency
Dim mc_ZetaTE As Currency
Dim mc_ZetaPE As Currency
Dim mc_Hemoglobina As Currency
Dim ms_Parasitosis As String

Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let IdCuentaAtencion(iValue As Long)
   ml_IdCuentaAtencion = iValue
End Property
Property Get IdCuentaAtencion() As Long
   IdCuentaAtencion = ml_IdCuentaAtencion
End Property
Property Let Peso(cValue As Currency)
   mc_Peso = cValue
End Property
Property Get Peso() As Currency
   Peso = mc_Peso
End Property
Property Let Talla(cValue As Currency)
   mc_Talla = cValue
End Property
Property Get Talla() As Currency
   Talla = mc_Talla
End Property
Property Let idDxNutricional(iValue As Long)
   ml_idDxNutricional = iValue
End Property
Property Get idDxNutricional() As Long
   idDxNutricional = ml_idDxNutricional
End Property
Property Let GrafXedadEnMeses(iValue As Long)
   ml_GrafXedadEnMeses = iValue
End Property
Property Get GrafXedadEnMeses() As Long
   GrafXedadEnMeses = ml_GrafXedadEnMeses
End Property
Property Let GrafYpercentilTE(iValue As Long)
   ml_GrafYpercentilTE = iValue
End Property
Property Get GrafYpercentilTE() As Long
   GrafYpercentilTE = ml_GrafYpercentilTE
End Property
Property Let GrafYpercentilPT(iValue As Long)
   ml_GrafYpercentilPT = iValue
End Property
Property Get GrafYpercentilPT() As Long
   GrafYpercentilPT = ml_GrafYpercentilPT
End Property
Property Let GrafYpercentilPE(iValue As Long)
   ml_GrafYpercentilPE = iValue
End Property
Property Get GrafYpercentilPE() As Long
   GrafYpercentilPE = ml_GrafYpercentilPE
End Property
Property Let ZetaPT(cValue As Currency)
   mc_ZetaPT = cValue
End Property
Property Get ZetaPT() As Currency
   ZetaPT = mc_ZetaPT
End Property
Property Let ZetaTE(cValue As Currency)
   mc_ZetaTE = cValue
End Property
Property Get ZetaTE() As Currency
   ZetaTE = mc_ZetaTE
End Property
Property Let ZetaPE(cValue As Currency)
   mc_ZetaPE = cValue
End Property
Property Get ZetaPE() As Currency
   ZetaPE = mc_ZetaPE
End Property
Property Let Hemoglobina(cValue As Currency)
   mc_Hemoglobina = cValue
End Property
Property Get Hemoglobina() As Currency
   Hemoglobina = mc_Hemoglobina
End Property
Property Let Parasitosis(sValue As String)
   ms_Parasitosis = sValue
End Property
Property Get Parasitosis() As String
   Parasitosis = ms_Parasitosis
End Property
