VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dofarm_formDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: farm_formdet
'        Autor: Daniel Enrique Barrantes
'        Fecha: 11/08/2017 12:39:34 p.m.
'        Auspiciado por: Usaid/Politicas en Salud
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim ml_IdUsuarioAuditoria As Long
Dim ms_CODIGO_EJE As String
Dim ms_CODIGO_PRE As String
Dim ms_TIPSUM As String
Dim ms_ANNOMES As String
Dim ms_CODIGO_MED As String
Dim ml_SALDO As Long
Dim ml_PRECIO As Double
Dim ml_INGRE As Long
Dim ml_REINGRE As Long
Dim ml_VENTA As Long
Dim ml_SIS As Long
Dim ml_INTERSAN As Long
Dim ml_FAC_PERD As Long
Dim ml_DEFNAC As Long
Dim ml_EXO As Long
Dim ml_SOAT As Long
Dim ml_CREDHOSP As Long
Dim ml_OTR_CONV As Long
Dim ml_DEVOL As Long
Dim ml_VENCIDO As Long
Dim ml_MERMA As Long
Dim ml_DISTRI As Long
Dim ml_TRANSF As Long
Dim ml_VENTAINST As Long
Dim ml_DEV_VEN As Long
Dim ml_DEV_MERMA As Long
Dim ml_OTRAS_SAL As Long
Dim ml_STOCK_FIN As Long
Dim ml_STOCK_FIN1 As Long
Dim ml_REQ As Long
Dim ml_TOTAL As Long
Dim mda_FEC_EXP As Date
Dim ml_DO_SALDO As Long
Dim ml_DO_INGRE As Long
Dim ml_DO_CON As Long
Dim ml_DO_OTR As Long
Dim ml_DO_TOT As Long
Dim ml_DO_STK As Long
Dim mda_DO_FECEXP As Date
Dim mda_Fecha As Date
Dim ms_USUARIO As String
Dim ms_INDIPROC As String
Dim ms_SIT As String
Dim ms_INDISIGA As String
Dim ml_DSTKCERO As Long
Dim ml_MPTOREPO As Long
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let CODIGO_EJE(sValue As String)
   ms_CODIGO_EJE = sValue
End Property
Property Get CODIGO_EJE() As String
   CODIGO_EJE = ms_CODIGO_EJE
End Property
Property Let CODIGO_PRE(sValue As String)
   ms_CODIGO_PRE = sValue
End Property
Property Get CODIGO_PRE() As String
   CODIGO_PRE = ms_CODIGO_PRE
End Property
Property Let TIPSUM(sValue As String)
   ms_TIPSUM = sValue
End Property
Property Get TIPSUM() As String
   TIPSUM = ms_TIPSUM
End Property
Property Let ANNOMES(sValue As String)
   ms_ANNOMES = sValue
End Property
Property Get ANNOMES() As String
   ANNOMES = ms_ANNOMES
End Property
Property Let CODIGO_MED(sValue As String)
   ms_CODIGO_MED = sValue
End Property
Property Get CODIGO_MED() As String
   CODIGO_MED = ms_CODIGO_MED
End Property
Property Let SALDO(iValue As Long)
   ml_SALDO = iValue
End Property
Property Get SALDO() As Long
   SALDO = ml_SALDO
End Property
Property Let precio(iValue As Double)
   ml_PRECIO = iValue
End Property
Property Get precio() As Double
   precio = ml_PRECIO
End Property
Property Let INGRE(iValue As Long)
   ml_INGRE = iValue
End Property
Property Get INGRE() As Long
   INGRE = ml_INGRE
End Property
Property Let REINGRE(iValue As Long)
   ml_REINGRE = iValue
End Property
Property Get REINGRE() As Long
   REINGRE = ml_REINGRE
End Property
Property Let VENTA(iValue As Long)
   ml_VENTA = iValue
End Property
Property Get VENTA() As Long
   VENTA = ml_VENTA
End Property
Property Let SIS(iValue As Long)
   ml_SIS = iValue
End Property
Property Get SIS() As Long
   SIS = ml_SIS
End Property
Property Let INTERSAN(iValue As Long)
   ml_INTERSAN = iValue
End Property
Property Get INTERSAN() As Long
   INTERSAN = ml_INTERSAN
End Property
Property Let FAC_PERD(iValue As Long)
   ml_FAC_PERD = iValue
End Property
Property Get FAC_PERD() As Long
   FAC_PERD = ml_FAC_PERD
End Property
Property Let DEFNAC(iValue As Long)
   ml_DEFNAC = iValue
End Property
Property Get DEFNAC() As Long
   DEFNAC = ml_DEFNAC
End Property
Property Let EXO(iValue As Long)
   ml_EXO = iValue
End Property
Property Get EXO() As Long
   EXO = ml_EXO
End Property
Property Let SOAT(iValue As Long)
   ml_SOAT = iValue
End Property
Property Get SOAT() As Long
   SOAT = ml_SOAT
End Property
Property Let CREDHOSP(iValue As Long)
   ml_CREDHOSP = iValue
End Property
Property Get CREDHOSP() As Long
   CREDHOSP = ml_CREDHOSP
End Property
Property Let OTR_CONV(iValue As Long)
   ml_OTR_CONV = iValue
End Property
Property Get OTR_CONV() As Long
   OTR_CONV = ml_OTR_CONV
End Property
Property Let DEVOL(iValue As Long)
   ml_DEVOL = iValue
End Property
Property Get DEVOL() As Long
   DEVOL = ml_DEVOL
End Property
Property Let VENCIDO(iValue As Long)
   ml_VENCIDO = iValue
End Property
Property Get VENCIDO() As Long
   VENCIDO = ml_VENCIDO
End Property
Property Let MERMA(iValue As Long)
   ml_MERMA = iValue
End Property
Property Get MERMA() As Long
   MERMA = ml_MERMA
End Property
Property Let DISTRI(iValue As Long)
   ml_DISTRI = iValue
End Property
Property Get DISTRI() As Long
   DISTRI = ml_DISTRI
End Property
Property Let TRANSF(iValue As Long)
   ml_TRANSF = iValue
End Property
Property Get TRANSF() As Long
   TRANSF = ml_TRANSF
End Property
Property Let VENTAINST(iValue As Long)
   ml_VENTAINST = iValue
End Property
Property Get VENTAINST() As Long
   VENTAINST = ml_VENTAINST
End Property
Property Let DEV_VEN(iValue As Long)
   ml_DEV_VEN = iValue
End Property
Property Get DEV_VEN() As Long
   DEV_VEN = ml_DEV_VEN
End Property
Property Let DEV_MERMA(iValue As Long)
   ml_DEV_MERMA = iValue
End Property
Property Get DEV_MERMA() As Long
   DEV_MERMA = ml_DEV_MERMA
End Property
Property Let OTRAS_SAL(iValue As Long)
   ml_OTRAS_SAL = iValue
End Property
Property Get OTRAS_SAL() As Long
   OTRAS_SAL = ml_OTRAS_SAL
End Property
Property Let STOCK_FIN(iValue As Long)
   ml_STOCK_FIN = iValue
End Property
Property Get STOCK_FIN() As Long
   STOCK_FIN = ml_STOCK_FIN
End Property
Property Let STOCK_FIN1(iValue As Long)
   ml_STOCK_FIN1 = iValue
End Property
Property Get STOCK_FIN1() As Long
   STOCK_FIN1 = ml_STOCK_FIN1
End Property
Property Let REQ(iValue As Long)
   ml_REQ = iValue
End Property
Property Get REQ() As Long
   REQ = ml_REQ
End Property
Property Let TOTAL(iValue As Long)
   ml_TOTAL = iValue
End Property
Property Get TOTAL() As Long
   TOTAL = ml_TOTAL
End Property
Property Let FEC_EXP(daValue As Date)
   mda_FEC_EXP = daValue
End Property
Property Get FEC_EXP() As Date
   FEC_EXP = mda_FEC_EXP
End Property
Property Let DO_SALDO(iValue As Long)
   ml_DO_SALDO = iValue
End Property
Property Get DO_SALDO() As Long
   DO_SALDO = ml_DO_SALDO
End Property
Property Let DO_INGRE(iValue As Long)
   ml_DO_INGRE = iValue
End Property
Property Get DO_INGRE() As Long
   DO_INGRE = ml_DO_INGRE
End Property
Property Let DO_CON(iValue As Long)
   ml_DO_CON = iValue
End Property
Property Get DO_CON() As Long
   DO_CON = ml_DO_CON
End Property
Property Let DO_OTR(iValue As Long)
   ml_DO_OTR = iValue
End Property
Property Get DO_OTR() As Long
   DO_OTR = ml_DO_OTR
End Property
Property Let DO_TOT(iValue As Long)
   ml_DO_TOT = iValue
End Property
Property Get DO_TOT() As Long
   DO_TOT = ml_DO_TOT
End Property
Property Let DO_STK(iValue As Long)
   ml_DO_STK = iValue
End Property
Property Get DO_STK() As Long
   DO_STK = ml_DO_STK
End Property
Property Let DO_FECEXP(daValue As Date)
   mda_DO_FECEXP = daValue
End Property
Property Get DO_FECEXP() As Date
   DO_FECEXP = mda_DO_FECEXP
End Property
Property Let Fecha(daValue As Date)
   mda_Fecha = daValue
End Property
Property Get Fecha() As Date
   Fecha = mda_Fecha
End Property
Property Let USUARIO(sValue As String)
   ms_USUARIO = sValue
End Property
Property Get USUARIO() As String
   USUARIO = ms_USUARIO
End Property
Property Let INDIPROC(sValue As String)
   ms_INDIPROC = sValue
End Property
Property Get INDIPROC() As String
   INDIPROC = ms_INDIPROC
End Property
Property Let SIT(sValue As String)
   ms_SIT = sValue
End Property
Property Get SIT() As String
   SIT = ms_SIT
End Property
Property Let INDISIGA(sValue As String)
   ms_INDISIGA = sValue
End Property
Property Get INDISIGA() As String
   INDISIGA = ms_INDISIGA
End Property
Property Let DSTKCERO(iValue As Long)
   ml_DSTKCERO = iValue
End Property
Property Get DSTKCERO() As Long
   DSTKCERO = ml_DSTKCERO
End Property
Property Let MPTOREPO(iValue As Long)
   ml_MPTOREPO = iValue
End Property
Property Get MPTOREPO() As Long
   MPTOREPO = ml_MPTOREPO
End Property
