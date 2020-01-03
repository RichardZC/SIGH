VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntegracionEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: enumerados para integracion de pacientes con otros sistemas
'        Programado por: Garay M
'        Fecha: Octubre 2014
'
'------------------------------------------------------------------------------------
Option Explicit

Enum enumGenderCarestream
    Femenino = 1
    Masculino = 2
End Enum

Enum enumPatientClassKeyCarestream
    Ambulatorio = 1
    Emergencia = 2
    Hospitalizacion = 3
End Enum
