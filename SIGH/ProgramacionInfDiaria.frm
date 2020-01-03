VERSION 5.00
Begin VB.Form ProgramacionInfDiaria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informacion "
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   Icon            =   "ProgramacionInfDiaria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ProgramacionInfDiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Programación Información diaria
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Diario As PVDayView.PVDayView
Dim mo_Calendario As PVCalendar

Property Set Diario(oValue As PVDayView.PVDayView)
   Set mo_Diario = oValue
End Property
Property Get Diario() As PVDayView.PVDayView
   Set Diario = mo_Diario
End Property
Property Set Calendario(oValue As PVCalendar)
   Set mo_Calendario = oValue
End Property
Property Get Calendario() As PVCalendar
   Set mo_Calendario = mo_Diario
End Property

Private Sub Form_Load()
Dim daDiaSeleccionado As Date
Dim Calendario As PVCalendar
Dim Diario As PVDayView.PVDayView
Dim programacion As PVAppointment
Dim sTitulo As String
Dim sHoras() As String
Dim iHoraIni As Integer
Dim iHoraFin As Integer
Dim bTurnoProgramado As Boolean
Dim doTurno As doTurno
Dim doProgramacion As DOProgramacionMedica

    Set Calendario = mo_Calendario
    Set Diario = mo_Diario


    daDiaSeleccionado = Calendario.Value
    
    Me.Caption = "Informacion del día: " & Format(daDiaSeleccionado, sighEntidades.DevuelveFechaSoloFormato_DMY)
    
    Do While daDiaSeleccionado <> 0
        sTitulo = ""
        Set programacion = Diario.AppointmentSet.Get(daDiaSeleccionado)
        bTurnoProgramado = False
        
        Do While Not programacion Is Nothing
            Set doProgramacion = programacion.DataVariant
            
            'Verifica que la programacion sea del mismo dia
            If Format(programacion.StartDateTime, sighEntidades.DevuelveFechaSoloFormato_DMY) <> daDiaSeleccionado Then
                Exit Do
            End If
            
            '''''''''''''''''''''''''''
            sTitulo = sTitulo + "Codigo: " + programacion.Description + Chr(13)
            sTitulo = sTitulo + "Descripcion: " + doProgramacion.Descripcion + Chr(13)
            sTitulo = sTitulo + "Hora Inicio: " + doProgramacion.HoraInicio + Chr(13)
            sTitulo = sTitulo + "Hora Fin: " + doProgramacion.HoraFin + Chr(13)
            sTitulo = sTitulo + Chr(13)
            
            Set programacion = Diario.AppointmentSet.GetNext(programacion)
        Loop
        
        daDiaSeleccionado = Calendario.NextSelectedDate(daDiaSeleccionado)
    Loop

End Sub

