VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{22ACD161-99EB-11D2-9BB3-00400561D975}#1.0#0"; "PVCALE~1.OCX"
Object = "{8FFC5771-EE23-11D3-9DC0-00A0CC3A1AD6}#1.0#0"; "PVDAYV~1.OCX"
Begin VB.UserControl ucSIcitas 
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13440
   ScaleHeight     =   8400
   ScaleWidth      =   13440
   Begin VB.Frame fraMedico 
      Height          =   7845
      Left            =   75
      TabIndex        =   4
      Top             =   510
      Width           =   3345
      Begin VB.Frame fraPtoCarga 
         Caption         =   "Actualiza Punto Carga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2670
         Left            =   120
         TabIndex        =   8
         Top             =   2865
         Width           =   3135
         Begin VB.CommandButton btnAceptar 
            Caption         =   "Actualizar"
            DisabledPicture =   "ucSIcitas.ctx":0000
            DownPicture     =   "ucSIcitas.ctx":0460
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   165
            Picture         =   "ucSIcitas.ctx":08D5
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   1800
            Width           =   1365
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   165
            MaxLength       =   99
            TabIndex        =   12
            Top             =   1365
            Width           =   525
         End
         Begin VB.TextBox txtCuposXdia 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   165
            MaxLength       =   99
            TabIndex        =   10
            Top             =   570
            Width           =   525
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "múltiplo de 60"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   3
            Left            =   780
            TabIndex        =   13
            Top             =   1425
            Width           =   1170
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "N° minutos de atención"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   165
            TabIndex        =   11
            Top             =   1110
            Width           =   1950
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "N° Cupos por día"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   165
            TabIndex        =   9
            Top             =   330
            Width           =   1380
         End
      End
      Begin MSDataListLib.DataList lstMedicos 
         Height          =   2370
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Haga click sobre el nombre del médico para seleccionarlo"
         Top             =   435
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   4180
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Puntos de Carga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   6
         Top             =   210
         Width           =   1350
      End
   End
   Begin VB.Frame fraProgramacion 
      Height          =   7845
      Left            =   3480
      TabIndex        =   0
      Top             =   510
      Width           =   9915
      Begin PVDayView.PVDayView Diario 
         Height          =   7785
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Haga click con el botón derecho del mouse para agregar una programación"
         Top             =   195
         Width           =   3945
         _Version        =   65536
         DOYAlignment    =   2
         UseCustomCaption=   -1  'True
         Caption         =   ""
         Appearance      =   1
         BorderStyle     =   1
         Increments      =   0
         SelectMode      =   1
         EnableDayChange =   0   'False
         UseControlPanelSettings=   0   'False
         TimeSeparator   =   ":"
         AMString        =   "AM"
         PMString        =   "PM"
         BusinessHoursBegin=   0
         BusinessHoursEnd=   0
         TopIndex        =   3
         TimeBackColor   =   16577517
         SelectedTimeBackColor=   8388608
         AppointmentsForeColor=   0
         AppointmentsBackColor=   16777215
         AppointmentsBarColor=   16737792
         BeginProperty TimeFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty AppointmentsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseStandardDialogs=   0   'False
         FreeTimeColor   =   16777215
         BusyTimeColor   =   16711680
      End
      Begin PVATLCALENDARLib.PVCalendar Calendario 
         Height          =   6765
         Left            =   4140
         TabIndex        =   2
         ToolTipText     =   "Seleccione uno o mas días y haga click con el boton derecho de mouse para agregar un programación"
         Top             =   225
         Width           =   5595
         _Version        =   524288
         BorderStyle     =   1
         Appearance      =   1
         FirstDay        =   1
         Frame           =   1
         SelectMode      =   2
         DisplayFormat   =   0
         DateOrientation =   0
         CustomTextOrientation=   2
         ImageOrientation=   8
         DOWText0        =   "Domingo"
         DOWText1        =   "Lunes"
         DOWText2        =   "Martes"
         DOWText3        =   "Miercoles"
         DOWText4        =   "Jueves"
         DOWText5        =   "Viernes"
         DOWText6        =   "Sabado"
         MonthText0      =   "Enero"
         MonthText1      =   "Febrero"
         MonthText2      =   "MArzo"
         MonthText3      =   "Abril"
         MonthText4      =   "Mayo"
         MonthText5      =   "Junio"
         MonthText6      =   "Julio"
         MonthText7      =   "Agosto"
         MonthText8      =   "Setiembre"
         MonthText9      =   "Octubre"
         MonthText10     =   "Noviembre"
         MonthText11     =   "Diciembre"
         HeaderBackColor =   15780518
         HeaderForeColor =   0
         DisplayBackColor=   13405544
         DisplayForeColor=   0
         DayBackColor    =   16577517
         DayForeColor    =   0
         SelectedDayForeColor=   16777215
         SelectedDayBackColor=   16737792
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DOWFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DaysFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MultiLineText   =   -1  'True
         EditMode        =   0
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label 
         Caption         =   "Pulsar ENTER para ver pacientes citados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   150
         Width           =   3930
      End
      Begin VB.Label lblTotalHoras 
         Caption         =   "Total de Horas Programadas:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4230
         TabIndex        =   3
         Top             =   7470
         Width           =   5505
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Citas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13530
   End
   Begin VB.Menu mnuProgramacion 
      Caption         =   "mnuProgramacion"
   End
   Begin VB.Menu mnuCalendario 
      Caption         =   "mnuCalendario"
      Begin VB.Menu mnuCalAgregarProgramacion 
         Caption         =   "Agregar Programacion"
      End
      Begin VB.Menu mnuCalEliminarProgSelecionada 
         Caption         =   "Eliminar programaciones seleccionados"
      End
   End
   Begin VB.Menu mnuDiario 
      Caption         =   "mnuDiario"
      Begin VB.Menu mnuDiarioAgregarProgramacion 
         Caption         =   "Agregar Programación"
      End
      Begin VB.Menu mnuDiarioModificarProgramacion 
         Caption         =   "Modificar Programación"
      End
      Begin VB.Menu mnuDiarioConsultarProgramacion 
         Caption         =   "Consultar Programación"
      End
      Begin VB.Menu mnuDiarioEliminarProgramacion 
         Caption         =   "Eliminar Programación"
      End
   End
End
Attribute VB_Name = "ucSIcitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de Programación Médica
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Turnos() As doTurno
Dim mo_ReglasConfiguarcionReslab As New SIGHNegocios.ReglasConfiguarcionReslab
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasLaboratorio As New SIGHNegocios.ReglasLaboratorio
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario As Long
Dim mb_SeHaModificadoProgramacion As Boolean
Dim ms_NombreUltimoMedicoSeleccionado As String
Dim mda_UltimaFechaSeleccionada As Date
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_cmbDepartamento As New sighentidades.ListaDespleglable
Dim mo_cmbEspecialidad As New sighentidades.ListaDespleglable
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lnMaximaHorasProgramadasXmedico As Integer
Dim oRsPuntosCarga As New Recordset
Dim ml_Area  As sghAreasLaboraEmpleado
Property Let Area(lValue As sghAreasLaboraEmpleado)
   ml_Area = lValue
End Property
Property Get Area() As sghAreasLaboraEmpleado
   Area = ml_Area
End Property


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Get DiarioVista() As PVDayView.PVDayView
   Set DiarioVista = UserControl.Diario
End Property
Property Get CalendarioVista() As PVCalendar
   Set CalendarioVista = UserControl.Calendario
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
'Property Let MenuAgregarEnabled(bValue As Boolean)
'   UserControl.mnuDiarioAgregarProgramacion.Enabled = bValue
'   UserControl.mnuCalAgregarProgramacion.Enabled = bValue
'End Property
'Property Let MenuModificarEnabled(bValue As Boolean)
'   UserControl.mnuDiarioModificarProgramacion.Enabled = bValue
'End Property
'Property Let MenuEliminarEnabled(bValue As Boolean)
'   UserControl.mnuDiarioEliminarProgramacion.Enabled = bValue
'   UserControl.mnuCalEliminarProgSelecionada.Enabled = bValue
'End Property
'Property Let MenuConsultarEnabled(bValue As Boolean)
'   UserControl.mnuDiarioConsultarProgramacion.Enabled = bValue
'End Property

Private Sub Calendario_Change(ByVal NewDate As Date)
    
    'Si cambia de mes o año pregunta a guardar los datos
    'If Month(mda_UltimaFechaSeleccionada) <> Month(NewDate) Or Year(mda_UltimaFechaSeleccionada) <> Year(NewDate) Then
        'If mb_SeHaModificadoProgramacion Then
        '    If MsgBox("Ud ha modificado la programación del médico " + Chr(13) + UCase(ms_NombreUltimoMedicoSeleccionado) + Chr(13) + ", si no guarda los cambios se perderán. " + Chr(13) + "¿Desea guardar esos cambios? ", vbExclamation + vbYesNo, "Programación médica") = vbYes Then
        '        GrabarProgramacionDelMes
        '    End If
        '    mb_SeHaModificadoProgramacion = False
        'End If
        LimpiarProgramaciones
        LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(NewDate), Year(NewDate)
    'End If
    
    mda_UltimaFechaSeleccionada = NewDate
    'mgaray COmentado para permitir elegir fechas dispersas
    'Diario.CurrentDate = NewDate
    Diario.Caption = Format(NewDate, "dddd, MMMM dd, yyyy")
    
End Sub

Private Sub Calendario_DateDblClick(ByVal DateClicked As Date)
'    Dim oProgInf As New ProgramacionInfDiaria
'
'    Set oProgInf.Diario = Diario
'    Set oProgInf.Calendario = Calendario
'    oProgInf.Show 1

End Sub

Private Sub Calendario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        
    If Button = 2 Then
        PopupMenu mnuCalendario
    End If

End Sub



Sub RefrescarListaMedicos()
    If oRsPuntosCarga.State = 1 Then
       Set oRsPuntosCarga = Nothing
    End If
    With oRsPuntosCarga
        .Fields.Append "idGrupo", adInteger
        .Fields.Append "NombreGrupo", adVarChar, 50, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
        If ml_Area = sghAreasLaboraEmpleado.sghImageneología Then
            .AddNew
            .Fields!idGrupo = 20
            .Fields!nombreGrupo = "Ecografía General"
            .Update
            .AddNew
            .Fields!idGrupo = 23
            .Fields!nombreGrupo = "Ecografía Obstétrica"
            .Update
            .AddNew
            .Fields!idGrupo = 21
            .Fields!nombreGrupo = "Rayos X"
            .Update
            .AddNew
            .Fields!idGrupo = 22
            .Fields!nombreGrupo = "Tomografía"
            .Update
        ElseIf ml_Area = sghAreasLaboraEmpleado.sghLaboratorio Then
            .AddNew
            .Fields!idGrupo = 3
            .Fields!nombreGrupo = "Anatomía Patológica"
            .Update
            .AddNew
            .Fields!idGrupo = 11
            .Fields!nombreGrupo = "Banco Sangre"
            .Update
            .AddNew
            .Fields!idGrupo = 2
            .Fields!nombreGrupo = "Patología Clínica"
            .Update
        End If
    End With
      
    lstMedicos.BoundColumn = "IdGrupo"
    lstMedicos.ListField = "NombreGrupo"
    Set lstMedicos.RowSource = oRsPuntosCarga
    lstMedicos.Tag = ""

End Sub



Private Sub Diario_KeyPress(ByVal KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       ucLaborPacienConCitas1.Visible = True
'       ucLaborPacienConCitas1.LlenaCitasPorFecha mda_UltimaFechaSeleccionada
'    End If
End Sub

Private Sub Diario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        PopupMenu mnuDiario
    End If
End Sub

Private Sub lstMedicos_Click()
    
    'Verifica que no sea el mismo medico
    If lstMedicos.Tag = lstMedicos.BoundText Then
        Exit Sub
    End If
    
    
    LimpiarProgramaciones
    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    
    lstMedicos.Tag = lstMedicos.BoundText
    ms_NombreUltimoMedicoSeleccionado = lstMedicos.Text
End Sub

Private Sub mnuCalAgregarProgramacion_Click()
'Dim oProgDetalle As New LaboratorioProgDetalle
'
'    If lstMedicos.BoundText = "" Then
'        MsgBox "Seleccione un GRUPO EXAMEN", vbInformation, "Programación médica"
'        Exit Sub
'    End If
'
'
'    oProgDetalle.FechaInicial = mda_UltimaFechaSeleccionada
'    oProgDetalle.IdProgramacion = 0
'    oProgDetalle.idGrupo = Val(lstMedicos.BoundText)
'    oProgDetalle.idUsuario = Me.idUsuario
'    oProgDetalle.Opcion = sghAgregar
'    oProgDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
'    oProgDetalle.lcNombrePc = mo_lcNombrePc
'    oProgDetalle.Show 1
'
'
'    LimpiarProgramaciones
'    'Dim lnIdGrupo9 As Long
'    'lnIdGrupo9 = oProgDetalle.idGrupo
'    LeerProgramacionDelMes oProgDetalle.idGrupo, Month(Diario.CurrentDate), Year(Diario.CurrentDate)
'    Unload oProgDetalle
    
End Sub




Public Sub mnuDiarioConsultarProgramacion_Click()
'Dim oProgDetalle As New LaboratorioProgDetalle
'
'    oProgDetalle.FechaInicial = mda_UltimaFechaSeleccionada
'    oProgDetalle.IdProgramacion = 0
'    oProgDetalle.idGrupo = Val(lstMedicos.BoundText)
'    oProgDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
'    oProgDetalle.idUsuario = Me.idUsuario
'    oProgDetalle.Opcion = sghConsultar
'    oProgDetalle.Show 1
'
'    LimpiarProgramaciones
'    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
'    Unload oProgDetalle
    
End Sub

'Private Sub ucLaborPacienConCitas1_SePulsoClicEnSalir(KeyCode As Boolean)
'    If KeyCode = True Then
'       ucLaborPacienConCitas1.Visible = False
'    End If
'End Sub

Private Sub UserControl_Initialize()
    
    'Calendario.AttachDayView Diario
    'Diario.AttachCalendar Calendario
    
End Sub

Public Function Inicializar()
'    Calendario.AttachDayView Diario
'    Diario.AttachCalendar Calendario
    
'    ConfigurarMenusProgramacionMedica
 '   lnMaximaHorasProgramadasXmedico = Val(lcBuscaParametro.SeleccionaFilaParametro(309))
    
 '   RefrescarListaMedicos
End Function

Private Sub UserControl_Resize()
   'On Error Resume Next
   
   lblNombre.Width = UserControl.Width
   
'   fraMedico.Height = UserControl.Height - 600
   lstMedicos.Height = fraMedico.Height - 1700
   fraProgramacion.Height = fraMedico.Height
   
   Diario.Height = fraProgramacion.Height - 330
   Calendario.Height = fraProgramacion.Height - 630
   
'   fraProgramacion.Width = UserControl.Width - fraMedico.Width - 200
   Calendario.Width = fraProgramacion.Width - 4280
   
   lblTotalHoras.Top = Calendario.Top + Calendario.Height + 50
End Sub
Private Sub UserControl_Terminate()
'    Calendario.AttachDayView Nothing
'    Diario.AttachCalendar Nothing
End Sub
Public Sub ConfigurarMenusProgramacionMedica()


    

End Sub



Public Sub mnuDiarioAgregarProgramacion_Click()

'    If mda_UltimaFechaSeleccionada = 0 Then
'       MsgBox "Elija la FECHA", vbInformation, ""
'       Exit Sub
'    End If
'    If Val(lstMedicos.BoundText) = 0 Then
'       MsgBox "Elija el GRUPO EXAMEN", vbInformation, ""
'       Exit Sub
'    End If
'    Dim oProgDetalle As New LaboratorioProgDetalle
'    oProgDetalle.FechaInicial = mda_UltimaFechaSeleccionada
'    oProgDetalle.IdProgramacion = 0
'    oProgDetalle.idGrupo = Val(lstMedicos.BoundText)
'    oProgDetalle.Opcion = sghAgregar
'    oProgDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
'    oProgDetalle.lcNombrePc = mo_lcNombrePc
'    oProgDetalle.Show 1
'
'    LimpiarProgramaciones
'    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
'    Unload oProgDetalle
    
End Sub

Public Sub mnuDiarioEliminarProgramacion_Click()
'Dim oProgDetalle As New LaboratorioProgDetalle
'
'    oProgDetalle.FechaInicial = mda_UltimaFechaSeleccionada
'    oProgDetalle.IdProgramacion = 0
'    oProgDetalle.idGrupo = Val(lstMedicos.BoundText)
'    oProgDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
'    oProgDetalle.idUsuario = Me.idUsuario
'    oProgDetalle.Opcion = sghEliminar
'    oProgDetalle.Show 1
'
'    LimpiarProgramaciones
'    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
'    Unload oProgDetalle

End Sub

Public Sub mnuDiarioModificarProgramacion_Click()
'Dim oProgDetalle As New LaboratorioProgDetalle
'
'    oProgDetalle.FechaInicial = mda_UltimaFechaSeleccionada
'    oProgDetalle.idGrupo = Val(lstMedicos.BoundText)
'    oProgDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
'    oProgDetalle.idUsuario = Me.idUsuario
'    oProgDetalle.Opcion = sghModificar
'    oProgDetalle.Show 1
'
'
'    LimpiarProgramaciones
'    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
'    Unload oProgDetalle
    
End Sub

Private Sub mnuCalEliminarProgSelecionada_Click()
'Dim daDiaSeleccionado As Date
'Dim programacion As PVAppointment
'Dim sTitulo As String
'Dim sHoras() As String
'Dim iHoraIni As Integer
'Dim iHoraFin As Integer
'Dim bTurnoProgramado As Boolean
'
'
'
'    daDiaSeleccionado = Calendario.Value
'    Do While daDiaSeleccionado <> 0
'        sTitulo = ""
'        Set programacion = Diario.AppointmentSet.Get(daDiaSeleccionado)
'
'        bTurnoProgramado = False
'        Do While Not programacion Is Nothing
'            'Verifica que la programacion sea del mismo dia
'            If Format(programacion.StartDateTime, sighentidades.DevuelveFechaSoloFormato_DMY) = daDiaSeleccionado Then
'                If Not mo_AdminProgramacionMedica.ProgramacionMedicaEliminar(programacion.DataVariant, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") Then
'                    MsgBox mo_AdminProgramacionMedica.MensajeError, vbInformation, "Programación médica"
'                Else
'                    Diario.AppointmentSet.Remove programacion.Key
'                    Calendario.DATEText(daDiaSeleccionado) = ""
'                End If
'            Else
'                Exit Sub
'            End If
'            Set programacion = Diario.AppointmentSet.GetNext(programacion)
'        Loop
'        daDiaSeleccionado = Calendario.NextSelectedDate(daDiaSeleccionado)
'    Loop
'
'    mb_SeHaModificadoProgramacion = True
'
'    LimpiarProgramaciones
'    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
'
End Sub

Sub LeerProgramacionDelMes(lIDMedico As Long, iMes As Integer, iAnio As Integer)
'Dim oProgramaciones As Collection
'Dim oProgramacion As DoLaboratorioProg
'Dim programacion As PVAppointment
'Dim sHoras() As String
'Dim iHoraIni As Double
'Dim iHoraFin As Double
'Dim daFechaIni As Date
'Dim sDescripcion As String
'
'Dim dCantidadHoras As Double
'Dim dCantidadHorasMes As Double
'Dim oRsTmp As New Recordset
'Dim lcNombreServicio As String
'Dim lcSql As String, lcFecha As String
'        lblTotalHoras.Caption = "Total Horas Programadas: "
'        'Obtiene las programaciones del medico del mes correspondiente
'        Set oProgramaciones = mo_ReglasLaboratorio.LaboratorioProgLeerPorMedicoYMes(lIDMedico, iMes, iAnio)
'        If oProgramaciones.Count > 0 Then
'            daFechaIni = oProgramaciones.Item(1).fecha
'            For Each oProgramacion In oProgramaciones
'
'                oProgramacion.IdUsuarioAuditoria = ml_idUsuario
'                'Agrega programacion
'                sHoras = Split(oProgramacion.HoraInicio, ":")
'                iHoraIni = Val(sHoras(0)) + IIf(Val(sHoras(1)) = 59, 60, Val(sHoras(1))) / 60
'
'                sHoras = Split(oProgramacion.HoraFinal, ":")
'                iHoraFin = Val(sHoras(0)) + IIf(Val(sHoras(1)) = 59, 60, Val(sHoras(1))) / 60
'
'                dCantidadHoras = Format(iHoraFin - iHoraIni, "##0.00")
'                dCantidadHorasMes = dCantidadHorasMes + oProgramacion.cuposCe
'                'busca Servicio
'                lcFecha = Calendario.Value
'
'
'                Set programacion = Diario.AppointmentSet.Add(lcNombreServicio + " - " + Str(oProgramacion.IdProgramacion), oProgramacion.fecha + iHoraIni / 24, oProgramacion.fecha + iHoraFin / 24)
'                programacion.DataVariant = oProgramacion
'                programacion.ReadOnly = True
'                sDescripcion = "(" & oProgramacion.cuposCe & ")"
'                If daFechaIni = oProgramacion.fecha Then
'                    'Si hay mas de una programación en la misma fecha, concatena los códigos
'                    'sDescripcion = sDescripcion + IIf(sDescripcion <> "", "/", "") & "(" & dCantidadHoras & ")"
'                Else
'                    'Si es la primera programación en el dia
'                    Calendario.DATEText(daFechaIni) = sDescripcion
'                    'sDescripcion = "(" & oProgramacion.cuposCE & ")"
'                    daFechaIni = oProgramacion.fecha
'                    Calendario.DATEForeColor(daFechaIni) = vbBlack
'                    If oProgramacion.estado <> 1 Then
'                       Calendario.DATEForeColor(daFechaIni) = vbRed
'                    End If
'                End If
'            Next
'            Calendario.DATEText(daFechaIni) = sDescripcion
'            lblTotalHoras.Caption = "Total CUPOS Programadas: " & Trim(Str(dCantidadHorasMes))
'
'
'        End If
'
'        Set oRsTmp = Nothing
End Sub

Sub LimpiarProgramaciones()
Dim programacion As PVAppointment
Dim lKey As Long
        
        Set programacion = Diario.AppointmentSet.GetFirst()
        Do While Not programacion Is Nothing
            Calendario.DATEText(Format(programacion.StartDateTime, sighentidades.DevuelveFechaSoloFormato_DMY)) = ""
            lKey = programacion.Key
            Set programacion = Diario.AppointmentSet.GetNext(programacion)
            Diario.AppointmentSet.Remove lKey
        Loop

End Sub

