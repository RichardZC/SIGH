VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{22ACD161-99EB-11D2-9BB3-00400561D975}#1.0#0"; "PVCALE~1.OCX"
Object = "{8FFC5771-EE23-11D3-9DC0-00A0CC3A1AD6}#1.0#0"; "PVDAYV~1.OCX"
Begin VB.UserControl ucHISListaLotes 
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13425
   ScaleHeight     =   8400
   ScaleWidth      =   13425
   Begin VB.Frame fraMedico 
      Height          =   7845
      Left            =   75
      TabIndex        =   4
      Top             =   510
      Width           =   3345
      Begin VB.ComboBox cmbDepartamento 
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
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   3105
      End
      Begin MSDataListLib.DataList lstMedicos 
         Height          =   6570
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Haga click sobre el nombre del médico para seleccionarlo"
         Top             =   1065
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   11589
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
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   150
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Responsables"
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
         Left            =   135
         TabIndex        =   7
         Top             =   825
         Width           =   1245
      End
   End
   Begin VB.Frame fraProgramacion 
      Height          =   7845
      Left            =   3480
      TabIndex        =   0
      Top             =   510
      Width           =   9915
      Begin PVDayView.PVDayView Diario 
         Height          =   7515
         Left            =   135
         TabIndex        =   1
         ToolTipText     =   "Haga click con el botón derecho del mouse para agregar una programación"
         Top             =   210
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
         Left            =   4155
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
      Caption         =   "Programación"
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
      TabIndex        =   9
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
         Caption         =   "Eliminar programaciones seleccionadas"
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
Attribute VB_Name = "ucHISListaLotes"
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
Dim mo_AdminProgramacionMedica As New SIGHNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasImagenes As New SIGHNegocios.ReglasImagenes
Dim lcBuscaParametro As New SIGHDatos.Parametros
Dim ml_idUsuario As Long
Dim mb_SeHaModificadoProgramacion As Boolean
Dim ms_NombreUltimoMedicoSeleccionado As String
Dim mda_UltimaFechaSeleccionada As Date
Dim mo_Teclado As New sighentidades.Teclado
Dim mo_cmbDepartamento As New sighentidades.ListaDespleglable
Public Event SePresionoTeclaEspecial(KeyCode As Integer)
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String
Dim lnMaximaHorasProgramadasXmedico As Integer
Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Get Departamento() As DataCombo
   Set DiarioVista = UserControl.cmbDepartamento
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
Property Let MenuAgregarEnabled(bValue As Boolean)
   UserControl.mnuDiarioAgregarProgramacion.Enabled = bValue
   UserControl.mnuCalAgregarProgramacion.Enabled = bValue
End Property
Property Let MenuModificarEnabled(bValue As Boolean)
   UserControl.mnuDiarioModificarProgramacion.Enabled = bValue
End Property
Property Let MenuEliminarEnabled(bValue As Boolean)
   UserControl.mnuDiarioEliminarProgramacion.Enabled = bValue
   UserControl.mnuCalEliminarProgSelecionada.Enabled = bValue
End Property
Property Let MenuConsultarEnabled(bValue As Boolean)
   UserControl.mnuDiarioConsultarProgramacion.Enabled = bValue
End Property

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
    Dim oProgInf As New ProgramacionInfDiaria

    Set oProgInf.Diario = Diario
    Set oProgInf.Calendario = Calendario
    oProgInf.Show 1

End Sub

Private Sub Calendario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
        
    If Button = 2 Then
        PopupMenu mnuCalendario
    End If

End Sub

Private Sub cmbDepartamento_Click()
       
        LimpiarProgramaciones
        RefrescarListaMedicos

End Sub
Private Sub cmbDepartamento_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbDepartamento
    RaiseEvent SePresionoTeclaEspecial(KeyCode)
End Sub
Private Sub cmbDepartamento_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Sub ConfirmarActualizacionDeDatosModificados()
        
        If mb_SeHaModificadoProgramacion Then
            If MsgBox("Ud ha modificado la programación del médico " + Chr(13) + UCase(ms_NombreUltimoMedicoSeleccionado) + Chr(13) + ", si no guarda los cambios se perderán. " + Chr(13) + "¿Desea guardar esos cambios? ", vbExclamation + vbYesNo, "Programación médica") = vbYes Then
                GrabarProgramacionDelMes
            End If
            mb_SeHaModificadoProgramacion = False
        End If
        LimpiarProgramaciones
        LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)

End Sub
Private Sub cmbDepartamento_LostFocus()
'   If cmbDepartamento.Text <> "" Then
'       mo_cmbDepartamento.BoundText = Val(Split(cmbDepartamento.Text, " = ")(0))
'   End If
End Sub






Sub RefrescarListaMedicos()
    
    
    lstMedicos.BoundColumn = "idEmpleado"
    lstMedicos.ListField = "ApNom"
    Set lstMedicos.RowSource = mo_ReglasFarmacia.EmpleadosDeImagen("dbo.EmpleadosCargos.idCargo =" & mo_ReglasFarmacia.EmpleadosDevuelveIdCargoSegunPuntoCarga(Val(mo_cmbDepartamento.BoundText)))
    'lstMedicos.Tag = ""
    
End Sub


Private Sub Diario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 2 Then
        PopupMenu mnuDiario
    End If
End Sub

Private Sub lstMedicos_Click()
    
    'Verifica que no sea el mismo medico
'    If lstMedicos.Tag = lstMedicos.BoundText Then
'        Exit Sub
'    End If
    
    
    LimpiarProgramaciones
    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    
    lstMedicos.Tag = lstMedicos.BoundText
    ms_NombreUltimoMedicoSeleccionado = lstMedicos.Text
End Sub

Private Sub mnuCalAgregarProgramacion_Click()
Dim oProgDetalle As New SiProgramacionDetalle
    'franklin 2017
    'If Val(mo_cmbDepartamento.BoundText) = 0 Then
    '    MsgBox "Seleccione el departamento", vbInformation, "Programación médica"
    '    Exit Sub
    'End If

    If lstMedicos.BoundText = "" Then
        MsgBox "Seleccione un médico", vbInformation, "Programación médica"
        Exit Sub
    End If
    
    Set oProgDetalle.Diario = Diario
    Set oProgDetalle.Calendario = Calendario
    oProgDetalle.idPuntoCarga = Val(mo_cmbDepartamento.BoundText)
    oProgDetalle.idMedico = Val(lstMedicos.BoundText)
    oProgDetalle.NombreMedico = lstMedicos.Text
    oProgDetalle.idUsuario = Me.idUsuario
    
    oProgDetalle.Opcion = sghAgregar
    oProgDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    oProgDetalle.lcNombrePc = mo_lcNombrePc
    oProgDetalle.Show 1
    
    mb_SeHaModificadoProgramacion = oProgDetalle.SeHaModificadoProgramacion
    
    LimpiarProgramaciones
    LeerProgramacionDelMes oProgDetalle.idMedico, Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    Unload oProgDetalle
    
End Sub

Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
    Case "Leer"
        If lstMedicos.BoundText = "" Then
            'MsgBox "Por favor seleccione un medico", vbInformation, "Programación Médica"
            Exit Sub
        End If
        
        LimpiarProgramaciones
        LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
        fraProgramacion.Enabled = True
        mb_SeHaModificadoProgramacion = False
        
    Case "Grabar"
        'Elimina todas las programaciones del mes y las vuelve a grabar
        GrabarProgramacionDelMes
        'LimpiarProgramaciones
        fraProgramacion.Enabled = True
        mb_SeHaModificadoProgramacion = False
        
    Case "Cancelar"
        LimpiarProgramaciones
        fraProgramacion.Enabled = True
        mb_SeHaModificadoProgramacion = False
    End Select
    
End Sub



Public Sub mnuDiarioConsultarProgramacion_Click()
Dim oProgDetalle As New SiProgramacionDetalle

    Set oProgDetalle.Diario = Diario
    Set oProgDetalle.Calendario = Calendario
    
    oProgDetalle.idMedico = Val(lstMedicos.BoundText)
    oProgDetalle.idPuntoCarga = Val(mo_cmbDepartamento.BoundText)
    oProgDetalle.NombreMedico = lstMedicos.Text
    oProgDetalle.idUsuario = Me.idUsuario
    
    oProgDetalle.Opcion = sghConsultar
    oProgDetalle.Show 1

    mb_SeHaModificadoProgramacion = oProgDetalle.SeHaModificadoProgramacion
    LimpiarProgramaciones
    LeerProgramacionDelMes oProgDetalle.idMedico, Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    Unload oProgDetalle
    
End Sub

Private Sub UserControl_Initialize()
    
    'Calendario.AttachDayView Diario
    'Diario.AttachCalendar Calendario
    Set mo_cmbDepartamento.MiComboBox = cmbDepartamento

    
End Sub



Public Function Inicializar()
   
    
    Calendario.AttachDayView Diario
    Diario.AttachCalendar Calendario
    
    ConfigurarMenusProgramacionMedica
    lnMaximaHorasProgramadasXmedico = Val(lcBuscaParametro.SeleccionaFilaParametro(309))
End Function

Private Sub UserControl_Resize()
   On Error Resume Next
   
   lblNombre.Width = UserControl.Width
   
   fraMedico.Height = UserControl.Height - 600
   lstMedicos.Height = fraMedico.Height - 1700
   fraProgramacion.Height = fraMedico.Height
   
   Diario.Height = fraProgramacion.Height - 330
   Calendario.Height = fraProgramacion.Height - 630
   
   fraProgramacion.Width = UserControl.Width - fraMedico.Width - 200
   Calendario.Width = fraProgramacion.Width - 4280
   
   lblTotalHoras.Top = Calendario.Top + Calendario.Height + 50
End Sub
Private Sub UserControl_Terminate()
    Calendario.AttachDayView Nothing
    Diario.AttachCalendar Nothing
End Sub
Public Sub ConfigurarMenusProgramacionMedica()
    Dim oRsPuntosCarga As New Recordset
    With oRsPuntosCarga
        .Fields.Append "idGrupo", adInteger
        .Fields.Append "NombreGrupo", adVarChar, 50, adFldIsNullable
        .LockType = adLockOptimistic
        .Open
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
    End With



    mo_cmbDepartamento.BoundColumn = "idGrupo"
    mo_cmbDepartamento.ListField = "nombreGrupo"
    Set mo_cmbDepartamento.RowSource = oRsPuntosCarga
    mo_cmbDepartamento.BoundText = "20"
    
    
    cmbDepartamento_Click

End Sub



Public Sub mnuDiarioAgregarProgramacion_Click()
Dim oProgDetalle As New SiProgramacionDetalle
'franklin 2017
'    If lstMedicos.BoundText = "" Then
'        MsgBox "Seleccione un médico", vbInformation, "Programación médica"
'        Exit Sub
'    End If
    
    Set oProgDetalle.Diario = Diario
    Set oProgDetalle.Calendario = Calendario
    oProgDetalle.idPuntoCarga = Val(mo_cmbDepartamento.BoundText)
    oProgDetalle.idMedico = Val(lstMedicos.BoundText)
    oProgDetalle.NombreMedico = lstMedicos.Text
    oProgDetalle.idUsuario = Me.idUsuario
    
    oProgDetalle.Opcion = sghAgregar
    oProgDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    oProgDetalle.lcNombrePc = mo_lcNombrePc
    oProgDetalle.Show 1
    mb_SeHaModificadoProgramacion = oProgDetalle.SeHaModificadoProgramacion
    
    LimpiarProgramaciones
    LeerProgramacionDelMes oProgDetalle.idMedico, Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    Unload oProgDetalle
    
End Sub

Public Sub mnuDiarioEliminarProgramacion_Click()
Dim oProgDetalle As New SiProgramacionDetalle

    Set oProgDetalle.Diario = Diario
    Set oProgDetalle.Calendario = Calendario
    oProgDetalle.idPuntoCarga = Val(mo_cmbDepartamento.BoundText)
    oProgDetalle.idMedico = Val(lstMedicos.BoundText)
    oProgDetalle.NombreMedico = lstMedicos.Text
    oProgDetalle.idUsuario = Me.idUsuario
    
    oProgDetalle.Opcion = sghEliminar
    oProgDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    oProgDetalle.lcNombrePc = mo_lcNombrePc
    oProgDetalle.Show 1

    LimpiarProgramaciones
    mb_SeHaModificadoProgramacion = oProgDetalle.SeHaModificadoProgramacion
    LeerProgramacionDelMes oProgDetalle.idMedico, Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    Unload oProgDetalle

End Sub

Public Sub mnuDiarioModificarProgramacion_Click()
Dim oProgDetalle As New SiProgramacionDetalle

    Set oProgDetalle.Diario = Diario
    Set oProgDetalle.Calendario = Calendario
    oProgDetalle.idPuntoCarga = Val(mo_cmbDepartamento.BoundText)
    oProgDetalle.idMedico = Val(lstMedicos.BoundText)
    oProgDetalle.NombreMedico = lstMedicos.Text
    oProgDetalle.idUsuario = Me.idUsuario
    
    oProgDetalle.Opcion = sghModificar
    oProgDetalle.lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
    oProgDetalle.lcNombrePc = mo_lcNombrePc
    oProgDetalle.Show 1

    mb_SeHaModificadoProgramacion = oProgDetalle.SeHaModificadoProgramacion
    LimpiarProgramaciones
    LeerProgramacionDelMes oProgDetalle.idMedico, Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    Unload oProgDetalle
    
End Sub

Private Sub mnuCalEliminarProgSelecionada_Click()
Dim daDiaSeleccionado As Date
Dim programacion As PVAppointment
Dim sTitulo As String
Dim sHoras() As String
Dim iHoraIni As Integer
Dim iHoraFin As Integer
Dim bTurnoProgramado As Boolean


    
    daDiaSeleccionado = Calendario.Value
    Do While daDiaSeleccionado <> 0
        sTitulo = ""
        Set programacion = Diario.AppointmentSet.Get(daDiaSeleccionado)
        
        bTurnoProgramado = False
        Do While Not programacion Is Nothing
            'Verifica que la programacion sea del mismo dia
            If Format(programacion.StartDateTime, sighentidades.DevuelveFechaSoloFormato_DMY) = daDiaSeleccionado Then
            
                If Not mo_ReglasImagenes.SiProgramacionEliminarVarios(programacion.DataVariant, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") Then
                    MsgBox mo_AdminProgramacionMedica.MensajeError, vbInformation, "Programación médica"
                Else
                    Diario.AppointmentSet.Remove programacion.Key
                    Calendario.DATEText(daDiaSeleccionado) = ""
                End If
            Else
                Exit Sub
            End If
            Set programacion = Diario.AppointmentSet.GetNext(programacion)
        Loop
        daDiaSeleccionado = Calendario.NextSelectedDate(daDiaSeleccionado)
    Loop

    mb_SeHaModificadoProgramacion = True
    
    LimpiarProgramaciones
    LeerProgramacionDelMes Val(lstMedicos.BoundText), Month(Diario.CurrentDate), Year(Diario.CurrentDate)
    
End Sub
Sub GrabarProgramacionDelMes()
Dim programacion As PVAppointment
Dim oProgramaciones As New Collection
Dim doProgramacion As DOProgramacionMedica
Dim daPrimerDiaDelMes As Date
Dim lIdDepartamento As Long
Dim lIDMedico As Long

        If lstMedicos.BoundText = "" Then
            MsgBox "No hay médicos seleccionados", vbExclamation, "Programación médica"
            Exit Sub
        End If
        
        lIDMedico = Val(lstMedicos.Tag)
        lIdDepartamento = mo_AdminProgramacionMedica.MedicosObtenerDepartamento(lIDMedico)
        If mo_AdminProgramacionMedica.MensajeError <> "" Then
            MsgBox mo_AdminProgramacionMedica.MensajeError, vbInformation, "Programación Medica"
            Exit Sub
        End If
        Set programacion = Diario.AppointmentSet.GetFirst()
        Do While Not programacion Is Nothing
            
            Set doProgramacion = New DOProgramacionMedica
            Set doProgramacion = programacion.DataVariant
            doProgramacion.IdDepartamento = lIdDepartamento
            doProgramacion.idMedico = lIDMedico
            oProgramaciones.Add doProgramacion
            Set programacion = Diario.AppointmentSet.GetNext(programacion)
        Loop
    
        If oProgramaciones.Count >= 0 Then
            If mo_AdminProgramacionMedica.ProgramacionMedicaGrabar(oProgramaciones, lIDMedico, Year(Diario.CurrentDate), Month(Diario.CurrentDate), ml_idUsuario) Then
                mb_SeHaModificadoProgramacion = False
                MsgBox "Los datos se guardarón correctamente", vbInformation, "Programación Médica"
            Else
                MsgBox mo_AdminProgramacionMedica.MensajeError, vbInformation, "Programación Médica"
            End If
        End If
        
End Sub

Sub LeerProgramacionDelMes(lIDMedico As Long, iMes As Integer, iAnio As Integer)
Dim oProgramaciones As Collection
Dim oProgramacion As DOSiProgramacion
Dim programacion As PVAppointment
Dim sHoras() As String
Dim iHoraIni As Double
Dim iHoraFin As Double
Dim daFechaIni As Date
Dim sDescripcion As String
Dim oDOTurno As New doTurno
Dim dCantidadHoras As Double
Dim dCantidadHorasMes As Double
Dim oRsTmp As New Recordset
Dim lcNombreServicio As String
Dim lcSql As String, lcFecha As String
        lblTotalHoras.Caption = "Total Horas Programadas: "
        'Obtiene las programaciones del medico del mes correspondiente
        Set oProgramaciones = mo_ReglasImagenes.SiProgramacionMedicaLeerPorResponsableYMes(lIDMedico, iMes, iAnio)
        If oProgramaciones.Count > 0 Then
            daFechaIni = oProgramaciones.Item(1).fecha
            For Each oProgramacion In oProgramaciones
                
                oProgramacion.IdUsuarioAuditoria = ml_idUsuario
                'Agrega programacion
                sHoras = Split(oProgramacion.HoraInicio, ":")
                iHoraIni = Val(sHoras(0)) + IIf(Val(sHoras(1)) = 59, 60, Val(sHoras(1))) / 60
                
                sHoras = Split(oProgramacion.HoraFin, ":")
                iHoraFin = Val(sHoras(0)) + IIf(Val(sHoras(1)) = 59, 60, Val(sHoras(1))) / 60
                
                dCantidadHoras = Format(iHoraFin - iHoraIni, "##0.00")
                dCantidadHorasMes = dCantidadHorasMes + dCantidadHoras
                'busca Servicio
                lcFecha = Calendario.Value
                Set oRsTmp = mo_ReglasImagenes.SiProgramacionXmedicoFechaHOra(Val(lstMedicos.BoundText), lcFecha, oProgramacion.HoraInicio)
                lcNombreServicio = ""
                If oRsTmp.RecordCount > 0 Then
                   lcNombreServicio = IIf(IsNull(oRsTmp.Fields!nombre), "", oRsTmp.Fields!nombre)
                   lcNombreServicio = lcNombreServicio & " (" & IIf(IsNull(oRsTmp.Fields!dturno), "", oRsTmp.Fields!dturno) & ")"
                End If
                oRsTmp.Close
                Set oDOTurno = mo_AdminProgramacionMedica.TurnosSeleccionarPorId(oProgramacion.IdTurno)
                Set programacion = Diario.AppointmentSet.Add(lcNombreServicio, oProgramacion.fecha + iHoraIni / 24, oProgramacion.fecha + iHoraFin / 24)
                programacion.DataVariant = oProgramacion
                programacion.ReadOnly = True
                If daFechaIni = oProgramacion.fecha Then
                    'Si hay mas de una programación en la misma fecha, concatena los códigos
                    sDescripcion = sDescripcion + IIf(sDescripcion <> "", "/", "") + oDOTurno.Codigo & "(" & dCantidadHoras & ")"
                Else
                    'Si es la primera programación en el dia
                    Calendario.DATEText(daFechaIni) = sDescripcion
                    sDescripcion = oDOTurno.Codigo & "(" & dCantidadHoras & ")"
                    daFechaIni = oProgramacion.fecha
                End If
            Next
            Calendario.DATEText(daFechaIni) = sDescripcion
            lblTotalHoras.Caption = "Total Horas Programadas: " & Trim(Str(dCantidadHorasMes))
            
        End If
        
        Set oRsTmp = Nothing
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

