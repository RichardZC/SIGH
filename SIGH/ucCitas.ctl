VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{22ACD161-99EB-11D2-9BB3-00400561D975}#1.0#0"; "PVCalendar9.ocx"
Object = "{8FFC5771-EE23-11D3-9DC0-00A0CC3A1AD6}#1.0#0"; "PVDayView9.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucCitas 
   ClientHeight    =   8610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11970
   ScaleHeight     =   8610
   ScaleWidth      =   11970
   Begin VB.Frame fraMedico 
      Height          =   7125
      Left            =   30
      TabIndex        =   3
      Top             =   1290
      Width           =   3375
      Begin MSDataListLib.DataList lstMedicos 
         Height          =   5310
         Left            =   180
         TabIndex        =   4
         Top             =   1590
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   9366
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo cmbEspecialidad 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   990
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDepartamento 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   390
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   150
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Especialidad"
         Height          =   345
         Left            =   180
         TabIndex        =   8
         Top             =   750
         Width           =   2235
      End
      Begin VB.Label Label3 
         Caption         =   "Medicos"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1350
         Width           =   1245
      End
   End
   Begin VB.Frame fraProgramacion 
      Height          =   7785
      Left            =   3450
      TabIndex        =   0
      Top             =   600
      Width           =   9915
      Begin PVDayView.PVDayView Diario 
         Height          =   7425
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   3945
         _Version        =   65536
         DOYAlignment    =   2
         UseCustomCaption=   -1  'True
         Caption         =   ""
         Appearance      =   0
         BorderStyle     =   1
         Increments      =   2
         SelectMode      =   1
         EnableDayChange =   0   'False
         UseControlPanelSettings=   0   'False
         TimeSeparator   =   ":"
         AMString        =   "AM"
         PMString        =   "PM"
         BusinessHoursBegin=   4.16666666666667E-02
         BusinessHoursEnd=   0.958333333333333
         TopIndex        =   28
         TimeBackColor   =   12648447
         SelectedTimeBackColor=   8388608
         AppointmentsForeColor=   0
         AppointmentsBackColor=   16777215
         AppointmentsBarColor=   16711680
         BeginProperty TimeFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty AppointmentsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Height          =   7425
         Left            =   4170
         TabIndex        =   2
         Top             =   240
         Width           =   5595
         _Version        =   524288
         BorderStyle     =   1
         Appearance      =   0
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
         HeaderBackColor =   13160660
         HeaderForeColor =   0
         DisplayBackColor=   13160660
         DisplayForeColor=   0
         DayBackColor    =   15330541
         DayForeColor    =   0
         SelectedDayForeColor=   16777215
         SelectedDayBackColor=   6956042
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DOWFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty DaysFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar toolbar 
      Height          =   540
      Left            =   30
      TabIndex        =   10
      Top             =   720
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   953
      ButtonWidth     =   1429
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refrescar"
            Key             =   "Refrescar"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00808080&
      Caption         =   "Asignación Citas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   30
      TabIndex        =   11
      Top             =   90
      Width           =   13365
   End
   Begin VB.Menu mnuAsignacionCitas 
      Caption         =   "mnuAsignacionCitas"
   End
   Begin VB.Menu mnuDiario 
      Caption         =   "mnuDiario"
      Begin VB.Menu mnuDiarioAgregarCita 
         Caption         =   "Agregar Cita"
      End
      Begin VB.Menu mnuDiarioModificarCita 
         Caption         =   "Modificar Cita"
      End
      Begin VB.Menu mnuDiarioEliminarCita 
         Caption         =   "Eliminar Cita"
      End
   End
End
Attribute VB_Name = "ucCitas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mo_AdminProgramacionMedica As New SIGHReglasNegocios.ReglasDeProgMedica
Dim mo_AdminServiciosHosp As New SIGHReglasNegocios.ReglasServiciosHosp

Property Get Departamento() As DataCombo
   Set DiarioVista = UserControl.cmbDepartamento
End Property
Property Get Especialidad() As DataCombo
   Set Especialidad = UserControl.cmbEspecialidad
End Property

Property Get DiarioVista() As PVDayView.PVDayView
   Set DiarioVista = UserControl.Diario
End Property
Property Get CalendarioVista() As PVCalendar
   Set CalendarioVista = UserControl.Calendario
End Property

Private Sub Calendario_Change(ByVal NewDate As Date)
    Diario.CurrentDate = NewDate
    Diario.Caption = Format(NewDate, "dddd, MMMM dd, yyyy")
End Sub

Private Sub Calendario_DateDblClick(ByVal DateClicked As Date)
Dim oProgInf As New ProgramacionInfDiaria

    Set oProgInf.Diario = Diario
    Set oProgInf.Calendario = Calendario
    oProgInf.Show 1

End Sub

Private Sub cmbDepartamento_Change()
       
        cmbEspecialidad.BoundColumn = "IdEspecialidad"
        cmbEspecialidad.ListField = "DescripcionLarga"
        On Error Resume Next
        Set cmbEspecialidad.RowSource = mo_AdminServiciosHosp.EspecialidadesSeleccionarporDepartamento(Val(cmbDepartamento.BoundText))
       
        cmbEspecialidad.BoundText = ""
        
        RefrescarListaMedicos

End Sub

Private Sub cmbDepartamento_LostFocus()
   If cmbDepartamento.Text <> "" Then
       cmbDepartamento.BoundText = Val(Split(cmbDepartamento.Text, " = ")(0))
   End If
End Sub

Private Sub cmbEspecialidad_Change()
        RefrescarListaMedicos
End Sub

Sub RefrescarListaMedicos()
    
    lstMedicos.BoundColumn = "IdMedico"
    lstMedicos.ListField = "Nombre"
    Set lstMedicos.RowSource = mo_AdminProgramacionMedica.MedicosFiltrarPorDptosYEspecialidad(Val(cmbDepartamento.BoundText), Val(cmbEspecialidad.BoundText))
    
End Sub

Private Sub cmbEspecialidad_LostFocus()
   If cmbEspecialidad.Text <> "" Then
       cmbEspecialidad.BoundText = Val(Split(cmbEspecialidad.Text, " = ")(0))
   End If
End Sub

Private Sub Diario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 2 Then
        PopupMenu mnuDiario
    End If
End Sub

Private Sub toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
    Case "Programar"
        If lstMedicos.BoundText = "" Then
            MsgBox "Por favor seleccione un medico", vbInformation, "Programación Médica"
            Exit Sub
        End If
        
        'Obtiene las programaciones del medico del mes correspondiente
        
    Case "Grabar"
        'Elimina todas las programaciones del mes y las vuelve a grabar
        
    End Select
    
    Select Case Button.Key
    Case "Programar"
        fraMedico.Enabled = False
        Calendario.EditMode = pcReadOnly
        toolbar.Buttons(1).Enabled = False
        toolbar.Buttons(3).Enabled = True
    Case "Grabar", "Cancelar"
        fraMedico.Enabled = True
        Calendario.EditMode = pcDropDown
        Button.Enabled = True
        toolbar.Buttons(1).Enabled = True
        toolbar.Buttons(3).Enabled = False
    End Select
    
End Sub

Private Sub UserControl_Initialize()
    
    Calendario.AttachDayView Diario
    Diario.AttachCalendar Calendario
    
End Sub
Private Sub UserControl_Resize()
   On Error Resume Next
   lblNombre.Width = UserControl.Width - 100
   
   fraMedico.Height = UserControl.Height - 1300
   lstMedicos.Height = fraMedico.Height - 1700
   fraProgramacion.Height = fraMedico.Height + 650
   Diario.Height = fraProgramacion.Height - 330
   Calendario.Height = fraProgramacion.Height - 330
   fraProgramacion.Width = UserControl.Width - fraMedico.Width - 120
   Calendario.Width = fraProgramacion.Width - 4280
End Sub
Private Sub UserControl_Terminate()
    Calendario.AttachDayView Nothing
    Diario.AttachCalendar Nothing
End Sub
Public Sub ConfigurarAsignacionCitas()
    
    Set UserControl.cmbDepartamento.RowSource = mo_AdminServiciosHosp.DepartamentosSeleccionarTodos
    UserControl.cmbDepartamento.BoundColumn = "IdDepartamento"
    UserControl.cmbDepartamento.ListField = "DescripcionLarga"
    
End Sub

Private Sub mnuDiarioAgregarCita_Click()
Dim oAsignarCita As New CitasDetalle

    Set oAsignarCita.Diario = Diario
    Set oAsignarCita.Calendario = Calendario
    
    oAsignarCita.Opcion = sghAgregar
    oAsignarCita.Show 1
    
End Sub

Private Sub mnuDiarioEliminarCita_Click()
Dim oAsignarCita As New CitasDetalle

    Set oAsignarCita.Diario = Diario
    Set oAsignarCita.Calendario = Calendario
    
    oAsignarCita.Opcion = sghEliminar
    oAsignarCita.Show 1

End Sub

Private Sub mnuDiarioModificarCita_Click()
Dim oAsignarCita As New CitasDetalle

    Set oAsignarCita.Diario = Diario
    Set oAsignarCita.Calendario = Calendario
    
    oAsignarCita.Opcion = sghModificar
    oAsignarCita.Show 1

End Sub



