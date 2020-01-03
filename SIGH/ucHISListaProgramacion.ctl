VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{22ACD161-99EB-11D2-9BB3-00400561D975}#1.0#0"; "PVCALE~1.OCX"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.UserControl ucHISListaProgramacion 
   BackColor       =   &H80000016&
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14880
   ScaleHeight     =   8490
   ScaleWidth      =   14880
   Begin VB.Frame fraResponsable 
      Caption         =   "Busqueda Responsables"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7845
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   4185
      Begin VB.ComboBox cmbEspecialidad 
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
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   3975
      End
      Begin UltraGrid.SSUltraGrid grdMedicos 
         Height          =   6855
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   12091
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Responsables "
      End
      Begin VB.Label Label2 
         Caption         =   "Especialidad"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraCalendario 
      Caption         =   "Dias Programados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   4200
      TabIndex        =   0
      Top             =   600
      Width           =   10695
      Begin VB.Frame FraBotones 
         Height          =   855
         Left            =   3960
         TabIndex        =   8
         Top             =   6840
         Width           =   6615
         Begin Threed.SSCommand btnAgregarProgramacion 
            Height          =   465
            Left            =   2520
            TabIndex        =   9
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   820
            _Version        =   262144
            PictureFrames   =   1
            Picture         =   "ucHISListaProgramacion.ctx":0000
            Caption         =   "Agregar"
            PictureAlignment=   9
         End
         Begin Threed.SSCommand btnEliminarProgramacion 
            Height          =   465
            Left            =   3960
            TabIndex        =   10
            Top             =   240
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   820
            _Version        =   262144
            PictureFrames   =   1
            Picture         =   "ucHISListaProgramacion.ctx":2F8C
            Caption         =   "Quitar"
            PictureAlignment=   9
            ShapeSize       =   1
         End
      End
      Begin PVATLCALENDARLib.PVCalendar Calendario 
         Height          =   6765
         Left            =   3960
         TabIndex        =   1
         ToolTipText     =   "Seleccione uno o mas días y haga click con el boton derecho de mouse para agregar un programación"
         Top             =   120
         Width           =   6675
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
      Begin UltraGrid.SSUltraGrid grdProgramadoMes 
         Height          =   7455
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   13150
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108884
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Programacion"
      End
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Programación Medica en la MicroRed"
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
      TabIndex        =   6
      Top             =   0
      Width           =   14880
   End
End
Attribute VB_Name = "ucHISListaProgramacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Control para lista de programación ingresada
'        Programado por: Cachay F
'        Fecha: Agosto 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Teclado As New sighEntidades.Teclado
Dim mo_Formulario As New sighEntidades.Formulario
Dim mo_Apariencia As New sighEntidades.GridInfragistic
Dim mo_cmbEspecialidad As New sighEntidades.ListaDespleglable

Dim mo_DatosParametro As New SIGHDatos.Parametros       'Representa la fecha y hora del servidor
Dim mo_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS Galenhos
Dim mo_AdminServiciosHosp As New SIGHNegocios.ReglasServiciosHosp
Dim lcBuscaParametro As New SIGHDatos.Parametros

Dim ms_FechaActual As String
Dim ml_idUsuario As Long                                'Indica el ID del Usuario que esta en session activa.
Dim mi_Opcion As sghOpciones
Dim ml_lnIdTablaLISTBARITEMS As Long
Dim ms_lcNombrePc As String

Dim ml_IdEstablecimiento As Long        'Contiene el ID del Establecimiento de referencia al responsable medico
Dim ml_IdEstablecimientoLocal As Long   'Contiene el ID establecimiento central de la MR
Dim mi_MesProgramacion As Integer
Dim mi_AnioProgramacion As Integer
Dim ms_IdMedicoSeleccionado As String

Const ML_DIAPROGRAMADO As Long = &HCC8D68
Const ML_DIANOPROGRAMADO As Long = &HFCF3ED

'===================================== PROPIEDADES ======================================
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   ml_lnIdTablaLISTBARITEMS = lValue
End Property
Property Get lnIdTablaLISTBARITEMS() As Long
   lnIdTablaLISTBARITEMS = ml_lnIdTablaLISTBARITEMS
End Property
Property Let lcNombrePc(sValue As String)
   ms_lcNombrePc = sValue
End Property
Property Get lcNombrePc() As String
   lcNombrePc = ms_lcNombrePc
End Property
Property Get IdMedicoSelecionado() As Long
   IdMedicoSelecionado = ms_IdMedicoSeleccionado
End Property

'===================================== EVENTOS ======================================
Private Sub btnAgregarProgramacion_Click()
    AgregarProgramacion
End Sub

Public Sub AgregarProgramacion()
Dim orsTemp As Recordset
Set orsTemp = grdMedicos.DataSource

If orsTemp.RecordCount = 0 Then
    Call MsgBox("Seleccione un responsable", vbInformation, "HIS - Programación de responsables")
    Exit Sub
End If

If ms_IdMedicoSeleccionado <> "" Then
    Dim ms_FechaInicial As String
    Dim ms_FechaFinal As String
    
    'Obtiener las fechas
    If Calendario.SelectedDateCount = 1 Then
        ms_FechaInicial = Format(Calendario.Value, sighEntidades.DevuelveFechaSoloFormato_DMY)
        ms_FechaFinal = 0
    ElseIf Calendario.SelectedDateCount > 1 Then
        Dim DiaSeleccionado As Date
        DiaSeleccionado = Calendario.Value
        ms_FechaInicial = Format(Calendario.Value, sighEntidades.DevuelveFechaSoloFormato_DMY)
        Do While DiaSeleccionado <> 0
            ms_FechaFinal = Format(DiaSeleccionado, sighEntidades.DevuelveFechaSoloFormato_DMY)
            DiaSeleccionado = Calendario.NextSelectedDate(DiaSeleccionado)
        Loop
    End If
    If Val(mo_cmbEspecialidad.BoundText) = 0 Then
            Call MsgBox("No ha seleccionado una especialidad", vbExclamation, "HIS - Programación de responsables")
            Exit Sub
    End If
    If CDate(ms_FechaInicial) < CDate(mo_DatosParametro.RetornaFechaServidorSQL) Then
        If DateDiff("d", CDate(ms_FechaInicial), CDate(mo_DatosParametro.RetornaFechaServidorSQL)) > Val(lcBuscaParametro.SeleccionaFilaParametro(330)) Then
            Call MsgBox("No se pueden registrar programaciones hasta " & lcBuscaParametro.SeleccionaFilaParametro(330) & " dias antes de la Fecha Actual.", vbExclamation, "HIS - Programación de responsables")
            Exit Sub
        End If
    End If
    
    Dim ms_resultado As String
    ms_resultado = ValidarDatos(ms_FechaInicial, ms_FechaFinal, ms_IdMedicoSeleccionado)
    If Len(ms_resultado) = 0 Then
        Dim mo_frmDetalleProgramacionMedica As New SIGHhisDigitacion.DetalleProgHIS
        mo_frmDetalleProgramacionMedica.Opcion = sghAgregar
        mo_frmDetalleProgramacionMedica.idMedico = ms_IdMedicoSeleccionado
        mo_frmDetalleProgramacionMedica.IdEspecialidad = Val(mo_cmbEspecialidad.BoundText)
        mo_frmDetalleProgramacionMedica.idUsuario = ml_idUsuario
        mo_frmDetalleProgramacionMedica.DescripcionEspecialidad = cmbEspecialidad.Text
        mo_frmDetalleProgramacionMedica.FechaInicial = ms_FechaInicial
        mo_frmDetalleProgramacionMedica.FechaFinal = ms_FechaFinal
        mo_frmDetalleProgramacionMedica.MostrarFormulario
        
        'Adicionar Programacion al detalle
        If mo_frmDetalleProgramacionMedica.BotonPresionado = sghAceptar Then
            RefrescarProgramacionMedico
            Set mo_frmDetalleProgramacionMedica = Nothing
        End If
    Else
        Call MsgBox("No se ingreso programación, hay fechas ya reservadas " & vbCrLf & ms_resultado, vbInformation, App.Title)
        Set mo_frmDetalleProgramacionMedica = Nothing
        Exit Sub
    End If
Else
    Call MsgBox("Seleccione un responsable", vbInformation, App.Title)
End If
End Sub

Public Sub ModificarProgramacion(Opciones As sghOpciones)
    If ms_IdMedicoSeleccionado <> "" Then
        Dim oTablaHIS_ProgMedEstMR As New DOHIS_ProgMedEstMR
        If Not grdProgramadoMes.ActiveRow Is Nothing Then
            If grdProgramadoMes.ActiveRow.Selected Then
                Dim mo_frmDetalleProgramacionMedica As New SIGHhisDigitacion.DetalleProgHIS
                mo_frmDetalleProgramacionMedica.Opcion = Opciones
                mo_frmDetalleProgramacionMedica.idMedico = ms_IdMedicoSeleccionado
                mo_frmDetalleProgramacionMedica.IdEspecialidad = Val(mo_cmbEspecialidad.BoundText)
                mo_frmDetalleProgramacionMedica.idUsuario = ml_idUsuario
                mo_frmDetalleProgramacionMedica.DescripcionEspecialidad = cmbEspecialidad.Text
                mo_frmDetalleProgramacionMedica.IdHisProgMedEstMR = CInt(grdProgramadoMes.ActiveRow.Cells("IdHisProgMedEstMR").Value)
                mo_frmDetalleProgramacionMedica.MostrarFormulario
                'Adicionar Programacion al detalle
                If mo_frmDetalleProgramacionMedica.BotonPresionado = sghAceptar Then
                    RefrescarProgramacionMedico
                    Set mo_frmDetalleProgramacionMedica = Nothing
                End If
            Else
                Call MsgBox("Debe seleccionar una programación.", vbInformation, UserControl.lblNombre.Caption)
            End If
        Else
            Call MsgBox("Debe seleccionar una programación.", vbInformation, UserControl.lblNombre.Caption)
        End If
    End If
End Sub

Private Sub btnEliminarProgramacion_Click()
    EliminarProgramacion
End Sub

Public Sub EliminarProgramacion()
    Dim oTablaHIS_ProgMedEstMR As New DOHIS_ProgMedEstMR
    If Not grdProgramadoMes.ActiveRow Is Nothing Then
        If grdProgramadoMes.ActiveRow.Selected Then
            If MsgBox("¿Esta seguro de eliminar la programación seleccionada?", vbOKCancel, UserControl.lblNombre.Caption) = vbCancel Then Exit Sub
       
            oTablaHIS_ProgMedEstMR.IdHisProgMedEstMR = CInt(grdProgramadoMes.ActiveRow.Cells("IdHisProgMedEstMR").Value)
            If mo_ReglasHIS.EliminarRegistroProgramacionMedica(oTablaHIS_ProgMedEstMR) Then
                Call MsgBox("Se eliminó correctamente la programacion seleccionada.", vbInformation, UserControl.lblNombre.Caption)
                RefrescarProgramacionMedico
            Else
                Call MsgBox("No se pudo eliminar la programacion seleccionada.", vbInformation, UserControl.lblNombre.Caption)
            End If
        Else
            Call MsgBox("Debe seleccionar una programación.", vbInformation, UserControl.lblNombre.Caption)
        End If
    Else
            Call MsgBox("Debe seleccionar una programación.", vbInformation, UserControl.lblNombre.Caption)
    End If
End Sub

Private Sub cmbEspecialidad_Click()
    Dim orsTemp As New ADODB.Recordset
    If cmbEspecialidad.Text <> "" Then 'Actualizado 01102014
        Set orsTemp = mo_ReglasHIS.ObtenerListaMedicosMR(Val(cmbEspecialidad.ItemData(cmbEspecialidad.ListIndex)))
        Set grdMedicos.DataSource = orsTemp
        grdMedicos.Update
        Set orsTemp = mo_ReglasHIS.ObtenerDatosProgramacionMedica(0, 0, 0, mi_AnioProgramacion, mi_MesProgramacion, 0)
        Set grdProgramadoMes.DataSource = orsTemp
        If orsTemp.RecordCount = 0 Then
            LimpiarProgramaciones
        End If
    End If
End Sub

Private Sub grdMedicos_Click()
    RefrescarProgramacionMedico
End Sub

Private Sub grdMedicos_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    Layout.Override.RowSizingArea = ssRowSizingAreaEntireRow
    Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
    Layout.Override.AllowDelete = ssAllowDeleteNo
    
    With grdMedicos.Bands(0)
        .Columns("IdEmpleado").Hidden = True
        .Columns("IdMedico").Hidden = True
        .Columns("Nombre").Header.Caption = "Responsable"
        .Columns("Nombre").Width = 3600
        .Columns("IdEstablecimientoExterno").Hidden = True
    End With
End Sub

Private Sub UserControl_Resize()
   On Error Resume Next
   
   lblNombre.Width = UserControl.Width
   
   fraResponsable.Height = UserControl.Height - 600
   grdMedicos.Height = fraResponsable.Height - 900
   fraCalendario.Height = fraResponsable.Height
      
   grdProgramadoMes.Height = fraCalendario.Height - 330
   Calendario.Height = fraCalendario.Height - 1130
   
   FraBotones.Top = Calendario.Top + Calendario.Height + 60
   
   fraCalendario.Width = UserControl.Width - fraResponsable.Width - 200
   grdProgramadoMes.Width = 4550
   Calendario.Left = grdProgramadoMes.Left + grdProgramadoMes.Width + 100
   Calendario.Width = fraCalendario.Width - grdProgramadoMes.Width - 200
   
   FraBotones.Left = Calendario.Left
   FraBotones.Width = Calendario.Width
End Sub

Private Sub grdMedicos_AfterRowActivate()
'verificar si se cambia de medico para que se refresque el listado de programaciones
'    RefrescarProgramacionMedico
End Sub

'Configura la grilla de ingreso de programacion
Private Sub grdProgramadoMes_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
Layout.Override.RowSizingArea = ssRowSizingAreaEntireRow
Layout.Override.HeaderClickAction = ssHeaderClickActionSortMulti
Layout.Override.AllowDelete = ssAllowDeleteNo

With grdProgramadoMes.Bands(0)
    'Configuracion de Detalle de Atenciones
    .Columns("IdHisProgMedEstMR").Hidden = True
    
    .Columns("IdMedico").Hidden = True
    .Columns("IdServicio").Hidden = True
    
    .Columns("IdEstablecimiento").Hidden = True
    .Columns("Nombre").Header.Caption = "Establecimiento Destino"
    .Columns("Nombre").Width = 1600
    
    .Columns("Servicio").Header.Caption = "Servicio"
    .Columns("Servicio").Width = 1000
        
    .Columns("FechaProgramada").Header.Caption = "Dia programado"
    .Columns("FechaProgramada").Width = 1100
    .Columns("FechaProgramada").Format = "dd/mm/yyyy"
    
    .Columns("idTurno").Header.Caption = "Turno"
    .Columns("idTurno").Width = 950
    .Columns("idTurno").ValueList = "Turnos"
End With
End Sub

'===================================== METODOS ======================================
Public Function inicializar()
Set mo_cmbEspecialidad.MiComboBox = cmbEspecialidad

CargarComboBoxes

'Consulta la fecha del sistema
ms_FechaActual = mo_DatosParametro.RetornaFechaHoraServidorSQL
Calendario.Value = CDate(ms_FechaActual)

'Cargar el valor del mes y del Año con el control
mi_AnioProgramacion = Year(CDate(ms_FechaActual))
mi_MesProgramacion = Month(CDate(ms_FechaActual))

'Consultar id del Establecimiento Base
Dim oRcs_DatosEstablecimiento As ADODB.Recordset
Set oRcs_DatosEstablecimiento = mo_ReglasHIS.ObtenerDatosEstablecimientoPorUsuario(ml_idUsuario)

If oRcs_DatosEstablecimiento.RecordCount <> 0 Then
oRcs_DatosEstablecimiento.MoveFirst
Do While Not oRcs_DatosEstablecimiento.EOF
    ml_IdEstablecimientoLocal = CLng(oRcs_DatosEstablecimiento!IdEstablecimiento)
    oRcs_DatosEstablecimiento.MoveNext
Loop
End If

mo_Apariencia.ConfigurarFilasBiColores UserControl.grdMedicos, sighEntidades.GrillaConFilasBicolor
mo_Apariencia.ConfigurarFilasBiColores UserControl.grdProgramadoMes, sighEntidades.GrillaConFilasBicolor
End Function

'Carga los listados de servicios y de establecimiento
Sub CargarComboBoxes()
Dim oRcs_Lista As New Recordset

'Codigo del las Especialidades de los Establecimientos Externos
Set oRcs_Lista = mo_ReglasHIS.ListarEspecialidadesEstablecimientosExternos

'verifica si tiene servicios configurados para elos locales de MR
If oRcs_Lista.RecordCount = 0 Then
    MsgBox "No tiene establecimientos ni servicios configurados", vbExclamation, "HIS "
    Exit Sub
End If

oRcs_Lista.MoveFirst
mo_cmbEspecialidad.BoundColumn = "IdEspecialidad"
mo_cmbEspecialidad.ListField = "Nombre"
Set mo_cmbEspecialidad.RowSource = oRcs_Lista
If oRcs_Lista.RecordCount > 0 Then
    oRcs_Lista.MoveFirst
    mo_cmbEspecialidad.BoundText = oRcs_Lista.Fields!IdEspecialidad
End If
    


Set oRcs_Lista = Nothing

'Codigo del tipo de Turno
Set oRcs_Lista = mo_ReglasHIS.ListaTurnos
oRcs_Lista.MoveFirst
If Not grdProgramadoMes.ValueLists.Exists("Turnos") Then
grdProgramadoMes.ValueLists.Add ("Turnos")
    While Not oRcs_Lista.EOF
        With grdProgramadoMes.ValueLists("Turnos")
'            .ValueListItems.Add CInt(oRcs_Lista!IdHisTurno), CStr(oRcs_Lista!IdHisTurno & " - " & oRcs_Lista!Descripcion)
            .ValueListItems.Add CInt(oRcs_Lista!IdHisTurno), CStr(oRcs_Lista!Descripcion)
            .Appearance.Font.Name = "Tahoma"
            .Appearance.Font.Size = 8
        End With
        oRcs_Lista.MoveNext
    Wend
End If
End Sub

Sub RefrescarProgramacionMedico()
Dim orsTemp As Recordset
Set orsTemp = grdMedicos.DataSource

If orsTemp.RecordCount = 0 Then Exit Sub

grdMedicos.Update
If Not IsNull(grdMedicos.ActiveRow.Cells("IdMedico").Value) Then
    'Visualiza la programacion en la Grilla
    Dim oRcs_DetalleProgramacionTemp As New ADODB.Recordset
    'Cambio de fecha dependiendo de la fecha del calendario
    Dim fecha As Date
    fecha = CDate(Calendario.Value)
    mi_AnioProgramacion = Year(fecha)
    mi_MesProgramacion = Month(fecha)
    Set oRcs_DetalleProgramacionTemp = mo_ReglasHIS.ObtenerDatosProgramacionMedica(0, 0, CLng(grdMedicos.ActiveRow.Cells("IdMedico").Value), mi_AnioProgramacion, mi_MesProgramacion, 0)
    ms_IdMedicoSeleccionado = CStr(grdMedicos.ActiveRow.Cells("IdMedico").Value)
    Set grdProgramadoMes.DataSource = oRcs_DetalleProgramacionTemp
    
    'Visualiza la programacion en el Calendario
    LimpiarProgramaciones
    If oRcs_DetalleProgramacionTemp.RecordCount <> 0 Then
    oRcs_DetalleProgramacionTemp.MoveFirst
        Do While Not oRcs_DetalleProgramacionTemp.EOF
            Calendario.DATEText(CDate(oRcs_DetalleProgramacionTemp!FechaProgramada)) = CStr(oRcs_DetalleProgramacionTemp!Nombre)
            Calendario.DATEBackColor(CDate(oRcs_DetalleProgramacionTemp!FechaProgramada)) = ML_DIAPROGRAMADO
            oRcs_DetalleProgramacionTemp.MoveNext
        Loop
    End If
End If
End Sub

'Verifica los dias programados del responsable a programar
Private Function ValidarDatos(ms_FechaInicial As String, ms_FechaFinal As String, ms_IdMedicoSeleccionado As String) As String
Dim oRcs_DiasProgramados As New ADODB.Recordset

'Lee Lista (Dias) Programados en el Sistema
Set oRcs_DiasProgramados = mo_ReglasHIS.ListarProgramacionMedica_FechasMesActual(CLng(ms_IdMedicoSeleccionado), ms_FechaInicial)

If oRcs_DiasProgramados.RecordCount <> 0 Then
    Dim mo_DiasProgramados As New Collection
    Dim mo_DiasParaProgramar As New Collection
    Dim i As Date: Dim y As Integer: Dim ms_DiasDuplicados As String
    
    'Obtinene los Dias Programados para el Mes y Año Escogido
    oRcs_DiasProgramados.MoveFirst
    Do While Not oRcs_DiasProgramados.EOF
        mo_DiasProgramados.Add CDate(oRcs_DiasProgramados!FechaProgramada)
        oRcs_DiasProgramados.MoveNext
    Loop
    
    'Lista los Dias para Programar
    For i = CDate(ms_FechaInicial) To CDate(ms_FechaFinal)
        mo_DiasParaProgramar.Add i
    Next
    
    Dim x As Integer
    'Compara las dos Listas y almacena las coincidencias
    For x = 1 To mo_DiasProgramados.Count
        For y = 1 To mo_DiasParaProgramar.Count
            If CDate(mo_DiasProgramados.Item(x)) = CDate(mo_DiasParaProgramar.Item(y)) Then
                ms_DiasDuplicados = ms_DiasDuplicados & CDate(mo_DiasParaProgramar.Item(y)) & vbCrLf
            End If
        Next
    Next
    ValidarDatos = ms_DiasDuplicados
Else
    ValidarDatos = ""
End If
End Function

Sub LimpiarProgramaciones()
Dim oRcs_DiasDelMes As New ADODB.Recordset
Dim i As Integer: Dim mi_diasMesActual As Integer
Set oRcs_DiasDelMes = Nothing
oRcs_DiasDelMes.CursorType = adOpenStatic
oRcs_DiasDelMes.Fields.Append "FechaProgramada", adVarChar, 20, adFldIsNullable
oRcs_DiasDelMes.Open

'Obtencion del Dato Mes y Año Actuales
Dim fecha As Date
fecha = CDate(Calendario.Value)
mi_diasMesActual = diasdelmes(Year(fecha), Month(fecha))

For i = 1 To mi_diasMesActual
    oRcs_DiasDelMes.AddNew
    oRcs_DiasDelMes.Fields(0) = Format(CDate(i & "/" & Month(fecha) & "/" & Year(fecha)), "dd/mm/yyyy")
    oRcs_DiasDelMes.Update
Next
            
oRcs_DiasDelMes.MoveFirst
Do While Not oRcs_DiasDelMes.EOF
    Calendario.DATEText(CDate(oRcs_DiasDelMes!FechaProgramada)) = ""
    Calendario.DATEBackColor(CDate(oRcs_DiasDelMes!FechaProgramada)) = ML_DIANOPROGRAMADO
    oRcs_DiasDelMes.MoveNext
Loop
End Sub

'Actualizado 30092014
Private Sub Calendario_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
'    If Button = 2 Then
'        PopupMenu mnuCalendario
'    End If
End Sub

