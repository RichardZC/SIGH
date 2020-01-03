VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{22ACD161-99EB-11D2-9BB3-00400561D975}#1.0#0"; "PVCALE~1.OCX"
Begin VB.Form frmDetalleDia 
   Caption         =   "Días programados en el establecimiento"
   ClientHeight    =   5205
   ClientLeft      =   6975
   ClientTop       =   2580
   ClientWidth     =   5115
   LinkTopic       =   "form1"
   ScaleHeight     =   5205
   ScaleWidth      =   5115
   Begin VB.ComboBox cmbMes 
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
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar (ESC)"
      DisabledPicture =   "frmDetalleDia.frx":0000
      DownPicture     =   "frmDetalleDia.frx":04C4
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2640
      Picture         =   "frmDetalleDia.frx":09B0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4440
      Width           =   1365
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar (F2)"
      DisabledPicture =   "frmDetalleDia.frx":0E9C
      DownPicture     =   "frmDetalleDia.frx":12FC
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1080
      Picture         =   "frmDetalleDia.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1365
   End
   Begin PVATLCALENDARLib.PVCalendar Calendario 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   795
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
   Begin UltraGrid.SSUltraGrid ugvDetalleDias 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   5530
      _Version        =   131072
      GridFlags       =   17040384
      LayoutFlags     =   67108884
      MaxColScrollRegions=   50
      MaxRowScrollRegions=   50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ugvDetalleDias"
   End
   Begin MSMask.MaskEdBox mskfechaAnio 
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   582
      _Version        =   393216
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Caption         =   "Año"
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
      Left            =   1440
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Mes"
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
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00373842&
      Caption         =   "Días programados"
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
      TabIndex        =   3
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmDetalleDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Interfaz grafica de dias programados.
'        Programado por: Cachay F
'        Fecha: Febrero 2014
'
'------------------------------------------------------------------------------------

Option Explicit
Dim mr_ReglasHIS As New SIGHNegocios.ReglasHISGalenos   'Representa la Capa de Negocios del Modulo HIS GalenHos
Dim mo_Apariencia As New SIGHEntidades.GridInfragistic
Dim mo_cmbMes As New SIGHEntidades.ListaDespleglable
Dim ml_IdEstablecimiento As Long
Dim ml_IdServicio As Long
Dim ml_IdDia As Long
Dim ml_IdMedicoResponsable As Long
Dim mb_IdCerrado As Boolean
Dim ml_IdMes As Long
Dim ml_IdAnio As Long
Dim ml_IdTurno As Long
Dim mi_BotonPresionado As sghBotonDetallePresionado
Dim oRcs_DetalleDias As New Recordset
Const ML_DIAPROGRAMADO As Long = &HCC8D68
Const ML_DIANOPROGRAMADO As Long = &HFCF3ED
Dim mo_Formulario As New SIGHEntidades.Formulario

Property Let IdEstablecimiento(lValue As Long)
   ml_IdEstablecimiento = lValue
End Property
Property Get IdEstablecimiento() As Long
   IdEstablecimiento = ml_IdEstablecimiento
End Property

Property Let IdServicio(lValue As Long)
   ml_IdServicio = lValue
End Property
Property Get IdServicio() As Long
   IdServicio = ml_IdServicio
End Property

Property Let IdMedicoResponsable(lValue As Long)
   ml_IdMedicoResponsable = lValue
End Property
Property Get IdMedicoResponsable() As Long
   IdMedicoResponsable = ml_IdMedicoResponsable
End Property

Property Let IdDia(lValue As Long)
   ml_IdDia = lValue
End Property
Property Get IdDia() As Long
   IdDia = ml_IdDia
End Property

Property Let IdMes(lValue As Long)
   ml_IdMes = lValue
End Property
Property Get IdMes() As Long
   IdMes = ml_IdMes
End Property

Property Let IdAnio(lValue As Long)
   ml_IdAnio = lValue
End Property

Property Get IdAnio() As Long
   IdAnio = ml_IdAnio
End Property

Property Let IdTurno(lValue As Long)
   ml_IdTurno = lValue
End Property

Property Get IdTurno() As Long
   IdTurno = ml_IdTurno
End Property

Property Let BotonPresionado(oValue As sghBotonDetallePresionado)
   mi_BotonPresionado = oValue
End Property
Property Get BotonPresionado() As sghBotonDetallePresionado
   BotonPresionado = mi_BotonPresionado
End Property

Private Sub btnAceptar_Click()
    Dim ms_FechaInicial As String
    mi_BotonPresionado = sghAceptar
    ms_FechaInicial = Format(Me.ugvDetalleDias.ActiveRow.Cells("FechaProgramada").Value, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
    ml_IdDia = Day(ms_FechaInicial)
    Visible = False
End Sub


Private Sub Form_Load()
    cargacombos
    RefrescarProgramacionMedico
End Sub

Private Sub btnCancelar_Click()
    mi_BotonPresionado = sghCancelar
    ml_IdDia = 0
    Visible = False
End Sub

Public Sub MostrarFormulario()
    Me.Show 1
End Sub

Sub cargacombos()
    Set mo_cmbMes.MiComboBox = Me.cmbMes
    mo_cmbMes.BoundColumn = "IdMes"
    mo_cmbMes.ListField = "NombreMes"
    Set mo_cmbMes.RowSource = mr_ReglasHIS.ListaMeses
End Sub

Sub RefrescarProgramacionMedico()
    Dim PrimerDiaSeleccionado As Boolean
    mo_Formulario.HabilitarDeshabilitar cmbMes, False
    mo_Formulario.HabilitarDeshabilitar mskfechaAnio, False
    
    mo_cmbMes.BoundText = CInt(ml_IdMes)
    mskfechaAnio.Text = CInt(ml_IdAnio)
    
    If ml_IdMedicoResponsable <> 0 Then
        'Visualiza la programacion en la Grilla
        Dim oRcs_DetalleProgramacionTemp As New ADODB.Recordset
        Set oRcs_DetalleProgramacionTemp = mr_ReglasHIS.ObtenerDatosProgramacionMedica(ml_IdEstablecimiento, ml_IdServicio, ml_IdMedicoResponsable, CInt(ml_IdAnio), CInt(ml_IdMes), CInt(ml_IdTurno))
        'Visualiza la programacion en el Calendario
        If oRcs_DetalleProgramacionTemp.RecordCount <> 0 Then
            oRcs_DetalleProgramacionTemp.MoveFirst
            Set ugvDetalleDias.DataSource = oRcs_DetalleProgramacionTemp
            mo_Apariencia.ConfigurarFilasBiColores Me.ugvDetalleDias, SIGHEntidades.GrillaConFilasBicolor
        End If
    End If
End Sub

Private Sub ugvDetalleDias_DblClick()
    btnAceptar_Click 'Actualizado 01102014
End Sub

Private Sub ugvDetalleDias_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    With Me.ugvDetalleDias.Bands(0)
        .Columns("IdHisProgMedEstMR").Hidden = True
        .Columns("IdMedico").Hidden = True
        .Columns("IdServicio").Hidden = True
        .Columns("IdEstablecimiento").Hidden = True
        .Columns("Nombre").Hidden = True
        .Columns("FechaProgramada").Header.Caption = "Días programados"
        .Columns("FechaProgramada").Width = 4800
        .Columns("IdTurno").Hidden = True
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub ugvDetalleDias_KeyDown(KeyCode As UltraGrid.SSReturnShort, Shift As Integer)
    If KeyCode = 13 Or KeyCode = vbKeyF3 Then
        btnAceptar_Click
    End If
    If KeyCode = vbKeyEscape Then
        btnCancelar_Click
    End If
End Sub
