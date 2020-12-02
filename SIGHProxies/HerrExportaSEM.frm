VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form HerrExportaSEM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exporta datos al Sistema SEM"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10005
   Icon            =   "HerrExportaSEM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   30
      TabIndex        =   5
      Top             =   5580
      Width           =   9900
      Begin VB.CheckBox chkConsideraDxIngresos 
         Caption         =   "Considera Dx INGRESOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2010
         TabIndex        =   22
         Top             =   1155
         Width           =   2385
      End
      Begin VB.TextBox txtClave 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   7830
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   19
         Text            =   "1"
         Top             =   525
         Width           =   1890
      End
      Begin VB.ComboBox cmbMeses 
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
         ItemData        =   "HerrExportaSEM.frx":0CCA
         Left            =   2775
         List            =   "HerrExportaSEM.frx":0CCC
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   255
         Width           =   1635
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   315
         Left            =   2010
         TabIndex        =   17
         Top             =   255
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   556
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
      Begin VB.TextBox txtNCorrelativo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2010
         MaxLength       =   30
         TabIndex        =   10
         Text            =   "1"
         Top             =   570
         Width           =   2355
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   300
         Left            =   6375
         TabIndex        =   6
         Top             =   195
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         ForeColor       =   255
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   8310
         TabIndex        =   7
         Top             =   195
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   255
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Threed.SSOption ssoEmergencias 
         Height          =   255
         Left            =   1995
         TabIndex        =   13
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Emergencia"
      End
      Begin Threed.SSOption ssoHospitalizacion 
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   840
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Hospitalización"
      End
      Begin Threed.SSOption optMedicos 
         Height          =   255
         Left            =   5940
         TabIndex        =   23
         Top             =   870
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   450
         _Version        =   262144
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Médicos"
         Value           =   -1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clave para mostrar USUARIO"
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
         Left            =   5490
         TabIndex        =   21
         Top             =   570
         Width           =   2325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Datos de "
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
         Left            =   120
         TabIndex        =   12
         Top             =   900
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Correlativo N°"
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
         Left            =   120
         TabIndex        =   11
         Top             =   570
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha de Alta Médica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1950
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "hasta"
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
         Left            =   7815
         TabIndex        =   8
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consideraciones:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5430
      Left            =   30
      TabIndex        =   3
      Top             =   90
      Width           =   9885
      Begin VB.ListBox cmbConsideraciones 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   4890
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   9645
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2115
      Left            =   30
      TabIndex        =   1
      Top             =   7140
      Width           =   9945
      Begin VB.CommandButton cmdMarzo2017 
         Caption         =   "Exporta al SEM (versión desde Marzo 2017)"
         DisabledPicture =   "HerrExportaSEM.frx":0CCE
         DownPicture     =   "HerrExportaSEM.frx":112E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   105
         Picture         =   "HerrExportaSEM.frx":15A3
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1125
         Width           =   2415
      End
      Begin VB.CommandButton cmdVerJunio2016 
         Caption         =   "Exporta al SEM (versión desde Junio 2016)"
         DisabledPicture =   "HerrExportaSEM.frx":1A18
         DownPicture     =   "HerrExportaSEM.frx":1E78
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5022
         Picture         =   "HerrExportaSEM.frx":22ED
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   225
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdMayo2011 
         Caption         =   "Exporta al SEM (versión desde Mayo 2011)"
         DisabledPicture =   "HerrExportaSEM.frx":2762
         DownPicture     =   "HerrExportaSEM.frx":2BC2
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2560
         Picture         =   "HerrExportaSEM.frx":3037
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   225
         Width           =   2415
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Exporta al SEM (versión hasta Abril 2011)"
         DisabledPicture =   "HerrExportaSEM.frx":34AC
         DownPicture     =   "HerrExportaSEM.frx":390C
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   98
         Picture         =   "HerrExportaSEM.frx":3D81
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   225
         Width           =   2415
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "HerrExportaSEM.frx":41F6
         DownPicture     =   "HerrExportaSEM.frx":46BA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7485
         Picture         =   "HerrExportaSEM.frx":4BA6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   2385
      End
   End
End
Attribute VB_Name = "HerrExportaSEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Exporta información para el Sistema del MINSA SEM
'        Programado por: Barrantes D
'        Fecha: Enero 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Teclado As New sighentidades.Teclado
Dim mo_Formulario As New sighentidades.Formulario
Dim ml_idUsuario As Long
Dim mo_lcNombrePc  As String
Dim ml_Valor As Integer
Dim ml_Tabla As String
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes

Property Let lcNombrePc(lValue As String)
  mo_lcNombrePc = lValue
End Property

Property Let idUsuario(lIdValue As Long)
  ml_idUsuario = lIdValue
End Property

Private Sub btnAceptar_Click()
'    Dim oProcesos As New Procesos
'    oProcesos.ExportaSEMversion2009
'    Set oProcesos = Nothing
End Sub

Private Sub btnCancelar_Click()
  Unload Me
End Sub

Private Sub cboDatos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub



Private Sub cmbMeses_LostFocus()
    ActualizaFechas txtAnio.Text, cmbMeses.ListIndex
End Sub

Private Sub cmdMarzo2017_Click()
  

If wxFranklin = "*" Then Exit Sub

  If Me.optMedicos = False Then
    If (txtFechaInicio.Text = "" Or Not IsDate(txtFechaInicio.Text)) Or (txtFechaFin.Text = "" Or _
              Not IsDate(txtFechaFin.Text)) Or ssoEmergencias.Value = False And ssoHospitalizacion.Value = False Or _
              txtNCorrelativo.Text = "" Then
      MsgBox "Debe completar los datos requeridos", vbInformation, Me.Caption
      txtFechaInicio.SetFocus
      Exit Sub
    End If
  End If
  
  If MsgBox("Esta seguro de exportar datos?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
       Me.MousePointer = 11
       Frame4.Enabled = False
       cmdMarzo2017.Enabled = False
       Dim oProcesos As New Procesos
       oProcesos.idUsuario = ml_idUsuario
       oProcesos.lcNombrePc = mo_lcNombrePc
       If Me.optMedicos.Value = True Then
          oProcesos.ExportaSEM2017medicos Me.txtFechaInicio.Text, Me.txtFechaFin.Text
       Else
          oProcesos.ExportaSEM2017 Me.txtFechaInicio.Text, Me.txtFechaFin.Text, Me.txtNCorrelativo.Text, ml_Valor, ml_Tabla, _
                                IIf(txtClave.Text = Format(Date, "yyyymmdd"), True, False)
       End If
       Me.MousePointer = 1
       If oProcesos.MensajeError = "" Then
          Me.Visible = False
       End If
       Set oProcesos = Nothing
       Frame4.Enabled = True
       btnAceptar.Enabled = True
       'txtFechaInicio.SetFocus
       Me.MousePointer = 1
       Me.Visible = False
  End If
End Sub

Private Sub cmdMayo2011_Click()
  If (txtFechaInicio.Text = "" Or Not IsDate(txtFechaInicio.Text)) Or (txtFechaFin.Text = "" Or Not IsDate(txtFechaFin.Text)) Or ssoEmergencias.Value = False And ssoHospitalizacion.Value = False Or txtNCorrelativo.Text = "" Then
    MsgBox "Debe completar los datos requeridos", vbInformation, Me.Caption
    txtFechaInicio.SetFocus
    Exit Sub
  End If
  
  If MsgBox("Esta seguro de exportar datos?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
       Me.MousePointer = 11
       Frame4.Enabled = False
       cmdMayo2011.Enabled = False
       Dim oProcesos As New Procesos
       oProcesos.idUsuario = ml_idUsuario
       oProcesos.lcNombrePc = mo_lcNombrePc
       oProcesos.ExportaSEM Me.txtFechaInicio.Text, Me.txtFechaFin.Text, Me.txtNCorrelativo.Text, ml_Valor, ml_Tabla
       Me.MousePointer = 1
       If oProcesos.MensajeError = "" Then
          Me.Visible = False
       End If
       Set oProcesos = Nothing
       Frame4.Enabled = True
       btnAceptar.Enabled = True
       'txtFechaInicio.SetFocus
       Me.MousePointer = 1
       Me.Visible = False
  End If
End Sub

Private Sub cmdVerJunio2016_Click()
  

If wxFranklin = "*" Then Exit Sub

  
  If (txtFechaInicio.Text = "" Or Not IsDate(txtFechaInicio.Text)) Or (txtFechaFin.Text = "" Or Not IsDate(txtFechaFin.Text)) Or ssoEmergencias.Value = False And ssoHospitalizacion.Value = False Or txtNCorrelativo.Text = "" Then
    MsgBox "Debe completar los datos requeridos", vbInformation, Me.Caption
    txtFechaInicio.SetFocus
    Exit Sub
  End If
  
  If MsgBox("Esta seguro de exportar datos?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
       Me.MousePointer = 11
       Frame4.Enabled = False
       cmdVerJunio2016.Enabled = False
       Dim oProcesos As New Procesos
       oProcesos.idUsuario = ml_idUsuario
       oProcesos.lcNombrePc = mo_lcNombrePc
       oProcesos.ExportaSEM2016 Me.txtFechaInicio.Text, Me.txtFechaFin.Text, Me.txtNCorrelativo.Text, ml_Valor, ml_Tabla, _
                                IIf(txtClave.Text = Format(Date, "yyyymmdd"), True, False), _
                                IIf(chkConsideraDxIngresos.Value = 1, True, False)
       Me.MousePointer = 1
       If oProcesos.MensajeError = "" Then
          Me.Visible = False
       End If
       Set oProcesos = Nothing
       Frame4.Enabled = True
       btnAceptar.Enabled = True
       'txtFechaInicio.SetFocus
       Me.MousePointer = 1
       Me.Visible = False
  End If

End Sub

Private Sub Form_Load()
  mo_ReglasComunes.LlenaListBoxConTablaMensajesEnVentana cmbConsideraciones, "HerrExportaSEM"
  cmbConsideraciones.AddItem "******* Junio 2016 *******"
  cmbConsideraciones.AddItem "- Deben ODBC=SEM, tabla libre, apunte ..\galenhos\archivos"
  cmbConsideraciones.AddItem "- Deben existir las tablas vacias: sem_egre.dbf, sem_emer.dbf, sem_med.dbf"
  cmbConsideraciones.AddItem "- Llenar en opción GENERAL->SERVICIOS el dato CODIGO SERVICIO (..SEM) de 6 digitos"
  cmbConsideraciones.AddItem "- Llenar en opción FACT-CONFIG->FUENTES FINANCIAMIENTOS el dato CODIGO SISTEMA SEM"
  '
  txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
  txtFechaFin.Text = Date
  mo_Formulario.LlenaComboConMeses cmbMeses
  txtAnio.Text = Year(Date)
  cmbMeses.ListIndex = Month(Date) - 1
  'mo_Formulario.HabilitarDeshabilitar Me.txtNCorrelativo, False
  ml_Valor = 0
  ml_Tabla = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
        Case vbKeyF6
        Case vbKeyEscape
'           btnCancelar_Click
        Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub ssoEmergencias_Click(Value As Integer)
  ml_Valor = 2
  ml_Tabla = "emergenc"
End Sub

Private Sub ssoHospitalizacion_Click(Value As Integer)
  ml_Valor = 3
  ml_Tabla = "egresos"
End Sub






Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFechaFin_LostFocus()
  If Not IsDate(txtFechaFin.Text) Then
    MsgBox "Fecha Final incorrecta", vbInformation, Me.Caption
    txtFechaFin.Text = Date
    txtFechaFin.SetFocus
  End If
End Sub


Private Sub txtFechaInicio_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFechaInicio_LostFocus()
  If Not IsDate(txtFechaInicio.Text) Then
    MsgBox "Fecha Inicial incorrecta", vbInformation, Me.Caption
    txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual()
    txtFechaInicio.SetFocus
  End If
End Sub


Private Sub txtNCorrelativo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtAnio_LostFocus()
    If Val(txtAnio.Text) > 2000 Then
       ActualizaFechas txtAnio.Text, cmbMeses.ListIndex
    Else
       txtAnio.Text = Year(Date)
    End If
End Sub

Sub ActualizaFechas(lcANIO As String, lnMes As Integer)
    txtFechaInicio.Text = "01/" & Right("0" & Trim(str(lnMes + 1)), 2) & "/" & lcANIO
    txtFechaFin.Text = sighentidades.UltimaFechaDDMMYYDelMesActual1(CDate(txtFechaInicio.Text))
End Sub
