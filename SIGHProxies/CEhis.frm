VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form CEhis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Imprime Formato HIS"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   7185
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   9195
      Begin VB.ComboBox cmbIdResponsable 
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
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   180
         Width           =   5010
      End
      Begin VB.ComboBox cmbIdServicioCE 
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
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   555
         Width           =   5010
      End
      Begin Threed.SSOption optHIS 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   990
         Width           =   3930
         _ExtentX        =   6932
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
         Caption         =   "Formato HIS del MINSA"
         Value           =   -1
      End
      Begin VB.Frame Frame1 
         Height          =   3180
         Left            =   435
         TabIndex        =   4
         Top             =   1335
         Width           =   8670
         Begin SIGHProxies.ucElegirTurno ucElegirTurno1 
            Height          =   1830
            Left            =   6495
            TabIndex        =   34
            Top             =   180
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   3228
         End
         Begin VB.CheckBox chkExportaCPT 
            Caption         =   "Agrega a los CIE los procedimientos (CPT) realizados en el mismo servicio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   17
            Top             =   2085
            Value           =   1  'Checked
            Width           =   6345
         End
         Begin VB.Frame Frame2 
            Height          =   1785
            Left            =   150
            TabIndex        =   7
            Top             =   210
            Width           =   6285
            Begin VB.PictureBox progressRpt 
               Height          =   300
               Left            =   4410
               ScaleHeight     =   240
               ScaleWidth      =   1770
               TabIndex        =   8
               Top             =   135
               Visible         =   0   'False
               Width           =   1830
            End
            Begin Threed.SSOption optFecha 
               Height          =   255
               Left            =   90
               TabIndex        =   9
               Top             =   210
               Width           =   1545
               _ExtentX        =   2725
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
               Caption         =   "Por Fecha"
               Value           =   -1
            End
            Begin MSMask.MaskEdBox txtFechaInicio 
               Height          =   315
               Left            =   2175
               TabIndex        =   10
               Top             =   510
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin Threed.SSOption optMesAnio 
               Height          =   255
               Left            =   90
               TabIndex        =   11
               Top             =   960
               Width           =   1545
               _ExtentX        =   2725
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
               Caption         =   "Por Mes y Año"
            End
            Begin MSMask.MaskEdBox txtMes 
               Height          =   315
               Left            =   2190
               TabIndex        =   12
               Top             =   1350
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtAnio 
               Height          =   315
               Left            =   3540
               TabIndex        =   13
               Top             =   1320
               Width           =   885
               _ExtentX        =   1561
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
            Begin MSMask.MaskEdBox txtHinicio 
               Height          =   315
               Left            =   3975
               TabIndex        =   32
               Top             =   495
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox txtHfinal 
               Height          =   315
               Left            =   4815
               TabIndex        =   33
               Top             =   495
               Width           =   780
               _ExtentX        =   1376
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   9
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##:##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "F.Atención"
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
               Left            =   1260
               TabIndex        =   16
               Top             =   540
               Width           =   885
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
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
               Height          =   210
               Left            =   1830
               TabIndex        =   15
               Top             =   1380
               Width           =   315
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
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
               Height          =   210
               Left            =   3150
               TabIndex        =   14
               Top             =   1350
               Width           =   330
            End
         End
         Begin VB.CheckBox chkAgregaOrdenesMedicas 
            Caption         =   "Agrega a las Ordenes Médicas (Laboratorio/Imágenes)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   150
            TabIndex        =   6
            Top             =   2430
            Width           =   6345
         End
         Begin VB.CheckBox chkFormato2017 
            Caption         =   "Formato 2017 del MINSA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   135
            TabIndex        =   5
            Top             =   2760
            Value           =   1  'Checked
            Width           =   6345
         End
      End
      Begin Threed.SSOption optAtencionNino 
         Height          =   255
         Left            =   135
         TabIndex        =   19
         Top             =   4755
         Width           =   5250
         _ExtentX        =   9260
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
         Caption         =   "Registro diario de atención del niño (cred #)"
      End
      Begin Threed.SSOption optVacunacionYseg 
         Height          =   255
         Left            =   150
         TabIndex        =   20
         Top             =   5655
         Width           =   7515
         _ExtentX        =   13256
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
         Caption         =   "Registro diario de vacunación y seguimiento del niño y de la niña (inmunización #, LAB)"
      End
      Begin Threed.SSOption optSistemaInform 
         Height          =   255
         Left            =   135
         TabIndex        =   21
         Top             =   6495
         Width           =   5955
         _ExtentX        =   10504
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
         Caption         =   "Sistema de Información del estado nutricional (cred #)"
      End
      Begin MSMask.MaskEdBox txtFregistroDiario 
         Height          =   315
         Left            =   1365
         TabIndex        =   22
         Top             =   5055
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   18
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFvacunacion 
         Height          =   315
         Left            =   1410
         TabIndex        =   24
         Top             =   5955
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFsistemaInf 
         Height          =   315
         Left            =   1395
         TabIndex        =   26
         Top             =   6795
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Médico"
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
         Left            =   180
         TabIndex        =   31
         Top             =   225
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Consultorio CE"
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
         Left            =   180
         TabIndex        =   30
         Top             =   615
         Width           =   1170
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "F.Atención"
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
         Left            =   435
         TabIndex        =   27
         Top             =   6840
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "F.Atención"
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
         Left            =   435
         TabIndex        =   25
         Top             =   6000
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "F.Atención"
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
         Left            =   420
         TabIndex        =   23
         Top             =   5100
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   45
      TabIndex        =   0
      Top             =   7350
      Width           =   9195
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CEhis.frx":0000
         DownPicture     =   "CEhis.frx":0460
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
         Left            =   2451
         Picture         =   "CEhis.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CEhis.frx":0D4A
         DownPicture     =   "CEhis.frx":120E
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
         Left            =   4045
         Picture         =   "CEhis.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         Width           =   1365
      End
   End
End
Attribute VB_Name = "CEhis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Imprime Formato HIS
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_AdminArchivoClinico As New SIGHNegocios.ReglasArchivoClinico
Dim mo_cmbIdTipoHistoria As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdResponsable As New SIGHEntidades.ListaDespleglable
Dim mo_cmbIdServicioCE As New SIGHEntidades.ListaDespleglable
Dim ms_ReglasSeguridad As New ReglasDeSeguridad
Dim sMensaje As String
Dim mo_Teclado As New SIGHEntidades.Teclado



Private Sub btnAceptar_Click()

'If wxFranklin = "*" Then Exit Sub
    Dim oRptHistorias As New RptCEhis
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        oRptHistorias.IdResponsable = Val(mo_cmbIdResponsable.BoundText)
        oRptHistorias.IdServicioCE = Val(mo_cmbIdServicioCE.BoundText)
        oRptHistorias.TextoDelFiltro = cmbIdResponsable.Text
        oRptHistorias.TextoDelFiltro1 = cmbIdServicioCE.Text
        If optHIS.Value = True Then
            If cmbIdResponsable.Text = "" And cmbIdServicioCE.Text = "" Then
               Me.chkFormato2017.Value = 0
            End If
            
            Dim oRsTmp123 As New Recordset
            Dim mo_ReglasDeProgMedica As New ReglasDeProgMedica
            'oRptHistorias.IdResponsable = Val(mo_cmbIdResponsable.BoundText)
            'oRptHistorias.IdServicioCE = Val(mo_cmbIdServicioCE.BoundText)
            If optFecha.Value = True Then
                oRptHistorias.FechaInicio = txtFechaInicio.Text
                oRptHistorias.FechaFin = txtFechaInicio.Text
            Else
                oRptHistorias.FechaInicio = "01/" & Me.txtMes.Text & "/" & Me.txtAnio.Text
                oRptHistorias.FechaFin = Right("0" & SIGHEntidades.DevuelveUltimoDiaDelMes(Val(Me.txtMes.Text), Val(Me.txtAnio.Text)), 2) & "/" & Me.txtMes.Text & "/" & Me.txtAnio.Text
            End If
'            oRptHistorias.TextoDelFiltro = cmbIdResponsable.Text
'            oRptHistorias.TextoDelFiltro1 = cmbIdServicioCE.Text
            oRptHistorias.AgregaCPT = IIf(Me.chkExportaCPT.Value = 1, True, False)
            If Me.chkFormato2017.Value = 1 Then
               oRptHistorias.CrearReporte_excel2017 Me.hwnd, IIf(Me.chkAgregaOrdenesMedicas.Value = 1, True, False), Me.txtHinicio.Text, Me.txtHfinal.Text
            Else
               If cmbIdResponsable.Text <> "" And cmbIdServicioCE.Text <> "" Then
                    oRptHistorias.CrearReporte_excel Me.hwnd, IIf(Me.chkAgregaOrdenesMedicas.Value = 1, True, False), True
               Else
                    Set oRsTmp123 = mo_ReglasDeProgMedica.ProgramacionMedicaSeleccionarPorFechas(CDate(txtFechaInicio.Text & " 00:00:01"), CDate(txtFechaInicio.Text & " 23:59:59"))
                    If oRsTmp123.RecordCount > 0 Then
                       oRsTmp123.MoveFirst
                       Do While Not oRsTmp123.EOF
                          mo_cmbIdServicioCE.BoundText = oRsTmp123!idServicio
                          mo_cmbIdResponsable.BoundText = oRsTmp123!idMedico
                          oRptHistorias.TextoDelFiltro = cmbIdResponsable.Text
                          oRptHistorias.TextoDelFiltro1 = cmbIdServicioCE.Text
                          oRptHistorias.IdResponsable = oRsTmp123!idMedico
                          oRptHistorias.IdServicioCE = oRsTmp123!idServicio
                          oRptHistorias.CrearReporte_excel Me.hwnd, IIf(Me.chkAgregaOrdenesMedicas.Value = 1, True, False), False
                          oRsTmp123.MoveNext
                       Loop
                    End If
                    Me.Visible = False
               End If
            End If
            Me.MousePointer = 1
            Set oRptHistorias = Nothing
            Set oRsTmp123 = Nothing
            Set mo_ReglasDeProgMedica = Nothing
        ElseIf optAtencionNino.Value = True Then
            oRptHistorias.RegistroDiarioAtencionNino Me.hwnd, CDate(txtFregistroDiario.Text)
        ElseIf optVacunacionYseg.Value = True Then
            oRptHistorias.RegistroDiarioVacunacionMujeres Me.hwnd, CDate(Me.txtFvacunacion.Text)
        ElseIf optSistemaInform.Value = True Then
            oRptHistorias.RegistroDiarioNutricional Me.hwnd, CDate(Me.txtFsistemaInf.Text)
        End If
    End If
    Me.MousePointer = 1
End Sub
Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    If cmbIdResponsable.Text = "" Then
       sMensaje = "Elija el Médico"
    End If
    If cmbIdServicioCE.Text = "" Then
       sMensaje = "Elija el Consultorio"
    End If
    If optHIS.Value = True Then
        If Me.txtFechaInicio = SIGHEntidades.FECHA_VACIA_DMY Then
            sMensaje = "Ingrese la fecha de atención"
        Else
            If Not SIGHEntidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
                sMensaje = "La fecha de atención, no tiene el formato correcto"
            End If
        End If
        If SIGHEntidades.EsHora(Me.txtHinicio.Text) = False Then
           sMensaje = "La HORA INICIAL no tiene el formato adecuado"
           ValidaDatosObligatorios = False
           MsgBox sMensaje, vbInformation, ""
           Exit Function
        End If
        If SIGHEntidades.EsHora(Me.txtHfinal.Text) = False Then
           sMensaje = "La HORA FINAL no tiene el formato adecuado"
           ValidaDatosObligatorios = False
           MsgBox sMensaje, vbInformation, ""
           Exit Function
        End If
        If CDate("01/01/2000 " & Me.txtHinicio.Text) > CDate("01/01/2000 " & Me.txtHfinal.Text) Then
           sMensaje = "La HORA FINAL debe ser mayor a la HORA INICIAL"
        End If
    
    End If
    If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       ValidaDatosObligatorios = False
    Else
       ValidaDatosObligatorios = True
    End If
End Function


Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub



Private Sub cmbIdResponsable_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbIdResponsable
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
    Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
    Set mo_cmbIdServicioCE.MiComboBox = cmbIdServicioCE
End Sub


Private Sub Form_Load()
       txtHinicio.Text = "00:00"
       txtHfinal.Text = "23:59"
       Me.ucElegirTurno1.Inicializar

       Dim oBuscaMedicos As New SIGHNegocios.ReglasDeProgMedica
       Dim mo_AdminServHosp As New ReglasServiciosHosp
       Dim oRsTmp As New Recordset
       Dim lnIdEmpleado As Long, lnIdMedico As Long
       Dim oConexion As New Connection
       
       oConexion.CommandTimeout = 300
       oConexion.CursorLocation = adUseClient
       oConexion.Open SIGHEntidades.CadenaConexion
       
       mo_cmbIdResponsable.BoundColumn = "IdMedico"
       mo_cmbIdResponsable.ListField = "Dmedico"
       Set mo_cmbIdResponsable.RowSource = oBuscaMedicos.MedicosSeleccionarTodosOrdenadoAlfabeticamente
       
       Me.txtFechaInicio.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
       Me.txtAnio.Text = Year(Date)
       Me.txtMes.Text = Right("0" & Trim(Str(Month(Date))), 2)
       
       mo_cmbIdServicioCE.BoundColumn = "idServicio"
       mo_cmbIdServicioCE.ListField = "descripcionLarga"
       Set mo_cmbIdServicioCE.RowSource = mo_AdminServHosp.ServiciosSeleccionarPorTipoV2(1, sghFiltraAnuladosYactivos)
       '
       lnIdEmpleado = Val(SIGHEntidades.Usuario)
       If ms_ReglasSeguridad.UsuarioEsMedico(lnIdEmpleado) = True Then
          Set oRsTmp = oBuscaMedicos.MedicosXidEmpleado(lnIdEmpleado, oConexion)
          If oRsTmp.RecordCount > 0 Then
                lnIdMedico = oRsTmp.Fields!idMedico
                mo_cmbIdResponsable.BoundText = Trim(Str(lnIdMedico))
                oRsTmp.Close
                Set oRsTmp = oBuscaMedicos.ProgramacionMedicaSeleccionarPorMedicoFechaServicio(lnIdMedico, txtFechaInicio.Text, 0)
                If oRsTmp.RecordCount > 0 Then
                   mo_cmbIdServicioCE.BoundText = Trim(Str(oRsTmp.Fields!idServicio))
                End If
                oRsTmp.Close
          End If
       End If
       oConexion.Close
       Set oBuscaMedicos = Nothing
       Set mo_AdminServHosp = Nothing
       Set oRsTmp = Nothing
       Set oConexion = Nothing
       
       Me.txtFregistroDiario.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
       Me.txtFvacunacion.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
       Me.txtFsistemaInf.Text = Format(Date, SIGHEntidades.DevuelveFechaSoloFormato_DMY)
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> SIGHEntidades.FECHA_VACIA_DMY Then
        If Not SIGHEntidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = SIGHEntidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub ucElegirTurno1_SeModificoTurnos(lnTurno As SIGHEntidades.sghTurnos, lcHrMinicio As String, lcHrMfinal As String, lcHrTinicio As String, lcHrTfinal As String)
    Select Case lnTurno
    Case sghTurnoTarde
        txtHinicio.Text = lcHrTinicio
        txtHfinal.Text = lcHrTfinal
    Case sghTurnoManana
        txtHinicio.Text = lcHrMinicio
        txtHfinal.Text = lcHrMfinal
    Case sghTurnoAmbos
        txtHinicio.Text = "00:00"
        txtHfinal.Text = "23:59"
    End Select
End Sub
