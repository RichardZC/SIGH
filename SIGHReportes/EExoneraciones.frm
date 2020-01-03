VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form EExoneraciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exoneraciones - Hospitalización"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   6555
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EExoneraciones.frx":0000
         DownPicture     =   "EExoneraciones.frx":0460
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
         Left            =   1913
         Picture         =   "EExoneraciones.frx":08D5
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EExoneraciones.frx":0D4A
         DownPicture     =   "EExoneraciones.frx":120E
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
         Left            =   3443
         Picture         =   "EExoneraciones.frx":16FA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4380
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      Begin VB.ComboBox cmbServicioSocial 
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
         Left            =   2760
         TabIndex        =   15
         Top             =   1950
         Width           =   3690
      End
      Begin Threed.SSOption optRepFechasAdministrativas 
         Height          =   345
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   609
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
         Caption         =   "Usando Fechas Alta Administrativas"
      End
      Begin MSMask.MaskEdBox txtFechaInicio 
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         Top             =   690
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
      Begin MSMask.MaskEdBox txtFechaFin 
         Height          =   315
         Left            =   5040
         TabIndex        =   2
         Top             =   690
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
      Begin Threed.SSOption optFechasExoneraciones 
         Height          =   345
         Left            =   240
         TabIndex        =   9
         Top             =   1170
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   609
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
         Caption         =   "Usando Fechas de Exoneraciones"
      End
      Begin MSMask.MaskEdBox txtFexoneracion1 
         Height          =   315
         Left            =   2760
         TabIndex        =   10
         Top             =   1500
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
      Begin MSMask.MaskEdBox txtFexoneracion2 
         Height          =   315
         Left            =   5040
         TabIndex        =   11
         Top             =   1500
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
      Begin Threed.SSOption optFechasBoleta 
         Height          =   345
         Left            =   240
         TabIndex        =   16
         Top             =   3225
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   609
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
         Caption         =   "Usando fechas de Comprobante Pago"
         Value           =   -1
      End
      Begin MSMask.MaskEdBox finicio 
         Height          =   315
         Left            =   2220
         TabIndex        =   17
         Top             =   3675
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
      Begin MSMask.MaskEdBox ffinal 
         Height          =   315
         Left            =   4980
         TabIndex        =   18
         Top             =   3675
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
      Begin VB.Label Label 
         Caption         =   $"EExoneraciones.frx":1BE6
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
         Height          =   630
         Left            =   645
         TabIndex        =   21
         Top             =   2325
         Width           =   5745
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "al"
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
         Left            =   4230
         TabIndex        =   20
         Top             =   3735
         Width           =   120
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "F. comprobante"
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
         Left            =   570
         TabIndex        =   19
         Top             =   3705
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Empleado Servicio Social"
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
         Left            =   660
         TabIndex        =   14
         Top             =   2010
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "F. Exoneracion"
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
         Left            =   660
         TabIndex        =   13
         Top             =   1590
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "al"
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
         Left            =   4770
         TabIndex        =   12
         Top             =   1560
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "al"
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
         Left            =   4770
         TabIndex        =   4
         Top             =   750
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "F.Alta Administrativa"
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
         Left            =   690
         TabIndex        =   3
         Top             =   720
         Width           =   1650
      End
   End
End
Attribute VB_Name = "EExoneraciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Exoneraciones
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_cmbServicioSocial As New sighentidades.ListaDespleglable
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion

Private Sub btnAceptar_Click()
    Me.MousePointer = 11
    Dim oRptExoneraciones As New RptEExoneraciones
    If optRepFechasAdministrativas.Value = True Then
    
        If Me.txtFechaInicio = sighentidades.FECHA_VACIA_DMY Then
            MsgBox "Ingrese la fecha de alta administrativa inicial", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not sighentidades.EsFecha(Me.txtFechaInicio, "DD/MM/AAAA") Then
                MsgBox "La fecha de alta administrativa inicial no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        
        If Me.txtFechaFin = sighentidades.FECHA_VACIA_DMY Then
            MsgBox "Ingrese la fecha de alta administrativa final", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not sighentidades.EsFecha(Me.txtFechaFin, "DD/MM/AAAA") Then
                MsgBox "La fecha de alta administrativa final no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        If CDate(Me.txtFechaInicio.Text) > CDate(Me.txtFechaFin.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
           Exit Sub
        End If
    
        oRptExoneraciones.FechaInicio = txtFechaInicio.Text
        oRptExoneraciones.FechaFin = txtFechaFin.Text
        oRptExoneraciones.TextoDelFiltro = "Filtros:     F.Alta Administrativa: (" & txtFechaInicio.Text & " - " & txtFechaFin.Text & ")"
        oRptExoneraciones.CrearReporte_excel Me.hwnd
        
    ElseIf optFechasBoleta.Value = True Then
        If Me.finicio.Text = sighentidades.FECHA_VACIA_DMY Then
            MsgBox "Ingrese la fecha de comprobante inicial", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not sighentidades.EsFecha(Me.finicio.Text, "DD/MM/AAAA") Then
                MsgBox "La fecha de comprobante inicial no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        
        If Me.ffinal.Text = sighentidades.FECHA_VACIA_DMY Then
            MsgBox "Ingrese la fecha de comprobante final", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not sighentidades.EsFecha(Me.ffinal.Text, "DD/MM/AAAA") Then
                MsgBox "La fecha de comprobante final no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        If CDate(Me.finicio.Text) > CDate(Me.ffinal.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
           Exit Sub
        End If
        
        oRptExoneraciones.FechaInicio = Me.finicio.Text
        oRptExoneraciones.FechaFin = Me.ffinal.Text & " 23:59:59"
        oRptExoneraciones.TextoDelFiltro = "Filtros:     F.Boletas: (" & Me.finicio.Text & " - " & Me.ffinal.Text & ")"

        oRptExoneraciones.CrearReportePorFechasBoletas
    
    Else
    
        If Me.txtFexoneracion1 = sighentidades.FECHA_VACIA_DMY Then
            MsgBox "Ingrese la fecha de exoneración inicial", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not sighentidades.EsFecha(Me.txtFexoneracion1, "DD/MM/AAAA") Then
                MsgBox "La fecha de exoneración inicial no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        
        If Me.txtFexoneracion2 = sighentidades.FECHA_VACIA_DMY Then
            MsgBox "Ingrese la fecha de exoneración final", vbInformation, Me.Caption
            Exit Sub
        Else
            If Not sighentidades.EsFecha(Me.txtFexoneracion2, "DD/MM/AAAA") Then
                MsgBox "La fecha de exoneración final no tiene el formato correcto", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        If CDate(Me.txtFexoneracion1.Text) > CDate(Me.txtFexoneracion2.Text) Then
           MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
           Exit Sub
        End If
        
        oRptExoneraciones.FechaInicio = txtFexoneracion1.Text
        oRptExoneraciones.FechaFin = txtFexoneracion2.Text
        oRptExoneraciones.TextoDelFiltro = "Filtros:     F.Exoneración: (" & txtFexoneracion1.Text & " - " & txtFexoneracion2.Text & ")" & _
                                           IIf(cmbServicioSocial.Text = "", "", "  (Empleado: " & cmbServicioSocial.Text & ")")
                                           
        oRptExoneraciones.CrearReportePorEmpleadoDeServicioSocial Val(mo_cmbServicioSocial.BoundText)
    
    End If
    Set oRptExoneraciones = Nothing
    Me.MousePointer = 1
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub


Private Sub Form_Load()
    Me.txtFechaInicio.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
    Me.txtFechaFin.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
    Me.txtFexoneracion1.Text = sighentidades.PrimerFechaDDMMYYDelMesActual
    Me.txtFexoneracion2.Text = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
    Me.finicio = sighentidades.PrimerFechaDDMMYYDelMesActual
    Me.ffinal = Format(Date, sighentidades.DevuelveFechaSoloFormato_DMY)
    
    '
    Set mo_cmbServicioSocial.MiComboBox = cmbServicioSocial
    mo_cmbServicioSocial.BoundColumn = "IdEmpleado"
    mo_cmbServicioSocial.ListField = "Empleado"
    Set mo_cmbServicioSocial.RowSource = mo_ReglasFacturacion.EmpleadosSeleccionarPorFiltro("Where idLaboraArea=" & sghAreasLaboraEmpleado.sghSeguros & " and idLaboraSubArea= 9")
    
End Sub


Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub txtFechaFin_LostFocus()
    If txtFechaFin <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaFin, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaFin = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub


Private Sub txtFechaInicio_LostFocus()
    If txtFechaInicio <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFechaInicio, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFechaInicio = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub


Private Sub txtFexoneracion1_LostFocus()
    If txtFexoneracion1 <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFexoneracion1, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFexoneracion1 = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub

Private Sub txtFexoneracion2_LostFocus()
    If txtFexoneracion2 <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFexoneracion2, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFexoneracion2 = sighentidades.FECHA_VACIA_DMY
        End If
    End If
End Sub
