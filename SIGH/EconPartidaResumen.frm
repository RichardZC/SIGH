VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form EconPartidaResumen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen x Partidas"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "EconPartidaResumen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosHistoria 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9195
      Begin VB.CheckBox chkIncluyeNotaCredito 
         Alignment       =   1  'Right Justify
         Caption         =   "Incluye columna NOTA DE CREDITO"
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
         Left            =   5685
         Picture         =   "EconPartidaResumen.frx":0CCA
         TabIndex        =   15
         Top             =   990
         Value           =   1  'Checked
         Width           =   3375
      End
      Begin VB.CheckBox chkSoloCredito 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo CREDITOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7485
         TabIndex        =   14
         Top             =   615
         Visible         =   0   'False
         Width           =   1575
      End
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
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   615
         Width           =   3705
      End
      Begin VB.CheckBox chkvalor 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo Partidas Totalizadas Mayor a 0 (Cero)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Picture         =   "EconPartidaResumen.frx":0FDC
         TabIndex        =   11
         Top             =   990
         Width           =   4005
      End
      Begin VB.CheckBox chkExcel 
         Alignment       =   1  'Right Justify
         Caption         =   "En Excel"
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
         Left            =   120
         Picture         =   "EconPartidaResumen.frx":12EE
         TabIndex        =   8
         Top             =   1380
         Visible         =   0   'False
         Width           =   1125
      End
      Begin MSMask.MaskEdBox txtFdesde 
         Height          =   315
         Left            =   1500
         TabIndex        =   0
         Top             =   240
         Width           =   1350
         _ExtentX        =   2381
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
      Begin MSMask.MaskEdBox txtFhasta 
         Height          =   315
         Left            =   6930
         TabIndex        =   1
         Top             =   210
         Width           =   1350
         _ExtentX        =   2381
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
      Begin MSMask.MaskEdBox txtHrInicio 
         Height          =   315
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   750
         _ExtentX        =   1323
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
      Begin MSMask.MaskEdBox txtHrFin 
         Height          =   315
         Left            =   8310
         TabIndex        =   10
         Top             =   210
         Width           =   750
         _ExtentX        =   1323
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cajero"
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
         TabIndex        =   13
         Top             =   660
         Width           =   510
      End
      Begin VB.Label Label4 
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
         Height          =   210
         Left            =   6420
         TabIndex        =   7
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "F.Movimiento"
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
         TabIndex        =   6
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1110
      Left            =   0
      TabIndex        =   3
      Top             =   1950
      Width           =   9180
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "EconPartidaResumen.frx":1600
         DownPicture     =   "EconPartidaResumen.frx":1A60
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
         Left            =   3210
         Picture         =   "EconPartidaResumen.frx":1ED5
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "EconPartidaResumen.frx":234A
         DownPicture     =   "EconPartidaResumen.frx":280E
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
         Left            =   4740
         Picture         =   "EconPartidaResumen.frx":2CFA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "EconPartidaResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: reporte por partida en resumen
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_ReglasFarmacia As New SIGHNegocios.ReglasFarmacia
Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim ms_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim mo_ReglasCaja As New ReglasCaja
Dim sMensaje As String
Dim mo_Teclado As New sighentidades.Teclado
Dim ml_TextoDelFiltro As String
Dim lnIdProducto As Long
Dim mo_Formulario As New sighentidades.Formulario
Dim mo_cmbIdResponsable As New sighentidades.ListaDespleglable
Dim oRsCajeros As New Recordset
Dim lnIdAlmacen As Long
Dim ml_idUsuario As Long

Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property


Private Sub btnAceptar_Click()
    If ValidaDatosObligatorios Then
        Me.MousePointer = 11
        Dim oRpt As New RptEPartidaResumen
        If chkIncluyeNotaCredito.Value = 1 Then
          oRpt.CreaDatosReporteConNotaDeCredito IIf(chkExcel.Value = 1, True, False), _
                                Me.Caption, ml_TextoDelFiltro, _
                                CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)), _
                                CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)), _
                                0, Me.hwnd, IIf(chkvalor.Value = 1, True, False), Val(mo_cmbIdResponsable.BoundText)
        Else
        oRpt.CreaDatosReporte IIf(chkExcel.Value = 1, True, False), _
                              Me.Caption, ml_TextoDelFiltro, _
                              CDate(Format(txtFdesde.Text & " " & txtHrInicio & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)), _
                              CDate(Format(txtFhasta.Text & " " & txtHrFin & ":00", sighentidades.DevuelveFechaSoloFormato_DMY_HMS)), _
                              0, Me.hwnd, IIf(chkvalor.Value = 1, True, False), Val(mo_cmbIdResponsable.BoundText), _
                              IIf(Me.chkSoloCredito.Value = 1, True, False)
        End If
        Set oRpt = Nothing
        Me.MousePointer = 1
    End If
End Sub

Function ValidaDatosObligatorios() As Boolean
    sMensaje = ""
    ml_TextoDelFiltro = "FILTROS:    F.Movimiento: (" & txtFdesde.Text & " " & txtHrInicio.Text & "   al " & txtFhasta.Text & " " & txtHrFin.Text & ") " & IIf(cmbIdResponsable.Text = "", "", "  (Cajero: " & cmbIdResponsable.Text & ")") & IIf(Me.chkSoloCredito.Value = 1, "  (solo CREDITOS)", "  (sin considerar CREDITOS)")
    
    If Me.txtFdesde = sighentidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de movimiento inicial"
    Else
        If Not sighentidades.EsFecha(Me.txtFdesde, "DD/MM/AAAA") Then
            sMensaje = "La fecha de movimiento inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtFhasta = sighentidades.FECHA_VACIA_DMY Then
        sMensaje = "Ingrese la fecha de movimiento final"
    Else
        If Not sighentidades.EsFecha(Me.txtFhasta, "DD/MM/AAAA") Then
            sMensaje = "La fecha de movimiento final no tiene el formato correcto"
        End If
    End If
    
    If Me.txtHrInicio = sighentidades.HORA_VACIA_HM Then
        sMensaje = "Ingrese la hora de movimiento inicial"
    Else
        If Not sighentidades.EsHora(txtHrInicio) Then
            sMensaje = "La hora de movimiento inicial no tiene el formato correcto"
        End If
    End If
    
    If Me.txtHrFin = sighentidades.HORA_VACIA_HM Then
        sMensaje = "Ingrese la hora de movimiento final"
    Else
        If Not sighentidades.EsHora(txtHrFin) Then
            sMensaje = "La hora de movimiento final no tiene el formato correcto"
        End If
    End If
    If CDate(Me.txtFdesde.Text & " " & Me.txtHrInicio.Text) > CDate(Me.txtFhasta.Text & " " & Me.txtHrFin.Text) Then
       MsgBox "La FECHA FINAL debe ser mayor o igual a la FECHA INICIAL", vbInformation, "Reporte"
       Exit Function
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
    LimpiarVariablesDeMemoria
End Sub



Private Sub Form_Initialize()
   Set mo_cmbIdResponsable.MiComboBox = cmbIdResponsable
End Sub

Private Sub Form_Load()
    Dim oRsPermisos As New Recordset
    txtFdesde.Text = Date
    txtFhasta.Text = Date
    txtHrInicio.Text = "00:01"
    txtHrFin.Text = "23:59"
    Set oRsCajeros = mo_ReglasCaja.CajerosSeleccionarTodos()
    mo_cmbIdResponsable.BoundColumn = "IdEmpleado"
    mo_cmbIdResponsable.ListField = "DCajero"
    Set mo_cmbIdResponsable.RowSource = oRsCajeros
    If oRsCajeros.RecordCount > 0 Then
    
        Set oRsPermisos = ms_ReglasSeguridad.UsuariosRolesSeleccionarPermisosTodos(Val(sighentidades.Usuario))
        oRsPermisos.Filter = "idPermiso=1000"
        If oRsPermisos.RecordCount > 0 Then
          mo_cmbIdResponsable.BoundText = sighentidades.Usuario
          mo_Formulario.HabilitarDeshabilitar cmbIdResponsable, False
        End If
        oRsPermisos.Close
'
'
'       oRsCajeros.MoveFirst
'       oRsCajeros.Find "idEmpleado=" & sighentidades.Usuario
'       If Not oRsCajeros.EOF Then
'          mo_cmbIdResponsable.BoundText = sighentidades.Usuario
'          mo_Formulario.HabilitarDeshabilitar cmbIdResponsable, False
'       End If
    End If
    Set oRsPermisos = Nothing
End Sub



Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub






Private Sub Form_Unload(Cancel As Integer)
    LimpiarVariablesDeMemoria
End Sub

Private Sub txtFdesde_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFdesde

End Sub



Private Sub txtFdesde_LostFocus()
    If txtFdesde <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFdesde, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFdesde = sighentidades.FECHA_VACIA_DMY
        End If
    End If

End Sub

Private Sub txtFhasta_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtFhasta

End Sub

Private Sub txtFhasta_LostFocus()
    If txtFhasta <> sighentidades.FECHA_VACIA_DMY Then
        If Not sighentidades.EsFecha(txtFhasta, "DD/MM/AAAA") Then
            MsgBox "La fecha ingresada no es válida", vbInformation, Me.Caption
            txtFhasta = sighentidades.FECHA_VACIA_DMY
        End If
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    AdministrarKeyPreview KeyCode
End Sub

Sub LimpiarVariablesDeMemoria()
    On Error Resume Next
    Set mo_ReglasFarmacia = Nothing
    Set mo_Teclado = Nothing
    Set mo_ReglasFacturacion = Nothing
    Set mo_ReglasComunes = Nothing
    Set mo_Formulario = Nothing
End Sub

Private Sub txtHrFin_LostFocus()
    If txtHrFin <> sighentidades.HORA_VACIA_HM Then
        If Not sighentidades.EsHora(txtHrFin) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHrFin = sighentidades.HORA_VACIA_HM
        End If
    End If
End Sub

Private Sub txtHrInicio_LostFocus()
    If txtHrInicio <> sighentidades.HORA_VACIA_HM Then
        If Not sighentidades.EsHora(txtHrInicio) Then
            MsgBox "La hora ingresada no es válida", vbInformation, Me.Caption
            txtHrInicio = sighentidades.HORA_VACIA_HM
        End If
    End If
End Sub
