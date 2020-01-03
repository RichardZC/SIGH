VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGULTR~1.OCX"
Begin VB.Form ConfiguraLabRes 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   Icon            =   "ConfiguraLabRes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   45
      TabIndex        =   24
      Top             =   7695
      Width           =   11295
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "ConfiguraLabRes.frx":0CCA
         DownPicture     =   "ConfiguraLabRes.frx":118E
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
         Left            =   5910
         Picture         =   "ConfiguraLabRes.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "ConfiguraLabRes.frx":1B66
         DownPicture     =   "ConfiguraLabRes.frx":1FC6
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
         Left            =   4350
         Picture         =   "ConfiguraLabRes.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   225
         Width           =   1365
      End
   End
   Begin VB.Frame Frame 
      Caption         =   " Datos de la Configuración"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7650
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   11295
      Begin VB.TextBox txtCodigo 
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
         Left            =   1905
         TabIndex        =   36
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton btnModificaItem 
         Height          =   450
         Left            =   10545
         Picture         =   "ConfiguraLabRes.frx":28B0
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1665
         Width           =   660
      End
      Begin VB.CommandButton btnNuevoItem 
         DisabledPicture =   "ConfiguraLabRes.frx":2C80
         DownPicture     =   "ConfiguraLabRes.frx":3069
         Height          =   450
         Left            =   10545
         Picture         =   "ConfiguraLabRes.frx":3475
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1215
         Width           =   660
      End
      Begin VB.CommandButton btnEliminarItem 
         DisabledPicture =   "ConfiguraLabRes.frx":3881
         DownPicture     =   "ConfiguraLabRes.frx":3C0C
         Height          =   450
         Left            =   10545
         Picture         =   "ConfiguraLabRes.frx":3F9F
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Eliminar un valor del Lista Despegable"
         Top             =   2115
         Width           =   660
      End
      Begin VB.CommandButton btnCancelarItem 
         Cancel          =   -1  'True
         DisabledPicture =   "ConfiguraLabRes.frx":4330
         DownPicture     =   "ConfiguraLabRes.frx":47F4
         Enabled         =   0   'False
         Height          =   450
         Left            =   10545
         Picture         =   "ConfiguraLabRes.frx":4CE0
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3030
         Width           =   660
      End
      Begin VB.CommandButton btnGrabarItem 
         DisabledPicture =   "ConfiguraLabRes.frx":51CC
         DownPicture     =   "ConfiguraLabRes.frx":562C
         Enabled         =   0   'False
         Height          =   450
         Left            =   10545
         Picture         =   "ConfiguraLabRes.frx":5AA1
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2580
         Width           =   660
      End
      Begin UltraGrid.SSUltraGrid grdItems 
         Height          =   3075
         Left            =   165
         TabIndex        =   27
         Top             =   1245
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   5424
         _Version        =   131072
         GridFlags       =   17040384
         LayoutFlags     =   67108864
         Caption         =   "Items"
      End
      Begin VB.TextBox txtLabItemCpt 
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
         Left            =   6915
         TabIndex        =   21
         Top             =   780
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame 
         Caption         =   "Datos del Item"
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
         Height          =   3015
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   4485
         Width           =   11055
         Begin VB.TextBox txtValorCombo 
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
            Left            =   6840
            TabIndex        =   30
            Top             =   825
            Visible         =   0   'False
            Width           =   3975
         End
         Begin UltraGrid.SSUltraGrid grdOpciones 
            Height          =   2025
            Left            =   6120
            TabIndex        =   28
            Top             =   825
            Visible         =   0   'False
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   3572
            _Version        =   131072
            GridFlags       =   17040384
            LayoutFlags     =   67108864
            Caption         =   "Opciones de Lista Desplegable"
         End
         Begin VB.CommandButton btnItem 
            Caption         =   "..."
            Height          =   280
            Left            =   5520
            TabIndex        =   14
            ToolTipText     =   "Agregar un Item"
            Top             =   825
            Width           =   375
         End
         Begin VB.CommandButton btnGrupo 
            Caption         =   "..."
            Height          =   280
            Left            =   5520
            TabIndex        =   13
            ToolTipText     =   "Agregar un Grupo de Item"
            Top             =   345
            Width           =   375
         End
         Begin VB.TextBox txtOrden 
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
            Left            =   1920
            TabIndex        =   12
            Top             =   2505
            Width           =   975
         End
         Begin VB.TextBox txtValorReferencial 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1920
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1800
            Width           =   3975
         End
         Begin VB.ComboBox cboTipo 
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
            ItemData        =   "ConfiguraLabRes.frx":5F16
            Left            =   1935
            List            =   "ConfiguraLabRes.frx":5F26
            TabIndex        =   10
            Top             =   1290
            Width           =   3975
         End
         Begin VB.ComboBox cboItem 
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
            Left            =   1935
            TabIndex        =   9
            Top             =   810
            Width           =   3555
         End
         Begin VB.ComboBox cboGrupo 
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
            Left            =   1935
            TabIndex        =   8
            Top             =   330
            Width           =   3555
         End
         Begin VB.TextBox txtMetodo 
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
            Left            =   6840
            TabIndex        =   7
            Top             =   360
            Width           =   3960
         End
         Begin VB.CommandButton btnQuitar 
            DisabledPicture =   "ConfiguraLabRes.frx":5F59
            DownPicture     =   "ConfiguraLabRes.frx":62E4
            Enabled         =   0   'False
            Height          =   315
            Left            =   10260
            Picture         =   "ConfiguraLabRes.frx":6677
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Eliminar un valor del Lista Despegable"
            Top             =   1155
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.CommandButton btnAgregar 
            DisabledPicture =   "ConfiguraLabRes.frx":6A08
            DownPicture     =   "ConfiguraLabRes.frx":6DF1
            Enabled         =   0   'False
            Height          =   315
            Left            =   10260
            Picture         =   "ConfiguraLabRes.frx":71FD
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Agregar Valor del Lista Despegable"
            Top             =   840
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Valor "
            Height          =   195
            Index           =   3
            Left            =   6135
            TabIndex        =   29
            Top             =   825
            Width           =   405
         End
         Begin VB.Label Label 
            Caption         =   "Orden del Resultado"
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
            Index           =   4
            Left            =   240
            TabIndex        =   20
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label Label 
            Caption         =   "Valor Referencial"
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
            Index           =   3
            Left            =   240
            TabIndex        =   19
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label 
            Caption         =   "Tipo de dato"
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
            Index           =   2
            Left            =   240
            TabIndex        =   18
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label 
            Caption         =   "Item"
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
            Index           =   1
            Left            =   240
            TabIndex        =   17
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label 
            Caption         =   "Grupo de Item"
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
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Método"
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
            Index           =   0
            Left            =   6120
            TabIndex        =   15
            Top             =   390
            Width           =   630
         End
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   3345
         TabIndex        =   3
         Top             =   360
         Width           =   7755
      End
      Begin VB.CommandButton btnBusqueda 
         Caption         =   "..."
         Height          =   285
         Left            =   2940
         TabIndex        =   2
         ToolTipText     =   "Selecciona un Servicio"
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox cboGrupoExamen 
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
         Left            =   1905
         TabIndex        =   1
         Top             =   765
         Width           =   3675
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Servicio ( CPT)"
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
         Left            =   270
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grupo de Exámen"
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
         Left            =   285
         TabIndex        =   22
         Top             =   780
         Width           =   1470
      End
   End
End
Attribute VB_Name = "ConfiguraLabRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Configura Resultados de Laboratorio
'        Programado por: Madrid S
'        Fecha: Julio 2014
'
'------------------------------------------------------------------------------------

Dim mo_Apariencia As New sighentidades.GridInfragistic 'configuracion de la grilla
Dim mo_Teclado As New sighentidades.Teclado 'configuracion del uso de teclado
Dim mo_Formulario As New sighentidades.Formulario

Dim mo_cboGrupo As New ListaDespleglable 'manejo de combos
Dim mo_cboItem As New ListaDespleglable 'manejo de combo
Dim mo_cboGrupoExamen As New ListaDespleglable 'manejo de grupo de examen

Dim mi_Opcion As sghOpciones 'manejo de opciones
Dim mo_CatalogoServicios As New DOCatalogoServicio 'estructura del catalogo de servicios

Dim mb_ExistenDatos As Boolean
Dim ml_IdProducto As Long
Dim mo_AdminComun As New ReglasConfiguarcionReslab

Dim mrs_Conflab As New ADODB.Recordset
Dim mrs_Opciones As New ADODB.Recordset

Dim ml_idUsuario As Long
Dim ms_MensajeError As String
Dim mo_lnIdTablaLISTBARITEMS As Long
Dim mo_lcNombrePc As String

Dim conta As Integer
Dim lb_Switch As Boolean
Dim mb_modifica As Boolean
Dim llGrupo As Long
Dim llItem As Long
Dim llGrupoItem As Long


Property Let lcNombrePc(lValue As String)
   mo_lcNombrePc = lValue
End Property
Property Get lcNombrePc() As String
    lcNombrePc = mo_lcNombrePc
End Property
Property Let lnIdTablaLISTBARITEMS(lValue As Long)
   mo_lnIdTablaLISTBARITEMS = lValue
End Property
Property Get lnIdTablaLISTBARITEMS() As Long
   lnIdTablaLISTBARITEMS = mo_lnIdTablaLISTBARITEMS
End Property
Property Let ExistenDatos(bValue As Boolean)
   mb_ExistenDatos = bValue
End Property
Property Get ExistenDatos() As Boolean
   ExistenDatos = mb_ExistenDatos
End Property
Property Let Opcion(iValue As sghOpciones)
   mi_Opcion = iValue
End Property
Property Get Opcion() As sghOpciones
   Opcion = mi_Opcion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Let idUsuario(lValue As Long)
   ml_idUsuario = lValue
End Property
Property Get idUsuario() As Long
   idUsuario = ml_idUsuario
End Property
Property Let idProducto(lValue As Long)
   ml_IdProducto = lValue
End Property
Property Get idProducto() As Long
   idProducto = ml_IdProducto
End Property

Private Sub btnAceptar_Click()
    Dim oDoLabItemsCPT As New DoLabItemsCPT
    Select Case Me.Opcion
        Case sghAgregar
            If txtLabItemCpt.Text <> "" Or cboGrupoExamen.Text <> "" Or mrs_Conflab.RecordCount > 0 Then
                If mo_AdminComun.LabItemCPTAgregar(mrs_Conflab, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") Then
                    MsgBox "Los datos se agregaron correctamente", vbInformation, Me.Caption
                    LimpiarFormulario
                Else
                    MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
                End If
            Else
                MsgBox "No ha completado los datos minimos para crear una nueva configuración", vbInformation, Me.Caption
            End If
        Case sghModificar
            If mo_AdminComun.LabItemCPTModificar(mrs_Conflab, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") Then
                    MsgBox "Los datos se modificaron correctamente", vbInformation, Me.Caption
'                    LimpiarFormulario
                    Unload Me
                Else
                    MsgBox "No se pudo modifcar los datos" + Chr(13) + mo_AdminComun.MensajeError, vbExclamation, Me.Caption
                End If
        Case sghEliminar
            If mo_AdminComun.VerificaExamen(Me.idProducto) Then
                If MsgBox("Está seguro de liminar la configuración de resultados", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                    oDoLabItemsCPT.idProductoCPT = Me.idProducto
                    oDoLabItemsCPT.IdUsuarioAuditoria = Me.idUsuario
                    If mo_AdminComun.LabItemsCPTEliminar(oDoLabItemsCPT, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, "") Then
                        MsgBox "Se eliminó la configuración correctamente", vbOKOnly + vbInformation, Me.Caption
                        Unload Me
                    End If
                End If
            Else
                MsgBox "No se puede borrar la configuración por que está siendo utilizada", vbOKOnly + vbInformation, Me.Caption
            End If
    End Select
End Sub

Private Sub btnAgregar_Click()
    mrs_Opciones.AddNew
    mrs_Opciones.Fields!Dato = InputBox("Ingrese una opción para la lista desplegable", "Ingreso de Opciones")
    mrs_Opciones.Update
    mrs_Opciones.MoveFirst
    Set grdOpciones.DataSource = mrs_Opciones
    grdOpciones.Refresh
End Sub


Private Sub btnCancelarItem_Click()
    Activacontroles
        mo_Formulario.HabilitarDeshabilitar cboGrupo, True
        mo_Formulario.HabilitarDeshabilitar cboItem, True
        mo_Formulario.HabilitarDeshabilitar cboGrupoExamen, True
        mo_Formulario.HabilitarDeshabilitar cboTipo, True
        'mo_Formulario.HabilitarDeshabilitar txtOrden, True
End Sub

Private Sub btnNuevoItem_Click()
'    conta = conta + 1
    conta = mrs_Conflab.RecordCount + 1
    lb_Switch = True
    Activacontroles
    cboItem.Text = ""
    cboTipo.Text = ""
    txtValorReferencial.Text = ""
    txtOrden.Text = conta
    txtMetodo.Text = ""
    Label10(3).Visible = False
    txtValorCombo.Visible = False
    creaOpciones
End Sub

Private Sub btnBusqueda_Click()
    Dim oFrm As New BuscarSSConfiguracion
    
        oFrm.MostrarFormulario
        If oFrm.IdRegistroSeleccionado <> 0 Then
            Me.txtLabItemCpt.Text = CStr(oFrm.IdRegistroSeleccionado)
            Set mo_CatalogoServicios = mo_AdminComun.CatalogoServicioSeleccionarPorId(Val(txtLabItemCpt.Text))
            Me.txtCodigo = mo_CatalogoServicios.Codigo
            Me.txtDescripcion = mo_CatalogoServicios.Nombre
            mo_Formulario.HabilitarDeshabilitar txtLabItemCpt, False
        End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnEliminarItem_Click()
    Select Case Me.Opcion
        Case sghAgregar
            mrs_Conflab.Delete
            mrs_Conflab.MoveFirst
            Set grdItems.DataSource = mrs_Conflab
            grdItems.Refresh
        Case sghModificar
            If mrs_Conflab.Fields!Estado Then
                MsgBox "No se puede eliminar un item de una configuración anterior", vbOKOnly + vbInformation, Me.Caption
            Else
                mrs_Conflab.Delete
                mrs_Conflab.MoveFirst
                Set grdItems.DataSource = mrs_Conflab
                grdItems.Refresh
            End If
    End Select
End Sub

Private Sub btnGrabarItem_Click()
    Activacontroles
    If lb_Switch Then
        If cboTipo.Text = "Lista Desplegable" Then
            mrs_Opciones.MoveFirst
            
            For x = 1 To mrs_Opciones.RecordCount
                mrs_Conflab.AddNew
                mrs_Conflab.Fields!IdUsuarioAuditoria = Me.idUsuario
                mrs_Conflab.Fields!idProductoCPT = txtLabItemCpt.Text
                conta = conta
                mrs_Conflab.Fields!ordenXresultado = conta
                mrs_Conflab.Fields!idGrupo = IIf(IsNull(mo_cboGrupoExamen.BoundText) Or mo_cboGrupoExamen.BoundText = "", llGrupo, Val(mo_cboGrupoExamen.BoundText))
                mrs_Conflab.Fields!idItemGrupo = IIf(IsNull(mo_cboGrupo.BoundText) Or mo_cboGrupo.BoundText = "", llGrupoItem, Val(mo_cboGrupo.BoundText))
                mrs_Conflab.Fields!idItem = IIf(IsNull(mo_cboItem.BoundText) Or mo_cboGrupo.BoundText = "", llItem, Val(mo_cboItem.BoundText))
                mrs_Conflab.Fields!Grupo_item = cboGrupo.Text
                mrs_Conflab.Fields!Item = cboItem.Text
                mrs_Conflab.Fields!idItem = Val(mo_cboItem.BoundText)
                mrs_Conflab.Fields!ValorSiEsCombo = mrs_Opciones.Fields!Dato
                mrs_Conflab.Fields!ValorReferencial = txtValorReferencial.Text
                mrs_Conflab.Fields!Metodo = txtMetodo.Text
                mrs_Conflab.Fields!SoloNumero = False
                mrs_Conflab.Fields!Solotexto = False
                mrs_Conflab.Fields!SoloCombo = True
                mrs_Conflab.Fields!SoloCheck = False
                mrs_Conflab.Fields!Estado = False
                mrs_Conflab.Update
                mrs_Opciones.MoveNext
                conta = conta + 1
            Next x
            conta = conta - 1
            ControlesOpciones (False)
        Else
            mrs_Conflab.AddNew
            mrs_Conflab.Fields!Grupo_item = cboGrupo.Text
            mrs_Conflab.Fields!Item = cboItem.Text
            mrs_Conflab.Fields!IdUsuarioAuditoria = Me.idUsuario
            mrs_Conflab.Fields!idProductoCPT = txtLabItemCpt.Text
            mrs_Conflab.Fields!ordenXresultado = Val(txtOrden.Text)
            mrs_Conflab.Fields!idGrupo = IIf(mo_cboGrupoExamen.BoundText = "", llGrupo, Val(mo_cboGrupoExamen.BoundText))
            mrs_Conflab.Fields!idItemGrupo = IIf(mo_cboGrupo.BoundText = "", llGrupoItem, Val(mo_cboGrupo.BoundText))
            mrs_Conflab.Fields!idItem = IIf(mo_cboItem.BoundText = "", llItem, Val(mo_cboItem.BoundText))
            mrs_Conflab.Fields!ValorSiEsCombo = ""
            mrs_Conflab.Fields!ValorReferencial = txtValorReferencial.Text
            mrs_Conflab.Fields!Metodo = txtMetodo.Text
            If cboTipo.ListIndex = 0 Then mrs_Conflab.Fields!SoloNumero = True Else mrs_Conflab.Fields!SoloNumero = False
            If cboTipo.ListIndex = 1 Then mrs_Conflab.Fields!Solotexto = True Else mrs_Conflab.Fields!Solotexto = False
            mrs_Conflab.Fields!SoloCombo = False
            If cboTipo.ListIndex = 3 Then mrs_Conflab.Fields!SoloCheck = True Else mrs_Conflab.Fields!SoloCheck = False
            mrs_Conflab.Fields!Estado = False
            mrs_Conflab.Update
        End If
    Else
        mrs_Conflab.Fields!Grupo_item = cboGrupo.Text
        mrs_Conflab.Fields!Item = cboItem.Text
        mrs_Conflab.Fields!IdUsuarioAuditoria = Me.idUsuario
        mrs_Conflab.Fields!idProductoCPT = txtLabItemCpt.Text
        mrs_Conflab.Fields!ordenXresultado = Val(txtOrden.Text)
        mrs_Conflab.Fields!idGrupo = IIf(Val(mo_cboGrupoExamen.BoundText) = 0, mrs_Conflab.Fields!idGrupo, Val(mo_cboGrupoExamen.BoundText))
        mrs_Conflab.Fields!idItemGrupo = IIf(Val(mo_cboGrupo.BoundText) = 0, mrs_Conflab.Fields!idItemGrupo, Val(mo_cboGrupo.BoundText))
        mrs_Conflab.Fields!idItem = IIf(Val(mo_cboItem.BoundText) = 0, mrs_Conflab.Fields!idItem, Val(mo_cboItem.BoundText))
        mrs_Conflab.Fields!ValorSiEsCombo = txtValorCombo.Text
        mrs_Conflab.Fields!ValorReferencial = txtValorReferencial.Text
        mrs_Conflab.Fields!Metodo = txtMetodo.Text
        If cboTipo.ListIndex = 0 Then mrs_Conflab.Fields!SoloNumero = True Else mrs_Conflab.Fields!SoloNumero = False
        If cboTipo.ListIndex = 1 Then mrs_Conflab.Fields!Solotexto = True Else mrs_Conflab.Fields!Solotexto = False
        mrs_Conflab.Fields!SoloCombo = False
        If cboTipo.ListIndex = 3 Then mrs_Conflab.Fields!SoloCheck = True Else mrs_Conflab.Fields!SoloCheck = False
        
        mrs_Conflab.Update
    End If
    Set grdItems.DataSource = mrs_Conflab
    Label10(3).Visible = True
    txtValorCombo.Visible = True
End Sub

Private Sub btnGrupo_Click()
    Dim rst As ADODB.Recordset
    Dim oFrmGI As New clLabItemsGrupo
        oFrmGI.idUsuario = Me.idUsuario
        oFrmGI.lcNombrePc = Me.lcNombrePc
        oFrmGI.lnIdTablaLISTBARITEMS = Me.lnIdTablaLISTBARITEMS
        oFrmGI.MostrarFormulario
        If oFrmGI.IdRegistroSeleccionado <> 0 Then
            Set rst = mo_AdminComun.LabItemsGruposSeleccionarTodos("")
            rst.Find ("iditemGrupo=" & oFrmGI.IdRegistroSeleccionado)
            cboGrupo.Text = rst.Fields!Grupo
            llGrupoItem = oFrmGI.IdRegistroSeleccionado
        End If
        rst.Close
        Set rst = Nothing
End Sub

Private Sub btnItem_Click()
    Dim rst As ADODB.Recordset
    Dim oFrmGI As New clLabItems
        oFrmGI.idUsuario = Me.idUsuario
        oFrmGI.lcNombrePc = Me.lcNombrePc
        oFrmGI.lnIdTablaLISTBARITEMS = Me.lnIdTablaLISTBARITEMS
        oFrmGI.MostrarFormulario
        If oFrmGI.IdRegistroSeleccionado <> 0 Then
            Set rst = mo_AdminComun.LabItemsSeleccionarTodos("")
            rst.Find ("idItem=" & oFrmGI.IdRegistroSeleccionado)
            cboItem.Text = rst.Fields!Item
            llItem = oFrmGI.IdRegistroSeleccionado
        End If
        rst.Close
        Set rst = Nothing
End Sub

Private Sub btnModificaItem_Click()
    Activacontroles
    If Not mb_modifica Then
        mo_Formulario.HabilitarDeshabilitar cboGrupo, False
        mo_Formulario.HabilitarDeshabilitar cboItem, False
        mo_Formulario.HabilitarDeshabilitar cboGrupoExamen, False
        mo_Formulario.HabilitarDeshabilitar cboTipo, False
       ' mo_Formulario.HabilitarDeshabilitar txtOrden, False
    Else
        mo_Formulario.HabilitarDeshabilitar cboGrupo, True
        mo_Formulario.HabilitarDeshabilitar cboItem, True
        mo_Formulario.HabilitarDeshabilitar cboGrupoExamen, True
        mo_Formulario.HabilitarDeshabilitar cboTipo, True
       ' mo_Formulario.HabilitarDeshabilitar txtOrden, True
    End If
    lb_Switch = False
End Sub

Private Sub btnQuitar_Click()
    mrs_Opciones.Delete
    mrs_Opciones.MoveFirst
    Set grdOpciones.DataSource = mrs_Opciones
    grdOpciones.Refresh
End Sub


Private Sub cboGrupoExamen_Click()
    Dim x As Integer
    If Me.Opcion = sghModificar Then
        mrs_Conflab.MoveFirst
        For x = o To mrs_Conflab.RecordCount - 1
            mrs_Conflab.Fields!idGrupo = mo_cboGrupoExamen.BoundText
            mrs_Conflab.MoveNext
        Next x
    End If
End Sub


Private Sub cboTipo_Validate(Cancel As Boolean)
    If cboTipo.Text = "Lista Desplegable" Then
        ControlesOpciones (True)
    Else
        ControlesOpciones (False)
    End If
End Sub
Private Sub ControlesOpciones(T As Boolean)
        grdOpciones.Visible = T
        btnAgregar.Visible = T
        btnQuitar.Visible = T
        btnAgregar.Enabled = T
        btnQuitar.Enabled = T
        grdOpciones.Enabled = T
End Sub

Private Sub Form_Activate()
    If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub

Private Sub Form_Initialize()
    Set mo_cboGrupo.MiComboBox = cboGrupo
    Set mo_cboItem.MiComboBox = cboItem
    Set mo_cboGrupoExamen.MiComboBox = cboGrupoExamen
    
End Sub

Private Sub Form_Load()
    mb_modifica = False
    CargarComboBoxes
    CreaTemporal
    Select Case mi_Opcion
    Case sghAgregar
        Me.Caption = "Agregar Configuración de un Procedimiento de Laboratorio"
        conta = 0
    Case sghModificar
        Me.Caption = "Modificar Configuración de un Procedimiento de Laboratorio"
        mo_Formulario.HabilitarDeshabilitar txtLabItemCpt, False
        btnBusqueda.Enabled = False
        CargarRecordsetATemporal
        txtValorCombo.Visible = True
        mb_modifica = mo_AdminComun.VerificaExamen(Me.idProducto)

    Case sghConsultar
        btnAceptar.Enabled = False
        Me.Caption = "Consultar Configuración de un Procedimiento de Laboratorio"
        mo_Formulario.HabilitarDeshabilitar txtLabItemCpt, False
        CargarRecordsetATemporal
        btnNuevoItem.Visible = False
        btnModificaItem.Visible = False
        btnEliminarItem.Visible = False
        btnGrabarItem.Visible = False
        btnCancelarItem.Visible = False
        btnBusqueda.Enabled = False
        txtValorCombo.Visible = True
    Case sghEliminar
        Me.Caption = "Eliminar Configuración de un Procedimiento de Laboratorio"
        mo_Formulario.HabilitarDeshabilitar txtLabItemCpt, False
        CargarRecordsetATemporal
        btnNuevoItem.Visible = False
        btnModificaItem.Visible = False
        btnEliminarItem.Visible = False
        btnGrabarItem.Visible = False
        btnCancelarItem.Visible = False
        btnBusqueda.Enabled = False
        txtValorCombo.Visible = True
    End Select

'    CargarDatosAlFormulario
    'mo_Formulario.HabilitarDeshabilitar txtOrden, False
    mo_Formulario.HabilitarDeshabilitar txtDescripcion, False
    mo_Formulario.HabilitarDeshabilitar txtCodigo, False
    mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
    mo_Apariencia.ConfigurarFilasBiColores grdItems, sighentidades.GrillaConFilasBicolor
    mo_Apariencia.ConfigurarFilasBiColores grdOpciones, sighentidades.GrillaConFilasBicolor
End Sub

Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub

Private Sub grdItems_AfterSelectChange(ByVal SelectChange As UltraGrid.Constants_SelectChange)
    CargarDatosALosControles
End Sub

Private Sub grdItems_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdItems.Bands(0).Columns("IdUsuarioAuditoria").Hidden = True
    grdItems.Bands(0).Columns("idProductoCpt").Hidden = True
    grdItems.Bands(0).Columns("idGrupo").Hidden = True
    grdItems.Bands(0).Columns("idItemGrupo").Hidden = True
    grdItems.Bands(0).Columns("Grupo_Item").Hidden = False
    grdItems.Bands(0).Columns("Grupo_Item").Header.Caption = "Grupo de Items"
    grdItems.Bands(0).Columns("Grupo_Item").Width = 2500
    grdItems.Bands(0).Columns("idItem").Hidden = True
    grdItems.Bands(0).Columns("Item").Header.Caption = "Items"
    grdItems.Bands(0).Columns("Item").Width = 2500
    grdItems.Bands(0).Columns("ValorSiEsCombo").Header.Caption = "Valor si es Combo"
    grdItems.Bands(0).Columns("ValorSiEsCombo").Width = 1500
    grdItems.Bands(0).Columns("ValorReferencial").Header.Caption = "Valor de Referencia"
    grdItems.Bands(0).Columns("ValorReferencial").Width = 1500
    grdItems.Bands(0).Columns("ordenXresultado").Header.Caption = "Orden"
    grdItems.Bands(0).Columns("ordenXresultado").Width = 600
    grdItems.Bands(0).Columns("Metodo").Width = 1000
    grdItems.Bands(0).Columns("SoloNumero").Header.Caption = "Numeros"
    grdItems.Bands(0).Columns("SoloNumero").Width = 700
    grdItems.Bands(0).Columns("SoloTexto").Header.Caption = "Texto"
    grdItems.Bands(0).Columns("SoloTexto").Width = 700
    grdItems.Bands(0).Columns("SoloCombo").Header.Caption = "Lista"
    grdItems.Bands(0).Columns("SoloCombo").Width = 700
    grdItems.Bands(0).Columns("SoloCheck").Header.Caption = "Check"
    grdItems.Bands(0).Columns("SoloCheck").Width = 700
    grdItems.Bands(0).Columns("Estado").Hidden = True
End Sub

Private Sub grdOpciones_InitializeLayout(ByVal Context As UltraGrid.Constants_Context, ByVal Layout As UltraGrid.SSLayout)
    grdOpciones.Bands(0).Columns("Dato").Hidden = False
    grdOpciones.Bands(0).Columns("Dato").Header.Caption = "Opciones"
    grdOpciones.Bands(0).Columns("Dato").Width = 3900
End Sub

Private Sub txtOrden_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtOrden
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtMetodo_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtMetodo
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtValorReferencial_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtValorReferencial
    AdministrarKeyPreview KeyCode
End Sub

Private Sub CreaTemporal()
    If mrs_Conflab.State = adStateOpen Then mrs_Conflab.Close
    With mrs_Conflab
        .Fields.Append "IdUsuarioAuditoria", adInteger, 4
        .Fields.Append "idProductoCpt", adInteger, 4, adFldIsNullable
        .Fields.Append "idGrupo", adInteger, 4, adFldIsNullable
        .Fields.Append "idItemGrupo", adInteger, 4, adFldIsNullable
        .Fields.Append "Grupo_Item", adVarChar, 100
        .Fields.Append "idItem", adInteger, 4, adFldIsNullable
        .Fields.Append "Item", adVarChar, 150
        .Fields.Append "ValorSiEsCombo", adVarChar, 150
        .Fields.Append "ValorReferencial", adVarChar, 150
        .Fields.Append "ordenXresultado", adInteger, 4, adFldIsNullable
        .Fields.Append "Metodo", adVarChar, 150
        .Fields.Append "SoloNumero", adBoolean, 1
        .Fields.Append "SoloTexto", adBoolean, 1
        .Fields.Append "SoloCombo", adBoolean, 1
        .Fields.Append "SoloCheck", adBoolean, 1
        .Fields.Append "Estado", adBoolean, 1
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
    End With
    creaOpciones
End Sub

Private Sub creaOpciones()
    If mrs_Opciones.State = adStateOpen Then mrs_Opciones.Close
    With mrs_Opciones
        .Fields.Append "Dato", adVarChar, 50, adFldIsNullable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
        
    End With
End Sub

Private Sub Activacontroles()
    btnAceptar.Enabled = Not (btnAceptar.Enabled)
    btnNuevoItem.Enabled = Not (btnNuevoItem.Enabled)
    btnModificaItem.Enabled = Not (btnModificaItem.Enabled)
    btnEliminarItem.Enabled = Not (btnEliminarItem.Enabled)
    btnGrabarItem.Enabled = Not (btnGrabarItem.Enabled)
    btnCancelarItem.Enabled = Not (btnCancelarItem.Enabled)
    Frame(2).Enabled = Not (Frame(2).Enabled)
'    mo_Formulario.HabilitarDeshabilitar cboGrupo, False
End Sub

Sub CargarComboBoxes()
    
    mo_cboGrupoExamen.BoundColumn = "idGrupo"
    mo_cboGrupoExamen.ListField = "NombreGrupo"
    Set mo_cboGrupoExamen.RowSource = mo_AdminComun.LabGruposSeleccionarTodos
    
    mo_cboGrupo.BoundColumn = "idItemGrupo"
    mo_cboGrupo.ListField = "Grupo"
    Set mo_cboGrupo.RowSource = mo_AdminComun.LabItemsGruposSeleccionarTodos("")

    mo_cboItem.BoundColumn = "idItem"
    mo_cboItem.ListField = "Item"
    Set mo_cboItem.RowSource = mo_AdminComun.LabItemsSeleccionarTodos("")

End Sub


Sub CargarDatosALosControles()
    txtOrden.Text = mrs_Conflab.Fields!ordenXresultado
    cboGrupo.Text = mrs_Conflab.Fields!Grupo_item
    llItemgrupo = mrs_Conflab.Fields!idGrupo
    llGrupoItem = mrs_Conflab.Fields!idItemGrupo
    cboItem.Text = mrs_Conflab.Fields!Item
    txtValorCombo = mrs_Conflab.Fields!ValorSiEsCombo
    txtValorReferencial.Text = mrs_Conflab.Fields!ValorReferencial
    txtMetodo.Text = mrs_Conflab.Fields!Metodo
    If mrs_Conflab.Fields!SoloNumero Then cboTipo.ListIndex = 0
    If mrs_Conflab.Fields!Solotexto Then cboTipo.ListIndex = 1
    If mrs_Conflab.Fields!SoloCombo Then
        cboTipo.ListIndex = 2
        txtValorCombo.Visible = True
        Label10(3).Visible = True
    Else
        txtValorCombo.Visible = False
        Label10(3).Visible = False
    End If
    If mrs_Conflab.Fields!SoloCheck Then cboTipo.ListIndex = 3
End Sub

Sub CargarRecordsetATemporal()
    Dim Entrada As New ADODB.Recordset
    Dim wRecItems As New ADODB.Recordset
    Dim wRecGrupoItems As New ADODB.Recordset
    Dim wRecGrupo As New ADODB.Recordset

    Set wRecItems = mo_AdminComun.LabItemsSeleccionarTodos("")
    Set wRecGrupoItems = mo_AdminComun.LabItemsGruposSeleccionarTodos("")
    Set wRecGrupo = mo_AdminComun.LabGruposSeleccionarTodos
    
    Set Entrada = mo_AdminComun.LabItemsCPTSeleccionarPorIdRecordset(Me.idProducto)
    If mo_AdminComun.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos" & Chr(13) & mo_AdminComun.MensajeError, vbInformation, Me.Caption
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Entrada.RecordCount > 0 Then
        Me.txtLabItemCpt = Entrada.Fields!idProductoCPT
        
        Set mo_CatalogoServicios = mo_AdminComun.CatalogoServicioSeleccionarPorId(Val(txtLabItemCpt.Text))
        Me.txtDescripcion = mo_CatalogoServicios.Nombre
        Me.txtCodigo = mo_CatalogoServicios.Codigo
        Entrada.MoveFirst
        wRecGrupo.MoveFirst
        wRecGrupo.Find ("idGrupo = " & Entrada.Fields!idGrupo)
        cboGrupoExamen.Text = wRecGrupo.Fields!NombreGrupo
        llGrupo = Entrada.Fields!idGrupo
        For x = 0 To Entrada.RecordCount - 1
            wRecItems.MoveFirst
            wRecItems.Find ("idItem = " & Entrada.Fields!idItem)
            wRecGrupoItems.MoveFirst
            wRecGrupoItems.Find ("idItemGrupo = " & Entrada.Fields!idItemGrupo)
            
            mrs_Conflab.AddNew
            mrs_Conflab.Fields!Grupo_item = wRecGrupoItems.Fields!Grupo
            mrs_Conflab.Fields!Item = wRecItems.Fields!Item
            mrs_Conflab.Fields!IdUsuarioAuditoria = Me.idUsuario
            mrs_Conflab.Fields!idProductoCPT = Entrada.Fields!idProductoCPT
            mrs_Conflab.Fields!ordenXresultado = Entrada.Fields!ordenXresultado
            mrs_Conflab.Fields!idGrupo = Entrada.Fields!idGrupo
            mrs_Conflab.Fields!idItemGrupo = Entrada!idItemGrupo
            mrs_Conflab.Fields!idItem = Entrada.Fields!idItem
            mrs_Conflab.Fields!ValorSiEsCombo = IIf(IsNull(Entrada.Fields!ValorSiEsCombo), "", Entrada.Fields!ValorSiEsCombo)
            mrs_Conflab.Fields!ValorReferencial = IIf(IsNull(Entrada.Fields!ValorReferencial), "", Entrada.Fields!ValorReferencial)
            mrs_Conflab.Fields!Metodo = IIf(IsNull(Entrada.Fields!Metodo), "", Entrada.Fields!Metodo)
            mrs_Conflab.Fields!SoloNumero = IIf(IsNull(Entrada.Fields!SoloNumero), False, Entrada.Fields!SoloNumero)
            mrs_Conflab.Fields!Solotexto = IIf(IsNull(Entrada.Fields!Solotexto), False, Entrada.Fields!Solotexto)
            mrs_Conflab.Fields!SoloCombo = IIf(IsNull(Entrada.Fields!SoloCombo), False, Entrada.Fields!SoloCombo)
            mrs_Conflab.Fields!SoloCheck = IIf(IsNull(Entrada.Fields!SoloCheck), False, Entrada.Fields!SoloCheck)
            mrs_Conflab.Fields!Estado = True
            mrs_Conflab.Update
            Entrada.MoveNext
        Next x
        mb_ExistenDatos = True
    Else
        MsgBox "No se puede configurar un grupo" & Chr(13) & mo_AdminComun.MensajeError, vbInformation, Me.Caption
        mb_ExistenDatos = False
        Exit Sub
    End If
    Set Me.grdItems.DataSource = mrs_Conflab
    wRecGrupoItems.Close
    wRecItems.Close
    wRecGrupo.Close
End Sub

Sub LimpiarFormulario()
    conta = 0
    txtLabItemCpt.Text = ""
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
    cboGrupoExamen.Text = ""
    cboGrupo.Text = ""
    cboItem.Text = ""
    cboTipo.Text = ""
    txtValorReferencial.Text = ""
    txtOrden.Text = ""
    txtMetodo.Text = ""
    txtValorCombo.Text = ""

    CreaTemporal

End Sub
