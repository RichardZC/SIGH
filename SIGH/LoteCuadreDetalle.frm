VERSION 5.00
Begin VB.Form LoteCuadreDetalle 
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   Icon            =   "LoteCuadreDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales"
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
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   7050
      Begin VB.ComboBox cmbLote 
         Height          =   315
         Left            =   900
         TabIndex        =   1
         Top             =   360
         Width           =   6045
      End
      Begin VB.TextBox txtCuadreUsuario 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label lblDiferencia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5160
         TabIndex        =   11
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label lblCalculado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   1020
         Width           =   1395
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencia (S/.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   5280
         TabIndex        =   5
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Lote"
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
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Calculado (S/.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1260
         TabIndex        =   2
         Top             =   780
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuadre Cajero (S/.)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3060
         TabIndex        =   3
         Top             =   780
         Width           =   1590
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1035
      Left            =   60
      TabIndex        =   8
      Top             =   1590
      Width           =   7035
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "LoteCuadreDetalle.frx":0CCA
         DownPicture     =   "LoteCuadreDetalle.frx":112A
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
         Left            =   2175
         Picture         =   "LoteCuadreDetalle.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "LoteCuadreDetalle.frx":1A14
         DownPicture     =   "LoteCuadreDetalle.frx":1ED8
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
         Left            =   3720
         Picture         =   "LoteCuadreDetalle.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "LoteCuadreDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MZD 22/06/2005 [Todo el Archivo]
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: PODiagnosticos
'        Autor: William Castro Grijalva
'        Fecha: 30/08/2004 12:17:18 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
Const ID_TIPO_MONEDA_SOLES = 1

Dim mo_Teclado As New SIGHCOmun.Teclado
Dim mo_Formulario As New SIGHCOmun.Formulario
Dim mo_CajaLoteCuadre As New DOCajaLoteCuadre
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim ml_IdLoteCuadre As Long
Dim mo_AdminCaja As New ReglasCaja
Dim mo_AdminComun As New ReglasComunes
Dim mo_cmbLote  As New SIGHCOmun.ListaDespleglable

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
Property Let IdUsuario(lValue As Long)
   ml_IdUsuario = lValue
End Property
Property Get IdUsuario() As Long
   IdUsuario = ml_IdUsuario
End Property
Property Let IdLoteCuadre(lValue As Long)
   ml_IdLoteCuadre = lValue
End Property
Property Get IdLoteCuadre() As Long
   IdLoteCuadre = ml_IdLoteCuadre
End Property

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
         CargarComboBoxes
     Case sghModificar
         Me.cmbLote.Enabled = False
         CargarDatosALosControles
     Case sghConsultar
         Frame1.Enabled = False
         CargarDatosALosControles
     Case sghEliminar
         Frame1.Enabled = False
         CargarDatosALosControles
 End Select
End Sub

Private Sub cmbLote_Click()
    'Obtenemos el Monto Calculado para el Lote
    Me.lblCalculado = mo_AdminCaja.LoteObtenerMontoCalculado(Val(mo_cmbLote.BoundText))
End Sub

Private Sub cmbLote_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, cmbLote
    AdministrarKeyPreview KeyCode
End Sub

Private Sub Form_Initialize()
    Set mo_cmbLote.MiComboBox = cmbLote
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Cuadre Caja"
       Case sghModificar
           Me.Caption = "Modificar Cuadre Caja"
       Case sghConsultar
           Me.Caption = "Consultar Cuadre Caja"
       Case sghEliminar
           Me.Caption = "Eliminar Cuadre Caja"
       End Select
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Activate()
   If mi_Opcion <> sghAgregar Then
       If Not mb_ExistenDatos Then
           Me.Visible = False
       End If
   End If
End Sub
Sub AdministrarKeyPreview(KeyCode As Integer)
   Select Case KeyCode
       Case vbKeyEscape
           btnCancelar_Click
       Case vbKeyF2
           btnAceptar_Click
       End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   AdministrarKeyPreview KeyCode
End Sub

Private Sub btnAceptar_Click()
   Select Case mi_Opcion
   Case sghAgregar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If AgregarDatos() Then
                   MsgBox " Los datos se agregaron correctamente", vbInformation, Me.Caption
                   LimpiarFormulario
               Else
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminCaja.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghModificar
       If ValidarDatosObligatorios() Then
           If ValidarReglas() Then
               If ModificarDatos() Then
                   MsgBox " Los datos se modificaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminCaja.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminCaja.MensajeError, vbExclamation, Me.Caption
               End If
           End If
   End Select
End Sub

Private Sub btnCancelar_Click()
   Me.Visible = False
End Sub

Function ValidarDatosObligatorios() As Boolean
   Dim sMensaje As String
   ValidarDatosObligatorios = False
   
   If mo_cmbLote.BoundText = "" Then
       sMensaje = sMensaje + "Ingrese el Lote para el Cuadre" + Chr(13)
   End If
   If Me.txtCuadreUsuario = "" Then
       sMensaje = sMensaje + "Ingrese el monto que informó el cajero" + Chr(13)
   End If
   
   If sMensaje <> "" Then
       MsgBox sMensaje, vbInformation, Me.Caption
       Exit Function
   End If
   ValidarDatosObligatorios = True
End Function
Function ValidarReglas() As Boolean
   ValidarReglas = False
   
   ValidarReglas = True
End Function
'------------------------------------------------------------------------------------
'   Cargar datos al objetos de datos
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
   Me.lblCalculado = Replace(Me.lblCalculado, ".", ",")
   Me.txtCuadreUsuario = Replace(Me.txtCuadreUsuario, ".", ",")
   Me.lblDiferencia = Replace(Me.lblDiferencia, ".", ",")
   
   With mo_CajaLoteCuadre
        .IdLote = Val(mo_cmbLote.BoundText)
        .Calculado = IIf(Me.lblCalculado = "", 0, CDbl(Me.lblCalculado))
        .CuadreUsuario = IIf(Me.txtCuadreUsuario = "", 0, CDbl(Me.txtCuadreUsuario))
        .Diferencia = IIf(Me.lblDiferencia = "", 0, CDbl(Me.lblDiferencia))
        
        .IdUsuarioAuditoria = Me.IdUsuario
   End With
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean
   CargaDatosAlObjetosDeDatos
   mo_CajaLoteCuadre.IdTipoFormaPago = ID_TIPO_MONEDA_SOLES
   AgregarDatos = mo_AdminCaja.LoteCuadreAgregar(mo_CajaLoteCuadre)
   
End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminCaja.LoteCuadreModificar(mo_CajaLoteCuadre)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminCaja.LoteCuadreEliminar(mo_CajaLoteCuadre)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

    Set mo_CajaLoteCuadre = mo_AdminCaja.LoteCuadreSeleccionarPorId(Me.IdLoteCuadre)
    If mo_AdminCaja.MensajeError <> "" Then
        MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminCaja.MensajeError, vbCritical, Me.Caption"
        mb_ExistenDatos = False
        Exit Sub
    End If
    If Not mo_CajaLoteCuadre Is Nothing Then
        'Cargamos el combo box
        CargarComboBoxes
        With mo_CajaLoteCuadre
            mo_cmbLote.BoundText = .IdLote
            Me.lblCalculado = .Calculado
            Me.txtCuadreUsuario = .CuadreUsuario
            Me.lblDiferencia = .Diferencia
            CalcularDiferencia
            
            mb_ExistenDatos = True
        End With
    Else
        mb_ExistenDatos = False
        Exit Sub
    End If
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Diagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

    Me.IdLoteCuadre = 0
    Me.lblCalculado = ""
    Me.txtCuadreUsuario = ""
    Me.lblDiferencia = ""
    Set mo_CajaLoteCuadre = New DOCajaLoteCuadre
    
    CargarComboBoxes
    mo_cmbLote.BoundText = ""
    
End Sub

Sub CargarComboBoxes()
       
    mo_cmbLote.BoundColumn = "IdLote"
    mo_cmbLote.ListField = "Descripcion"
    
    Set mo_cmbLote.RowSource = mo_AdminCaja.LoteSeleccionarPendientesParaLista(mo_CajaLoteCuadre.IdLote)

End Sub

Private Sub txtCuadreUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
    mo_Teclado.RealizarNavegacion KeyCode, txtCuadreUsuario
    AdministrarKeyPreview KeyCode
End Sub

Private Sub txtCuadreUsuario_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsDinero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtCuadreUsuario_LostFocus()
    CalcularDiferencia
End Sub

Private Sub CalcularDiferencia()
    Dim dCalculado As Double
    Dim dCuadre As Double
    Dim dDiferencia As Double
    'Obtenemos la diferencia
    lblCalculado = Replace(lblCalculado, ".", ",")
    txtCuadreUsuario = Replace(txtCuadreUsuario, ".", ",")
    dCalculado = Val(lblCalculado)
    dCuadre = Val(txtCuadreUsuario)
    dDiferencia = Round(dCuadre - dCalculado, 2)
    Me.lblDiferencia = dDiferencia
    If dDiferencia = 0 Then
        Me.lblDiferencia.ForeColor = &H0&   'NEGRO
    ElseIf dDiferencia < 0 Then
        Me.lblDiferencia.ForeColor = &HFF&  'ROJO
    Else
        Me.lblDiferencia.ForeColor = &HFF0000     'AZUL
    End If
End Sub
