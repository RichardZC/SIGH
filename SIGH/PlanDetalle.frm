VERSION 5.00
Object = "{5A9433E9-DD7B-4529-91B6-A5E8CA054615}#2.0#0"; "IGUltraGrid20.ocx"
Begin VB.Form PlanDetalle 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "PlanDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   60
      TabIndex        =   10
      Top             =   4770
      Width           =   7995
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "PlanDetalle.frx":0CCA
         DownPicture     =   "PlanDetalle.frx":118E
         Height          =   700
         Left            =   4065
         Picture         =   "PlanDetalle.frx":167A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "PlanDetalle.frx":1B66
         DownPicture     =   "PlanDetalle.frx":1FC6
         Height          =   700
         Left            =   2520
         Picture         =   "PlanDetalle.frx":243B
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipos y fuentes de financiamiento"
      Height          =   1095
      Left            =   60
      TabIndex        =   7
      Top             =   1080
      Width           =   7995
      Begin VB.ComboBox cmbIdFuenteFinanciamiento 
         Height          =   330
         Left            =   2325
         TabIndex        =   19
         Top             =   660
         Width           =   3270
      End
      Begin VB.ComboBox cmbIdTipoFinanciamiento 
         Height          =   330
         Left            =   2325
         TabIndex        =   18
         Top             =   300
         Width           =   3255
      End
      Begin VB.CommandButton btnEliminar 
         DisabledPicture =   "PlanDetalle.frx":28B0
         DownPicture     =   "PlanDetalle.frx":2C3B
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Picture         =   "PlanDetalle.frx":2FCE
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   315
         Width           =   1005
      End
      Begin VB.CommandButton btnAgregar 
         DisabledPicture =   "PlanDetalle.frx":335F
         DownPicture     =   "PlanDetalle.frx":3748
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5775
         Picture         =   "PlanDetalle.frx":3B54
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label lblIdTipoFinanciamiento 
         Caption         =   "Tipo de financiamiento"
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   330
         Width           =   1875
      End
      Begin VB.Label lblIdFuenteFinanciamiento 
         Caption         =   "Fuente de financiamiento"
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   690
         Width           =   2190
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del plan"
      Height          =   1065
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   7995
      Begin VB.TextBox txtCoaseguro 
         Height          =   315
         Left            =   3495
         TabIndex        =   3
         Top             =   630
         Width           =   1000
      End
      Begin VB.TextBox txtDeducible 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Top             =   630
         Width           =   1000
      End
      Begin VB.TextBox txtIdPlan 
         Height          =   315
         Left            =   1215
         TabIndex        =   0
         Top             =   270
         Width           =   1000
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   3495
         MaxLength       =   50
         TabIndex        =   1
         Top             =   270
         Width           =   4035
      End
      Begin VB.Label lblCoaseguro 
         Caption         =   "Coaseguro"
         Height          =   315
         Left            =   2370
         TabIndex        =   12
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label lblDeducible 
         Caption         =   "Deducible"
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblIdPlan 
         Caption         =   "Plan"
         Height          =   315
         Left            =   195
         TabIndex        =   6
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Descripción"
         Height          =   315
         Left            =   2370
         TabIndex        =   5
         Top             =   300
         Width           =   1005
      End
   End
   Begin UltraGrid.SSUltraGrid grdTiposFuentesFinanciamiento 
      Height          =   2475
      Left            =   75
      TabIndex        =   17
      Top             =   2235
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   4366
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
      Caption         =   "Tipos y fuentes de financiamiento"
   End
End
Attribute VB_Name = "PlanDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: POPlanes
'        Autor: William Castro Grijalva
'        Fecha: 30/08/2004 11:41:52 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------

Dim mo_Teclado As New SIGHComun.Teclado
Dim mo_Formulario As New SIGHComun.Formulario
Dim ml_IdUsuario As Long
Dim ms_MensajeError As String
Dim mi_Opcion As sghOpciones
Dim mb_ExistenDatos As Boolean
Dim mo_AdminFacturacion As New ReglasFacturacion
Dim ml_IdPlan As Long
Dim mrs_TipoYFuenteFinanciamiento As New Recordset
Dim mo_PlanFinanciamiento As New Collection
Dim mo_cmbIdTipoFinanciamiento As New ListaDespleglable
Dim mo_cmbIdFuenteFinanciamiento As New ListaDespleglable

Property Let IdPlan(lValue As Long)
   ml_IdPlan = lValue
End Property
Property Get IdPlan() As Long
   IdPlan = ml_IdPlan
End Property

Sub CargarComboBoxes()
Dim sSQL As String
Dim sMensaje As String
       
       mo_cmbIdTipoFinanciamiento.BoundColumn = "IdTipoFinanciamiento"
       mo_cmbIdTipoFinanciamiento.ListField = "DescripcionLarga"
       Set mo_cmbIdTipoFinanciamiento.RowSource = mo_AdminFacturacion.TiposFinanciamientoSeleccionarTodos()
       sMensaje = sMensaje + mo_AdminFacturacion.MensajeError

       If sMensaje <> "" Then
           MsgBox mo_AdminFacturacion.MensajeError, vbCritical, Me.Caption
       End If

End Sub
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



Private Sub cmbIdTipoFinanciamiento_Click()
Dim sMensaje As String

       If mo_cmbIdTipoFinanciamiento.BoundText = "" Then
        Exit Sub
       End If
       
       mo_cmbIdFuenteFinanciamiento.BoundColumn = "IdFuenteFinanciamiento"
       mo_cmbIdFuenteFinanciamiento.ListField = "DescripcionLarga"
       Set mo_cmbIdFuenteFinanciamiento.RowSource = mo_AdminFacturacion.FuentesFinanciamientoSeleccionarPorTipo(mo_cmbIdTipoFinanciamiento.BoundText)
        mo_cmbIdFuenteFinanciamiento.BoundText = ""
        
       sMensaje = sMensaje + mo_AdminFacturacion.MensajeError
       If sMensaje <> "" Then
           MsgBox sMensaje, vbCritical, Me.Caption
       End If
End Sub

Private Sub Form_Initialize()
    
    Set mo_cmbIdTipoFinanciamiento.MiComboBox = cmbIdTipoFinanciamiento
    Set mo_cmbIdFuenteFinanciamiento.MiComboBox = cmbIdFuenteFinanciamiento
End Sub

Private Sub txtDeducible_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDeducible
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtDeducible_LostFocus()
   mo_Formulario.MarcarComoVacio txtDeducible
End Sub

Private Sub txtDeducible_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtCoaseguro_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtCoaseguro
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtCoaseguro_LostFocus()
   mo_Formulario.MarcarComoVacio txtCoaseguro
End Sub

Private Sub txtCoaseguro_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsNumero(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtDescripcion
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtDescripcion_LostFocus()
   mo_Formulario.MarcarComoVacio txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub


Private Sub txtIdPlan_KeyDown(KeyCode As Integer, Shift As Integer)
   mo_Teclado.RealizarNavegacion KeyCode, txtIdPlan
AdministrarKeyPreview KeyCode
End Sub


Private Sub txtIdPlan_LostFocus()
   mo_Formulario.MarcarComoVacio txtIdPlan
End Sub

Private Sub txtIdPlan_KeyPress(KeyAscii As Integer)
   If Not mo_Teclado.CodigoAsciiEsEspecial(KeyAscii) Then
       If Not mo_Teclado.CodigoAsciiEsLetra(KeyAscii) Then
           KeyAscii = 0
       End If
   End If
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Planes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosAlFormulario()

 Select Case mi_Opcion
     Case sghAgregar
     Case sghModificar
         CargarDatosALosControles
     Case sghConsultar
         CargarDatosALosControles
     Case sghEliminar
         CargarDatosALosControles
 End Select
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Planes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub Form_Load()
       
        GenerarRecordsetTemporal
        
       Select Case mi_Opcion
       Case sghAgregar
           Me.Caption = "Agregar Planes"
       Case sghModificar
           Me.Caption = "Modificar Planes"
       Case sghConsultar
           Me.Caption = "Consultar Planes"
       Case sghEliminar
           Me.Caption = "Eliminar Planes"
       End Select
    
       CargarComboBoxes
       CargarDatosAlFormulario
       mo_Formulario.ConfigurarTipoLetra "Tahoma", "9", Me
End Sub

'------------------------------------------------------------------------------------
'   CargarDatosAlFormulario
'   Descripción:    Seleccionar un registro unico de la tabla Planes
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
                   MsgBox "No se pudo agregar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
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
                   MsgBox "No se pudo modificar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
               End If
           End If
       End If
   Case sghEliminar
           If ValidarReglas() Then
               If EliminarDatos() Then
                   MsgBox " Los datos se eliminaron correctamente", vbInformation, Me.Caption
                   Me.Visible = False
               Else
                   MsgBox "No se pudo eliminar los datos" + Chr(13) + mo_AdminFacturacion.MensajeError, vbExclamation, Me.Caption
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
   If Me.txtDescripcion.Text = "" Then
       sMensaje = sMensaje + "Ingrese el valor de Descripcion" + Chr(13)
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
'   Descripción:    Seleccionar un registro unico de la tabla Planes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargaDatosAlObjetosDeDatos()
Dim oPlanFinanciamiento As DOPlanFinanciamiento


   With mo_Planes
           .descripcion = Me.txtDescripcion.Text
           .IdPlan = Me.IdPlan
           .Coaseguro = Me.txtCoaseguro
           .Deducible = Me.txtDeducible
   End With
   
    'Busca IdPrestamo que se van a excluir
    Dim oRow As SSRow
    Set oRow = Me.grdTiposFuentesFinanciamiento.GetRow(ssChildRowFirst)
    If Not oRow Is Nothing Then
        Set oPlanFinanciamiento = New DOPlanFinanciamiento
        oPlanFinanciamiento.IdPlan = Me.txtIdPlan
        oPlanFinanciamiento.IdFuenteFinanciamiento = Val(oRow.Cells("IdFuenteFinanciamiento").Value)
        oPlanFinanciamiento.IdTipoFinanciamiento = Val(oRow.Cells("IdTipoFinanciamiento").Value)
        oPlanFinanciamiento.IdUsuarioAuditoria = ml_IdUsuario
        mo_PlanFinanciamiento.Add oPlanFinanciamiento
        Do While oRow.HasNextSibling
            Set oRow = oRow.GetSibling(ssSiblingRowNext)
            Set oPlanFinanciamiento = New DOPlanFinanciamiento
            oPlanFinanciamiento.IdPlan = Me.txtIdPlan
            oPlanFinanciamiento.IdFuenteFinanciamiento = oRow.Cells("IdFuenteFinanciamiento").Value
            oPlanFinanciamiento.IdTipoFinanciamiento = oRow.Cells("IdTipoFinanciamiento").Value
            oPlanFinanciamiento.IdUsuarioAuditoria = ml_IdUsuario
            mo_PlanFinanciamiento.Add oPlanFinanciamiento
        Loop
    End If
   
End Sub

'------------------------------------------------------------------------------------
'        Agregar Datos
'------------------------------------------------------------------------------------

Function AgregarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   AgregarDatos = mo_AdminFacturacion.PlanesAgregar(mo_Planes, mo_PlanFinanciamiento)

End Function

'------------------------------------------------------------------------------------
'        Modificar Datos
'------------------------------------------------------------------------------------

Function ModificarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   ModificarDatos = mo_AdminFacturacion.PlanesModificar(mo_Planes, mo_PlanFinanciamiento)

End Function

'------------------------------------------------------------------------------------
'        Eliminar Datos
'------------------------------------------------------------------------------------

Function EliminarDatos() As Boolean

   CargaDatosAlObjetosDeDatos
   EliminarDatos = mo_AdminFacturacion.PlanesModificar(mo_Planes, mo_PlanFinanciamiento)

End Function

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Planes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub CargarDatosALosControles()

       Set mo_Planes = mo_AdminFacturacion.PlanesSeleccionarPorId(Me.IdPlan)
        
        If mo_AdminFacturacion.MensajeError <> "" Then
             MsgBox "No se pudo obtener los datos + Chr(13) + mo_AdminFacturacion.MensajeError, vbCritical, Me.Caption"
             mb_ExistenDatos = False
             Exit Sub
        End If
        
        If Not mo_Planes Is Nothing Then
           
           With mo_Planes
                Me.txtDescripcion.Text = .descripcion
                Me.txtIdPlan.Text = .IdPlan
                Me.txtCoaseguro.Text = .Coaseguro
                Me.txtDeducible.Text = .Deducible
                mb_ExistenDatos = True
           End With
           
           Dim rsPlanesFinanciamiento As New Recordset
           Set rsPlanesFinanciamiento = mo_AdminFacturacion.PlanesFinanciamientoSeleccionarPorPlan(Me.IdPlan)
           Do While Not rsPlanesFinanciamiento.EOF
                With mrs_TipoYFuenteFinanciamiento
                    .AddNew
                    .Fields!IdTipoFinanciamiento = rsPlanesFinanciamiento!IdTipoFinanciamiento
                    .Fields!tipoFinanciamiento = rsPlanesFinanciamiento!tipoFinanciamiento
                    .Fields!IdFuenteFinanciamiento = "" & rsPlanesFinanciamiento!IdFuenteFinanciamiento
                    .Fields!FuenteFinanciamiento = "" & rsPlanesFinanciamiento!FuenteFinanciamiento
                End With
                rsPlanesFinanciamiento.MoveNext
           Loop
        Else
            mb_ExistenDatos = False
            Exit Sub
        End If
   
End Sub

'------------------------------------------------------------------------------------
'   Llenar Datos Al Formulario
'   Descripción:    Seleccionar un registro unico de la tabla Planes
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Sub LimpiarFormulario()

        Me.txtDescripcion.Text = ""
        Me.txtIdPlan.Text = ""
        
        Dim i As Integer
        For i = mo_PlanFinanciamiento.Count To 1 Step -1
            mo_PlanFinanciamiento.Remove i
        Next i
               
End Sub


Sub GenerarRecordsetTemporal()
    
    With mrs_TipoYFuenteFinanciamiento
          .Fields.Append "IdTipoFinanciamiento", adVarChar, 10
          .Fields.Append "TipoFinanciamiento", adVarChar, 255
          .Fields.Append "IdFuenteFinanciamiento", adVarChar, 10
          .Fields.Append "FuenteFinanciamiento", adVarChar, 255
          .CursorType = adOpenStatic
          .LockType = adLockOptimistic
          .Open
    End With
    
    Set Me.grdTiposFuentesFinanciamiento.DataSource = mrs_TipoYFuenteFinanciamiento
    
End Sub

Private Sub btnAgregar_Click()
    
    On Error Resume Next
    mrs_TipoYFuenteFinanciamiento.MoveFirst
    Do While Not mrs_TipoYFuenteFinanciamiento.EOF
        If mo_cmbIdTipoFinanciamiento.BoundText = mrs_TipoYFuenteFinanciamiento!IdTipoFinanciamiento _
        And mo_cmbIdFuenteFinanciamiento.BoundText = mrs_TipoYFuenteFinanciamiento!IdFuenteFinanciamiento Then
            MsgBox "La combinacion de tipo y fuente de financiamiento ya existe", vbExclamation, Me.Caption
            Exit Sub
        End If
        mrs_TipoYFuenteFinanciamiento.MoveNext
    Loop
    
    With mrs_TipoYFuenteFinanciamiento
        .AddNew
        .Fields!IdTipoFinanciamiento = mo_cmbIdTipoFinanciamiento.BoundText
        .Fields!tipoFinanciamiento = Me.cmbIdTipoFinanciamiento.Text
        .Fields!IdFuenteFinanciamiento = mo_cmbIdFuenteFinanciamiento.BoundText
        .Fields!FuenteFinanciamiento = Me.cmbIdFuenteFinanciamiento.Text
    End With
    
End Sub

Private Sub btnQuitar_Click()
    On Error Resume Next
    With mrs_TipoYFuenteFinanciamiento
        If Not .EOF And Not .BOF Then
           .Delete
           .Update
        End If
    End With

    Set Me.grdTiposFuentesFinanciamiento.DataSource = mrs_TipoYFuenteFinanciamiento

End Sub
