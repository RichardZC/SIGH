VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CierreCtaAtencion 
   Caption         =   "Cierre de Cuentas de Atención"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   Icon            =   "CierreCtaAtencion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   3
      Top             =   3990
      Width           =   5835
      Begin VB.CommandButton btnCancelar 
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "CierreCtaAtencion.frx":000C
         DownPicture     =   "CierreCtaAtencion.frx":04D0
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
         Left            =   2910
         Picture         =   "CierreCtaAtencion.frx":09BC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1335
      End
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "CierreCtaAtencion.frx":0EA8
         DownPicture     =   "CierreCtaAtencion.frx":1308
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
         Left            =   1410
         Picture         =   "CierreCtaAtencion.frx":177D
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   210
         Width           =   1365
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
      Height          =   3795
      Left            =   90
      TabIndex        =   2
      Top             =   30
      Width           =   5850
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
         ForeColor       =   &H80000004&
         Height          =   2790
         Left            =   180
         TabIndex        =   5
         Top             =   270
         Width           =   5505
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   345
         Left            =   150
         TabIndex        =   4
         Top             =   3210
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "CierreCtaAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ml_IdUsuario As Long
Dim lcHorasCE As String: Dim lcHorasHosp As String
Dim lcBuscaParametro As New SIGHDatos.Parametros

Property Let IdUsuario(lIdValue As Long)
    ml_IdUsuario = lIdValue
End Property


'***************daniel barrantes**************
'***************Busca si YA ESTA APERTURADA LA CAJA/TURNO/CAJERO/DIA, si es asi ya no APERTURARLA OTRA VEZ
'***************Cierra la ULTIMA CAJA/TURNO/CAJERO que es menor al DIA
Private Sub btnAceptar_Click()
    If MsgBox("Esta seguro que desea CERRAR la cuenta de atención", vbQuestion + vbYesNo, "Facturación") = vbYes Then
        Dim oRsTmp As New ADODB.Recordset
        Dim oConexion As New ADODB.Connection
        Dim lcSql As String
        Dim ml_idCuentaAtencion  As Long
        Dim mo_ReglasFacturacion As New SIGHNegocios.ReglasFacturacion
        Dim mo_DOCuentaAtencion As New DOCuentaAtencion
        Dim lbContinua As Boolean
        Dim lnCant As Long
        Dim lcErrores As String
        Dim ldFecha1 As Date, ldFecha2 As Date

        Me.MousePointer = 1
        oConexion.CommandTimeout = 300
        oConexion.Open sighcomun.CadenaConexion
        lcErrores = ""
        lcSql = "SELECT     dbo.Atenciones.FechaIngreso,dbo.Atenciones.HoraIngreso, dbo.Atenciones.FechaEgreso, dbo.Atenciones.IdCuentaAtencion, dbo.Atenciones.IdDestinoAtencion, " & _
                      " dbo.FacturacionCuentasAtencion.IdEstado, dbo.Pacientes.NroHistoriaClinica, dbo.Atenciones.IdTipoServicio," & _
                      " dbo.Atenciones.FechaEgresoAdministrativo" & _
                " FROM         dbo.Pacientes RIGHT OUTER JOIN" & _
                      " dbo.FacturacionCuentasAtencion ON dbo.Pacientes.IdPaciente = dbo.FacturacionCuentasAtencion.IdPaciente RIGHT OUTER JOIN" & _
                      " dbo.Atenciones ON dbo.FacturacionCuentasAtencion.IdCuentaAtencion = dbo.Atenciones.IdCuentaAtencion" & _
                " Where  dbo.FacturacionCuentasAtencion.IdEstado=1  and dbo.Atenciones.idEstadoAtencion=1"
        oRsTmp.Open lcSql, oConexion, adOpenKeyset, adLockOptimistic
        If oRsTmp.RecordCount > 0 Then
            ProgressBar1.Min = 0
            ProgressBar1.Max = oRsTmp.RecordCount
            lnCant = 1
            oRsTmp.MoveFirst
            Do While Not oRsTmp.EOF
                ProgressBar1.Value = lnCant
                lnCant = lnCant + 1
                lbContinua = False
                ml_idCuentaAtencion = oRsTmp.Fields!idCuentaAtencion
                Select Case oRsTmp.Fields!IdTipoServicio
                Case 1   'ce
                     ldFecha1 = CDate(oRsTmp.Fields!FechaIngreso & " " & oRsTmp.Fields!HoraIngreso)
                     ldFecha2 = Now
                      If DateDiff("h", ldFecha1, ldFecha2) >= Val(lcHorasCE) Then
                         lbContinua = True
                      End If
                Case 3   'hospitalizacion
                     ldFecha1 = CDate(oRsTmp.Fields!FechaIngreso & " " & oRsTmp.Fields!HoraIngreso)
                     ldFecha2 = Now
                      If DateDiff("h", ldFecha1, ldFecha2) >= Val(lcHorasHosp) Then
                         lbContinua = True
                      End If
                End Select
                If lbContinua = True Then
                    Set mo_DOCuentaAtencion = mo_ReglasFacturacion.CuentasAtencionSeleccionarPorId(ml_idCuentaAtencion)
                    mo_DOCuentaAtencion.IdUsuarioAuditoria = ml_IdUsuario
                    If mo_ReglasFacturacion.CuentasAtencionCerrar(mo_DOCuentaAtencion) = False Then
                       lcErrores = lcErrores & oRsTmp.Fields!NroHistoriaClinica & ", "
                    End If
                End If
                oRsTmp.MoveNext
            Loop
        End If
        oRsTmp.Close
        oConexion.Close
        Set oRsTmp = Nothing
        Set oConexion = Nothing
        Me.MousePointer = 11
        If lcErrores <> "" Then
           MsgBox lcErrores, vbInformation, "Cierre"
        End If
        Me.Visible = False
    End If
End Sub

Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub

Private Sub Form_Load()
  
  lcHorasCE = lcBuscaParametro.SeleccionaFilaParametro(209)
  lcHorasHosp = lcBuscaParametro.SeleccionaFilaParametro(233)


  cmbConsideraciones.AddItem "Se cerrarán las 'Cuentas de Atención' siempre y   "
  cmbConsideraciones.AddItem "cuando:                                           "
  cmbConsideraciones.AddItem "1- Las 'Cuentas de Atención' de CONSULTA EXTERNA  "
  cmbConsideraciones.AddItem "   que haigan pasado " & lcHorasCE & " horas después de su  "
  cmbConsideraciones.AddItem "   'CITA'.                                        "
  cmbConsideraciones.AddItem "2- Las 'Cuentas de Atención' de HOSPITALIZACION   "
  cmbConsideraciones.AddItem "   que no esten registradas la CAMA, despues de   "
  cmbConsideraciones.AddItem "   " & lcHorasHosp & "  horas de su Admisión."

End Sub


