VERSION 5.00
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGTHRE~1.OCX"
Begin VB.Form RpCatServicios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Catálogo de Servicios"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "rpCatServicios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   6750
      Begin Threed.SSOption optServicios 
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   2295
         _ExtentX        =   4048
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
         Caption         =   "Lista todos los Servicios"
         Value           =   -1
      End
      Begin Threed.SSOption optPtosCarga 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   870
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
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
         Caption         =   "Por Punto de Carga"
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   60
      TabIndex        =   2
      Top             =   2100
      Width           =   6735
      Begin VB.CommandButton btnAceptar 
         Caption         =   "Aceptar (F2)"
         DisabledPicture =   "rpCatServicios.frx":0CCA
         DownPicture     =   "rpCatServicios.frx":112A
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
         Left            =   1830
         Picture         =   "rpCatServicios.frx":159F
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.CommandButton btnCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar (ESC)"
         DisabledPicture =   "rpCatServicios.frx":1A14
         DownPicture     =   "rpCatServicios.frx":1ED8
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
         Left            =   3360
         Picture         =   "rpCatServicios.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   1365
      End
   End
End
Attribute VB_Name = "RpCatServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Catálogo de Servicios
'        Programado por: Barrantes D
'        Fecha: Setiembre 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Private Sub btnAceptar_Click()
     Me.MousePointer = 11
     If optServicios.Value = True Then
        Dim oRepEnGeneral As New clCatalogoServicios
        oRepEnGeneral.ListaServiciosEnGeneral Me.hwnd
        Set oRepEnGeneral = Nothing
     ElseIf optPtosCarga.Value = True Then
        Dim oRepPtoCarga As New clCatalogoServicios
        oRepPtoCarga.ListaServiciosPorPuntosDeCarga Me.hwnd
        Set oRepPtoCarga = Nothing
     End If
     Me.MousePointer = 1
End Sub



Private Sub btnCancelar_Click()
    Me.Visible = False
End Sub
