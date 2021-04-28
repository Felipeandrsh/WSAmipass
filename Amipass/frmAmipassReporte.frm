VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAmipassReporte 
   Caption         =   "Reporte de Transacciones Amipass"
   ClientHeight    =   7980
   ClientLeft      =   8550
   ClientTop       =   3285
   ClientWidth     =   11445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3075
      Left            =   480
      TabIndex        =   9
      Top             =   4605
      Width           =   10545
      Begin MSComctlLib.ListView LVAmiReporte 
         Height          =   2535
         Left            =   135
         TabIndex        =   10
         Top             =   270
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Turno"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo Transaccion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ID Transaccion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Rut Completo"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Respuesta Json"
      Height          =   3500
      Left            =   5500
      TabIndex        =   6
      Top             =   1000
      Width           =   5500
      Begin VB.TextBox txtSalida 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   330
         Width           =   5190
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reporte de Transacciones (Diario)"
      Height          =   3500
      Left            =   500
      TabIndex        =   0
      Top             =   1000
      Width           =   5000
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3400
         TabIndex        =   8
         Top             =   2800
         Width           =   1300
      End
      Begin VB.TextBox txtCodeLocal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2000
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "0"
         Top             =   950
         Width           =   2700
      End
      Begin VB.TextBox txtDate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2000
         TabIndex        =   2
         Top             =   400
         Width           =   2700
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   2000
         TabIndex        =   1
         Top             =   2800
         Width           =   1300
      End
      Begin VB.Label lblCodLocal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cod Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   705
         TabIndex        =   5
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1110
         TabIndex        =   4
         Top             =   525
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmAmipassReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variables Venta
Dim sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String, sSalida As String, sInputJson As String

'Variables reporte
Dim sCodLocalReporte As String, sFecha As String, sRutCliente As String

Dim sRespuesta As String, sNumTransaccion As String

Public itmx As ListItem

Dim Obj_response As Object


Private Sub cmdReporte_Click()
     
    Dim resp As Object
    Dim SpMensajeError As String
    
    sCodLocalReporte = txtCodeLocal.Text
    sFecha = txtDate.Text
    sRespuesta = getTransactionReports(sFecha, sCodLocalReporte)
        
    'Transforma a Json
    Set Obj_response = JSON.parse(sRespuesta)
    
    'Verifica respuesta
    If Obj_response.Item("status") <> 200 Then
        SpMensajeError = "Estado: " & Obj_response.Item("status") & " " & vbCrLf & _
                         "Respuesta: No se pudo obtener reporte"
        MsgBox SpMensajeError
        Exit Sub
    End If
         
    frmAmipassReporte.LVAmiReporte.ListItems.Clear
    'Iteramos
    For Each resp In Obj_response.Item("response")
        
        'LLENA GRILLA
        Set itmx = frmAmipassReporte.LVAmiReporte.ListItems.Add(, , resp.Item("sTurno"))
        itmx.SubItems(1) = resp.Item("sTipoTransaccion")
        itmx.SubItems(2) = resp.Item("dTransaccion")
        itmx.SubItems(3) = resp.Item("idTransaccion")
        itmx.SubItems(4) = Format(resp.Item("nMonto"), "#,##0")
        itmx.SubItems(5) = resp.Item("sRutCompleto")
    
    Next
    
    txtSalida = sRespuesta
    
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    sCodLocal = "76449"
    sFecha = "2021-03-29"
    
    txtDate = sFecha
    txtCodeLocal = sCodLocal
End Sub
