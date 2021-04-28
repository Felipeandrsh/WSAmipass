VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAmipassAnulacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación Amipass"
   ClientHeight    =   8910
   ClientLeft      =   10740
   ClientTop       =   3360
   ClientWidth     =   11280
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
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11280
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3075
      Left            =   480
      TabIndex        =   13
      Top             =   4545
      Width           =   10545
      Begin MSComctlLib.ListView LvAmiAnulacion 
         Height          =   2535
         Left            =   150
         TabIndex        =   14
         Top             =   255
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod Respuesta"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cod Autorizacion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Respuesta Json"
      Height          =   3500
      Left            =   5500
      TabIndex        =   10
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
         TabIndex        =   11
         Top             =   330
         Width           =   5190
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Anulacion"
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
         TabIndex        =   12
         Top             =   2800
         Width           =   1300
      End
      Begin VB.TextBox txtNumTransaction 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   5
         Text            =   "0"
         Top             =   400
         Width           =   2700
      End
      Begin VB.TextBox txtQRCode 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   4
         Text            =   "0"
         Top             =   950
         Width           =   2700
      End
      Begin VB.TextBox txtMonto 
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
         Top             =   2050
         Width           =   2700
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
         TabIndex        =   2
         Text            =   "0"
         Top             =   1500
         Width           =   2700
      End
      Begin VB.CommandButton cmdAnularTransaccion 
         Caption         =   "Anular"
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° Transaccion"
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
         Left            =   225
         TabIndex        =   9
         Top             =   530
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo QR"
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
         Left            =   660
         TabIndex        =   8
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Local"
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
         Left            =   435
         TabIndex        =   7
         Top             =   1630
         Width           =   1395
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Left            =   1185
         TabIndex        =   6
         Top             =   2180
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmAmipassAnulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variables Venta
Dim sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String, sSalida As String, sInputJson As String

'Variables reporte
Dim sCodLocalReporte As String, sFecha As String, sRutCliente As String

Dim Jrespuesta As Object, itmx As ListItem

Dim sRespuesta As String, sNumTransaccion As String

Private Sub cmdAnularTransaccion_Click()
    Dim SpMensajeError As String
    
    sNumTransaccion = txtNumTransaction.Text
    sCodLocal = txtCodeLocal.Text
    sMonto = txtMonto.Text
    sCodigoQR = txtQRCode.Text
    sRespuesta = postCancelTransaction(sNumTransaccion, sCodigoQR, sCodLocal, sMonto, sCodigoQR)
    
    'Transforma a Json
    Set Jrespuesta = JSON.parse(sRespuesta)
    
    'Verifica respuesta
    If Jrespuesta.Item("status") <> 200 Then
        SpMensajeError = "Estado: " & Jrespuesta.Item("status") & " " & vbCrLf & _
                         "Respuesta: No se pudo completar transaccion"
        MsgBox SpMensajeError
        Exit Sub
    End If
    
    frmAmipassAnulacion.LvAmiAnulacion.ListItems.Clear
    
    'LLENA GRILLA
    Set itmx = frmAmipassAnulacion.LvAmiAnulacion.ListItems.Add(, , Jrespuesta.Item("response").Item("CodRespuesta"))
    itmx.SubItems(1) = Jrespuesta.Item("response").Item("DesRespuesta")
    itmx.SubItems(2) = Jrespuesta.Item("response").Item("CodAutorizacion")
    itmx.SubItems(3) = Jrespuesta.Item("response").Item("Fecha")
    itmx.SubItems(4) = Jrespuesta.Item("response").Item("Monto")
    
    txtSalida = sRespuesta

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    sCodigoQR = ""
    sMonto = "0"
    sCodLocal = "76449"
    
    txtNumTransaction = ""
    txtQRCode = ""
    txtCodeLocal = sCodLocal
    txtMonto = "1"
End Sub
