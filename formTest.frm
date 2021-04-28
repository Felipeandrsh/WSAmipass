VERSION 5.00
Begin VB.Form formTest 
   Caption         =   "Form1"
   ClientHeight    =   9750
   ClientLeft      =   1500
   ClientTop       =   4305
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9750
   ScaleWidth      =   20250
   Begin VB.Frame Frame6 
      Caption         =   "Anulacion"
      Height          =   3735
      Left            =   10980
      TabIndex        =   31
      Top             =   435
      Width           =   4860
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
         Height          =   495
         Left            =   3210
         TabIndex        =   41
         Top             =   2865
         Width           =   1500
      End
      Begin VB.TextBox txtCodeLocalCancel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1860
         TabIndex        =   37
         Top             =   1500
         Width           =   2670
      End
      Begin VB.TextBox txtAmountCancel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1860
         TabIndex        =   36
         Top             =   2055
         Width           =   2685
      End
      Begin VB.TextBox txtQRCodeCancel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1860
         TabIndex        =   35
         Top             =   945
         Width           =   2670
      End
      Begin VB.TextBox txtNmTransactionCancel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1860
         TabIndex        =   32
         Top             =   375
         Width           =   2670
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         Height          =   495
         Left            =   615
         TabIndex        =   40
         Top             =   2085
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Height          =   495
         Left            =   15
         TabIndex        =   39
         Top             =   1605
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
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
         Height          =   495
         Left            =   600
         TabIndex        =   38
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Height          =   495
         Left            =   75
         TabIndex        =   33
         Top             =   510
         Width           =   1740
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Cangear Amipesos"
      Height          =   3360
      Left            =   5625
      TabIndex        =   25
      Top             =   4275
      Width           =   5220
      Begin VB.CommandButton cmdCangearAmipesos 
         Caption         =   "Cangear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3510
         TabIndex        =   30
         Top             =   2655
         Width           =   1500
      End
      Begin VB.TextBox txtCodeLocalCange 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1500
         TabIndex        =   29
         Top             =   1185
         Width           =   2595
      End
      Begin VB.TextBox txtRutCustomerCange 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1500
         TabIndex        =   28
         Top             =   465
         Width           =   2535
      End
      Begin VB.Label Label5 
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
         Height          =   495
         Left            =   360
         TabIndex        =   27
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Rut Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   285
         TabIndex        =   26
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Verifica Transaccion"
      Height          =   3735
      Left            =   5655
      TabIndex        =   19
      Top             =   435
      Width           =   5175
      Begin VB.CommandButton cmdVerificaTransaccion 
         Caption         =   "Verificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3465
         TabIndex        =   24
         Top             =   2910
         Width           =   1500
      End
      Begin VB.TextBox txtCodLocalTransaction 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1935
         TabIndex        =   21
         Top             =   1365
         Width           =   2610
      End
      Begin VB.TextBox txtNumTransaction 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1935
         TabIndex        =   20
         Top             =   645
         Width           =   2610
      End
      Begin VB.Label Label3 
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
         Height          =   495
         Left            =   495
         TabIndex        =   23
         Top             =   1410
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Left            =   255
         TabIndex        =   22
         Top             =   705
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Respuesta Json"
      Height          =   7380
      Left            =   15975
      TabIndex        =   17
      Top             =   270
      Width           =   5505
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
         Height          =   6825
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   5190
      End
   End
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
      Height          =   495
      Left            =   19695
      TabIndex        =   16
      Top             =   7770
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      Caption         =   "Post Venta"
      Height          =   3750
      Left            =   480
      TabIndex        =   3
      Top             =   420
      Width           =   5040
      Begin VB.CommandButton cmdTest 
         Caption         =   "Pagar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3180
         TabIndex        =   15
         Top             =   2880
         Width           =   1500
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1665
         TabIndex        =   14
         Top             =   2070
         Width           =   2385
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1650
         TabIndex        =   13
         Top             =   1500
         Width           =   2385
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1665
         TabIndex        =   12
         Top             =   945
         Width           =   2370
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   11
         Top             =   375
         Width           =   2370
      End
      Begin VB.Label Label1 
         Caption         =   "Promo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   780
         TabIndex        =   10
         Top             =   2145
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Index           =   3
         Left            =   345
         TabIndex        =   9
         Top             =   1635
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Index           =   1
         Left            =   795
         TabIndex        =   8
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Index           =   0
         Left            =   330
         TabIndex        =   7
         Top             =   555
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reporte de Transacciones (Diario)"
      Height          =   3360
      Left            =   345
      TabIndex        =   0
      Top             =   4305
      Width           =   5160
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
         Height          =   495
         Left            =   3345
         TabIndex        =   6
         Top             =   2565
         Width           =   1500
      End
      Begin VB.TextBox txtDateReport 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1740
         TabIndex        =   5
         Top             =   795
         Width           =   2310
      End
      Begin VB.TextBox txtCodeLocalReport 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1755
         TabIndex        =   4
         Text            =   " "
         Top             =   1545
         Width           =   2310
      End
      Begin VB.Label lblFecha 
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
         Height          =   495
         Left            =   750
         TabIndex        =   2
         Top             =   915
         Width           =   1215
      End
      Begin VB.Label lblCodLocal 
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
         Height          =   495
         Left            =   420
         TabIndex        =   1
         Top             =   1695
         Width           =   1215
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   10035
      TabIndex        =   34
      Top             =   4365
      Width           =   1215
   End
End
Attribute VB_Name = "formTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variables Venta
Dim sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String, sSalida As String, sInputJson As String

'Variables reporte
Dim sCodLocalReporte As String, sFecha As String, sRutCliente As String

Dim sRespuesta As String, sNumTransaccion As String

Private Sub Form_Load()
   
    sCodigoQR = ""
    sMonto = "0"
    sCodLocal = "76449"
    sPromo = "0"
    sFecha = "2021-03-29"
    
    Text1 = sCodigoQR
    Text2 = sMonto
    Text3 = sCodLocal
    Text4 = sPromo
    txtDateReport = sFecha
    txtCodeLocalReport = sCodLocal
    txtCodLocalTransaction = sCodLocal
    txtNumTransaction = ""
    txtRutCustomerCange = "NN33028"
    txtCodeLocalCange = sCodLocal
    
    txtNmTransactionCancel = ""
    txtQRCodeCancel = ""
    txtCodeLocalCancel = sCodLocal
    txtAmountCancel = ""
    
End Sub

Private Sub cmdTest_Click()
    
    sCodigoQR = Text1.Text
    sMonto = Text2.Text
    sCodLocal = Text3.Text
    sPromo = Text4.Text
    
    sRespuesta = postCreateTransaction(sCodigoQR, sMonto, sCodLocal, sPromo)
    
    txtSalida = sRespuesta
    sCodigoQR = ""
    sMonto = "0"
    

End Sub

Private Sub cmdReporte_Click()
    
    sCodLocalReporte = txtCodeLocalReport.Text
    sFecha = txtDateReport.Text
    sRespuesta = getTransactionReports(sFecha, sCodLocalReporte)
    txtSalida = sRespuesta
    
    
End Sub

Private Sub cmdVerificaTransaccion_Click()

    sNumTransaccion = txtNumTransaction.Text
    sCodLocal = txtCodLocalTransaction.Text
    sRespuesta = getTransactionData(sNumTransaccion, sCodLocal)
    
    txtSalida = sRespuesta
    
End Sub

Private Sub cmdCangearAmipesos_Click()
    
    sRutCliente = txtRutCustomerCange.Text
    sCodLocal = txtCodeLocalCange.Text
    sRespuesta = getCustomerCange(sRutCliente, sCodLocal)
    txtSalida = sRespuesta

End Sub

Private Sub cmdAnularTransaccion_Click()
    
    sNumTransaccion = txtNmTransactionCancel.Text
    sCodLocal = txtCodeLocalCancel.Text
    sMonto = txtAmountCancel.Text
    sCodigoQR = txtQRCodeCancel.Text
    sRespuesta = postCancelTransaction(sNumTransaccion, sCodigoQR, sCodLocal, sMonto, sCodigoQR)
    txtSalida = sRespuesta

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub VScroll1_Change()

End Sub

