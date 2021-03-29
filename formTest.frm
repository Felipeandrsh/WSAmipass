VERSION 5.00
Begin VB.Form formTest 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
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
      Left            =   2085
      TabIndex        =   9
      Top             =   2805
      Width           =   1830
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
      Left            =   2070
      TabIndex        =   8
      Top             =   2100
      Width           =   1800
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
      Left            =   2055
      TabIndex        =   7
      Top             =   1500
      Width           =   1800
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
      Left            =   2055
      TabIndex        =   6
      Top             =   870
      Width           =   1785
   End
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
      Left            =   2025
      TabIndex        =   1
      Top             =   3660
      Width           =   1665
   End
   Begin VB.TextBox txtSalida 
      Height          =   5595
      Left            =   4605
      TabIndex        =   0
      Top             =   45
      Width           =   9015
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
      Left            =   900
      TabIndex        =   5
      Top             =   2190
      Width           =   1215
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
      Left            =   1125
      TabIndex        =   4
      Top             =   2820
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
      Left            =   1200
      TabIndex        =   3
      Top             =   1530
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
      Left            =   765
      TabIndex        =   2
      Top             =   930
      Width           =   1215
   End
End
Attribute VB_Name = "formTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String
Dim a As String
Dim sInputJson As String
Dim jRespueta As Object
Dim sSalida As String

Private Sub Form_Load()
   
    sCodigoQR = ""
    sMonto = "0"
    sCodLocal = "76449"
    sPromo = "0"
    
    Text1 = sCodigoQR
    Text2 = sMonto
    Text3 = sCodLocal
    Text4 = sPromo
    
End Sub

Private Sub cmdTest_Click()
    
    sCodigoQR = Text1.Text
    sMonto = Text2.Text
    sCodLocal = Text3.Text
    sPromo = Text4.Text
    
    sSalida = callAmipassPay(sCodigoQR, sMonto, sCodLocal, sPromo)
    
    'Json to String
    txtSalida = sSalida

End Sub

