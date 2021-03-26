VERSION 5.00
Begin VB.Form FrmMAin 
   ClientHeight    =   7755
   ClientLeft      =   9420
   ClientTop       =   4860
   ClientWidth     =   6855
   LinkTopic       =   "Formulario"
   ScaleHeight     =   7755
   ScaleWidth      =   6855
   Begin VB.TextBox EdPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "e10adc3949ba59abbe56e057f20f883e"
      Top             =   480
      Width           =   3135
   End
   Begin VB.TextBox EdUser 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "pruebas.api@contapyme.com"
      Top             =   120
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   6690
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   6855
   End
   Begin VB.CommandButton ButtonAction 
      Caption         =   "Iniciar sesión y solicitar listado de terceros"
      Height          =   735
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmMAin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'https://www.youtube.com/watch?v=QqbDCA1qo3w
'https://stackoverflow.com/questions/306937/vb6-cast-expression

Private Function GetAuth() As String
    'Se define el objeto que se enviará
    Dim JSONSend As Dictionary
    'Se define el objeto que tendrá los parámetros a enviar
    Dim JParams As Dictionary
    'Se crean los objetos previamente definidos
    Set JSONSend = New Dictionary
    Set JParams = New Dictionary
    'Se adicionan los parámetros necesarios para hacer el GetAuth
    JParams.Item("email") = EdUser.Text
    JParams.Item("password") = EdPassword.Text
    JParams.Item("idmaquina") = "1"
    'Se define el arreglo parámetros generales que se enviaran
    Dim ArrParam(3) As String
    'Se agregan los 4 parámetros al arreglo
    ArrParam(0) = JsonConverter.ConvertToJson(JParams)
    ArrParam(1) = "0"
    ArrParam(2) = "1001"
    ArrParam(3) = "123"
    'Se agrega el arreglo al objeto a enviar
    JSONSend.Item("_parameters") = ArrParam
    'Se define la URL a donde se realizará la petición
    Dim URL As String
    URL = "http://local.insoft.co:9000/datasnap/rest/TBasicoGeneral/""GetAuth""/"
    'Se define el objeto que tendrá la respuesta
    Dim JSONResult As Dictionary
    'Se realiza la petición con los datos previamente definidos
    Set JSONResult = PostRequest(URL, JSONSend)
    'Se defien el resultado como vacio por defecto
    GetAuth = ""
    'Se verifica que el JSON tenga un valor
    If JsonConverter.ConvertToJson(JSONResult) <> "" Then
        'Se verifica que no se existan eventualidades en la petición
        If JSONResult("result")(1)("encabezado")("resultado") <> "true" Then
            MsgBox JSONResult("result")(1)("encabezado")("mensaje")
        Else
            'Se asgina el keyAgente como resultado que se retornará
            GetAuth = JSONResult("result")(1)("respuesta")("datos")("keyagente")
        End If
    End If
End Function

Private Function GetListTerceros(keyAgent As String) As Dictionary
    'Se define y asigna valor a la URL que se solicitará
    Dim URLPost As String
    URLPost = "http://local.insoft.co:9000/datasnap/rest/TCatTerceros/""GetListaTerceros""/"
    'Se define el objeto que se enviará
    Dim JSONSend As Dictionary
    'Se define el objeto que tendrá los parámetros a enviar
    Dim JParams As Dictionary
    'Se crean los objetos previamente definidos
    Set JSONSend = New Dictionary
    Set JParams = New Dictionary
    'Se agrega el objeto con los datos de paginación que se enviaran
    Set JParams.Item("datospagina") = JsonConverter.ParseJson("{""cantidadregistros"":""99"",""pagina"":""1""}")
    'Se agregan el arreglo de campos que se solicitaran en el listado
    Set JParams.Item("camposderetorno") = JsonConverter.ParseJson("[""init"",""ntercero"",""napellido""]")
    'Se define el arreglo parámetros generales que se enviaran
    Dim arrParams(3) As String
    'Se agregan los 4 parámetros al arreglo
    arrParams(0) = JsonConverter.ConvertToJson(JParams)
    arrParams(1) = keyAgent
    arrParams(2) = "1001"
    arrParams(3) = "5555"
    'Se agrega el arreglo al objeto a enviar
    JSONSend.Item("_parameters") = arrParams
    'Se define el objeto que tendrá la respuesta
    Dim ObjResult As Dictionary
    'Se realiza la petición a la URL definida con el JSON previo
    Set ObjResult = PostRequest(URLPost, JSONSend)
    'Se define que el resultado será la respuesta que se encuentra dentro del JSON entregado
    Set GetListTerceros = ObjResult("result")(1)("respuesta")
End Function

Private Function PostRequest(URL As String, Data As Dictionary) As Dictionary
    'Se define el Objeto para realiza la petición POST
    Dim winH As WinHttp.WinHttpRequest
    'Se crea el objeto previamente definido
    Set winH = New WinHttp.WinHttpRequest
    'Se realiza la apertura de la conexión a la URL con verbo POST
    winH.Open "post", URL
    'Se envía el string con el JSON de datos
    winH.Send JsonConverter.ConvertToJson(Data)
    'Se retorna el diccionario
    Set PostRequest = JsonConverter.ParseJson(winH.ResponseText)
End Function



Private Sub ButtonAction_Click()
    'Se eliminan los elementos del listBox
    List1.Clear
    'Se define la variable que tendrá el keyAgente
    Dim sKey As String
    'Se solicita el KeyAgente
    sKey = GetAuth()
    'Se verifica que el keyAgente tenga un valor
    If sKey <> "" Then
        'Se difine la variable que tendra la lista de terceros
        Dim arrTerceros As Dictionary
        'Se define el string que tendra la información de cada tercero
        Dim StrTem As String
        'Se solicita y almacena el listado de terceros
        Set arrTerceros = GetListTerceros(sKey)
        'Se define una variable i para recorrer el listado
        Dim i As Integer
        'Se recorre el listado de terceros retornado
        For i = 1 To arrTerceros("datos").Count - 1
            'Se obtiene los datos de cada registro
            StrTem = arrTerceros("datos")(i)("init")
            StrTem = StrTem & " " & arrTerceros("datos")(i)("ntercero")
            StrTem = StrTem & " " & arrTerceros("datos")(i)("napellido")
            'Se agrega cada registro como item al ListBox
            List1.AddItem (StrTem)
        Next
    End If
End Sub
