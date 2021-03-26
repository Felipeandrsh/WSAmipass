Attribute VB_Name = "Module1"
Dim sTokenAutenticacion As String, sRespuesta As String, sParametros As String, sUrl As String, respuestaJson As String, sToken As String
Dim httpRequest As New WinHttpRequest
Dim iEstado As Integer
Dim oJson As Object

Public Function callAmipassPay(sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String) As String

    sToken = "1348901 "
    sParametros = "?NumeroTransaccion=" & sCodigoQR & "&Monto=" & sMonto & "&CodLocal=" & sCodLocal & "&CodPromocion=" & sPromo
    sUrl = "https://intpay.amipassqa.com/wspay/PayPAP" & sParametros
    'sUrl = "https://pay.amipass.com/wspayTest/PayPAP" & sParametros

    'Crea y envia solicitud
    With httpRequest
        sAuthorization = "Basic " & sToken
        .open "GET", sUrl, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", sAuthorization
        .send
    End With
    
    iEstado = httpRequest.Status
    If iEstado <> 200 Then
        sRespuesta = "Error " & iEstado & " al llamar a la API, " & httpRequest.statusText
        sErrorBody = httpRequest.responseText
    Else
        sRespuesta = httpRequest.responseText
    End If
    
    Set httpRequest = Nothing
    respuestaJson = "{status:'" & iEstado & "',response:'" & sRespuesta & "'}"
    
    callAmipassPay = respuestaJson
    
    'Convertimos cadena a Json
    'Set oJson = json.parse(respuestaJson)
    'Mostramos en String
    'MsgBox (json.toString(oJson))
  
End Function


