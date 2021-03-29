Attribute VB_Name = "Module1"
Dim sRespuesta As String, sParametros As String, sUrl As String, respuestaJson As String, sToken As String, sMsg As String
Dim httpRequest As New WinHttpRequest
Dim iEstado As Integer
Dim resJson As Object


Dim arrSplitStrings() As String

Public Function callAmipassPay(sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String) As String

On Error GoTo ErrHandler

    sToken = "1348901 "
    sParametros = "?NumeroTransaccion=" & sCodigoQR & "&Monto=" & sMonto & "&CodLocal=" & sCodLocal & "&CodPromocion=" & sPromo
    sUrl = "https://intpay.amipassqa.com/wspay/PayPAP" & sParametros

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
        sRespuesta = "{'Error " & iEstado & " al llamar a la API, " & httpRequest.statusText & "'}"
        sErrorBody = httpRequest.responseText
    Else
        sRespuesta = httpRequest.responseText
        
        'Quita \ y quita comillas de indices
        sRespuesta = Replace(sRespuesta, Chr(92), Chr(32))
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "CodRespuesta " & Chr(34) & "", "CodRespuesta")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "DesRespuesta " & Chr(34) & "", "DesRespuesta")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "CodAutorizacion " & Chr(34) & "", "CodAutorizacion")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "Fecha " & Chr(34) & "", "Fecha")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "Monto " & Chr(34) & "", "Monto")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "Saldo " & Chr(34) & "", "Saldo")
              
        'Cambia comillas dobles por comillas simples
        sRespuesta = Replace(sRespuesta, Chr(34), Chr(39))
        
        'Quita primera y ultima comilla
        sRespuesta = Mid(sRespuesta, 2, Len(sRespuesta) - 2)
        'MsgBox sRespuesta
      
    End If
    
    respuestaJson = "{status:'" & iEstado & "',response:" & sRespuesta & "}"
    Set httpRequest = Nothing
    callAmipassPay = respuestaJson
  
Exit Function

ErrHandler:
    sMsg = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
    MsgBox sRespuesta

End Function
