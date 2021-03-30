Attribute VB_Name = "Module1"
Dim sRespuesta As String, sParametros As String, sUrl As String, respuestaJson As String, sToken As String, sMsError As String
Dim httpRequest As New WinHttpRequest
Dim iEstado As Integer

Public Function callAmipass(sParametros As String)
   

On Error GoTo ErrHandler
    
    sToken = "1348901"
    sUrl = "https://intpay.amipassqa.com/wspay/" & sParametros
    
    With httpRequest
        sAuthorization = "Basic " & sToken
        .open "GET", sUrl, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", sAuthorization
        .send
    End With
    
Exit Function

ErrHandler:
sMsError = "Error #" & Err.Number & ": '" & Err.Description & "' from '" & Err.Source & "'"
MsgBox sMsError

End Function

Public Function postCreateTransaction(sCodigoQR As String, sMonto As String, sCodLocal As String, sPromo As String) As String

    sParametros = "PayPAP?NumeroTransaccion=" & sCodigoQR & "&Monto=" & sMonto & "&CodLocal=" & sCodLocal & "&CodPromocion=" & sPromo
    
   'Realiza llamada a API
    callAmipass (sParametros)
    
    iEstado = httpRequest.Status
    If iEstado <> 200 Then
        sRespuesta = "{'Error " & iEstado & " al llamar a la API, " & httpRequest.statusText & "'}"
        sErrorBody = httpRequest.responseText
    ElseIf httpRequest.responseText = "" & Chr(34) & "[]" & Chr(34) & "" Then
        iEstado = 404
        sRespuesta = "{'No se encontro transaccion, " & httpRequest.statusText & "'}"
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
    postCreateTransaction = respuestaJson

End Function

Public Function getTransactionReports(sFecha As String, sCodLocalReporte As String) As String

    sParametros = "TX?Fecha=" & sFecha & "&CodLocal=" & sCodLocalReporte
  
    'Realiza llamada a API
    callAmipass (sParametros)
    
    iEstado = httpRequest.Status
    If iEstado <> 200 Then
        sRespuesta = "{'Error " & iEstado & " al llamar a la API, " & httpRequest.statusText & "'}"
        sErrorBody = httpRequest.responseText
    ElseIf httpRequest.responseText = "" & Chr(34) & "[]" & Chr(34) & "" Then
        iEstado = 404
        sRespuesta = "{'No se encontro transaccion, " & httpRequest.statusText & "'}"
    Else
        sRespuesta = httpRequest.responseText
        
        'Quita \ y quita comillas de indices
        sRespuesta = Replace(sRespuesta, Chr(92), Chr(32))
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "sTurno " & Chr(34) & "", "sTurno")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "sTipoTransaccion " & Chr(34) & "", "sTipoTransaccion")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "dTransaccion " & Chr(34) & "", "dTransaccion")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "idTransaccion " & Chr(34) & "", "idTransaccion")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "nMonto " & Chr(34) & "", "nMonto")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "sRutCliente " & Chr(34) & "", "sRutCliente")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "sRutCompleto " & Chr(34) & "", "sRutCompleto")
              
        'Cambia comillas dobles por comillas simples
        sRespuesta = Replace(sRespuesta, Chr(34), Chr(39))
        
        'Quita primera y ultima comilla
        sRespuesta = Mid(sRespuesta, 2, Len(sRespuesta) - 2)
        'MsgBox sRespuesta
      
    End If
    
    respuestaJson = "{status:'" & iEstado & "',response:" & sRespuesta & "}"
    Set httpRequest = Nothing
    getTransactionReports = respuestaJson
  
End Function

Public Function getTransactionData(sNumTransaccion As String, CodLocal As String) As String

    sParametros = "VTX?nTransaccion=" & sNumTransaccion & "&CodLocal=" & CodLocal
    
    'Realiza llamada a API
    callAmipass (sParametros)
    
    iEstado = httpRequest.Status
    If iEstado <> 200 Then
        sRespuesta = "{'Error " & iEstado & " al llamar a la API, " & httpRequest.statusText & "'}"
        sErrorBody = httpRequest.responseText
    'Si respuesta es []
    ElseIf httpRequest.responseText = "" & Chr(34) & "[]" & Chr(34) & "" Then
        'MsgBox "No se encontro nada"
        iEstado = 404
        sRespuesta = "{'No se encontro transaccion, " & httpRequest.statusText & "'}"
    Else
        sRespuesta = httpRequest.responseText
        
        'Quita \ y quita comillas de indices
        sRespuesta = Replace(sRespuesta, Chr(92), Chr(32))
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "sTurno " & Chr(34) & "", "sTurno")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "sTipoTransaccion " & Chr(34) & "", "sTipoTransaccion")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "dTransaccion " & Chr(34) & "", "dTransaccion")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "idTransaccion " & Chr(34) & "", "idTransaccion")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "nMonto " & Chr(34) & "", "nMonto")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "sRutCliente " & Chr(34) & "", "sRutCliente")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "sRutCompleto " & Chr(34) & "", "sRutCompleto")
              
        'Cambia comillas dobles por comillas simples
        sRespuesta = Replace(sRespuesta, Chr(34), Chr(39))
        
        'Quita primera y ultima comilla
        sRespuesta = Mid(sRespuesta, 2, Len(sRespuesta) - 2)
        'MsgBox sRespuesta
      
    End If
    
    respuestaJson = "{status:'" & iEstado & "',response:" & sRespuesta & "}"
    Set httpRequest = Nothing
    getTransactionData = respuestaJson
    
End Function

Public Function getCustomerCange(sRutCliente As String, CodLocal As String) As String

    sParametros = "CA?RutCliente=" & sRutCliente & "&CodLocal=" & CodLocal
    
    'Realiza llamada a API
    callAmipass (sParametros)
    
    iEstado = httpRequest.Status
    If iEstado <> 200 Then
        sRespuesta = "{'Error " & iEstado & " al llamar a la API, " & httpRequest.statusText & "'}"
        sErrorBody = httpRequest.responseText
    'Si respuesta es []
    ElseIf httpRequest.responseText = "" & Chr(34) & "[]" & Chr(34) & "" Then
        'MsgBox "No se encontro nada"
        iEstado = 404
        sRespuesta = "{'No se encontro transaccion, " & httpRequest.statusText & "'}"
    Else
        sRespuesta = httpRequest.responseText
        
        'Quita \ y quita comillas de indices
        sRespuesta = Replace(sRespuesta, Chr(92), Chr(32))
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "bCobrado " & Chr(34) & "", "bCobrado")
        sRespuesta = Replace(sRespuesta, "" & Chr(34) & "sMensaje " & Chr(34) & "", "sMensaje")
              
        'Cambia comillas dobles por comillas simples
        sRespuesta = Replace(sRespuesta, Chr(34), Chr(39))
        
        'Quita primera y ultima comilla
        sRespuesta = Mid(sRespuesta, 2, Len(sRespuesta) - 2)
        'MsgBox sRespuesta
      
    End If
    
    respuestaJson = "{status:'" & iEstado & "',response:" & sRespuesta & "}"
    Set httpRequest = Nothing
    getCustomerCange = respuestaJson
    
End Function

Public Function postCancelTransaction(sNumTransaccion As String, sCodigoQRForm As String, sCodLocalForm As String, sMonto As String, sCodigoQR As String) As String

    sParametros = "ANPAP?nTransaccion=" & sNumTransaccion & "&sCodigoQR=" & sCodigoQRForm & "&sCodLocal=" & sCodLocalForm & "&nMonto=" & sMonto
    
    'Realiza llamada a API
    callAmipass (sParametros)
    
    iEstado = httpRequest.Status
    If iEstado <> 200 Then
        sRespuesta = "{'Error " & iEstado & " al llamar a la API, " & httpRequest.statusText & "'}"
        sErrorBody = httpRequest.responseText
    'Si respuesta es []
    ElseIf httpRequest.responseText = "" & Chr(34) & "[]" & Chr(34) & "" Then
        'MsgBox "No se encontro nada"
        iEstado = 404
        sRespuesta = "{'No se encontro transaccion, " & httpRequest.statusText & "'}"
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
    postCancelTransaction = respuestaJson
    
End Function

