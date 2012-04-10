<%
' Classic ASP CAS SAML - classicAspCasSaml.asp
' (c) 2012 Keene State College

' refer to README.txt for details


' Declare Option Explicit
' from whatever page includes this file


Class classicAspCasSaml
   Private ticket            
   Private requestID         
   Private issueInstant      

   Public SAMLValidateURL   
   Public SAMLResponseDoc



   Public Sub Class_Initialize()
     ' setup variables for the saml validation soap xml

     Dim mTS, dNow, d1970
     d1970 = CDate("1970-01-01")
     dNow = Now()
     mTS = DateDiff("s", d1970, dNow)

     ' define private attributes used by SAMLValidateString
     requestID = "_" & Request.ServerVariables("REMOTE_HOST") & "." & mTS
     ticket = Request.QueryString("ticket")
     issueInstant = Year(dNow) & "-" & Right("00"& Month(dNow), 2) & "-" & Right("00"& Day(dNow), 2) & "T" & _
       Right("00" & Hour(dNow), 2) & ":" & Right("00" & Minute(dNow), 2) & ":" & Right("00"& Second(dNow), 2) & "Z"

   End Sub

   ' sticking w/string vs XML DOM for simplicity
   Public Property Get SAMLValidateString()
      Dim str
      str = "<?xml version='1.0'?>" & vbCrLf & _
        "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<SOAP-ENV:Header/>" & _
        "<SOAP-ENV:Body>" & _
        "<samlp:Request " & _
        "xmlns:samlp=""urn:oasis:names:tc:SAML:1.0:protocol"" " & _
        "MajorVersion=""1"" " & _
        "MinorVersion=""1"" " & _
        "RequestID=""" & requestID & """ " &_
        "IssueInstant=""" & issueInstant & """>" & _
        "<samlp:AssertionArtifact>" & ticket & "</samlp:AssertionArtifact>" & _
        "</samlp:Request>" & _
        "</SOAP-ENV:Body>" & _
        "</SOAP-ENV:Envelope>"
      SAMLValidateString = str

   End Property



   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' submit a SOAP/SAML request to CAS's samlValidate
   Public Sub ValidateTicket()

     Dim actionURL, objReq, xmlDoc

     ' there should be ?TARGET= in the url.
     ' left of that question mark becomes the action url
     actionURL = ""
     If InStr(SAMLValidateURL, "?") > 0  Then
         actionURL = Left(SAMLValidateURL, InStr(SAMLValidateURL, "?") - 1)
     End If

      ' build xml object, make request, set headers, send xml post
     Set objReq = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
     objReq.open "POST", SAMLValidateURL, False
     objReq.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8"
     objReq.setRequestHeader "action", actionUrl
     objReq.setRequestHeader "SOAPAction", actionUrl
     objReq.setRequestHeader "User-Agent", "classicAspCasSaml 0.1"
     objReq.send SAMLValidateString


     ' the request has been sent.
     ' now create a new xmldom object
     Set xmlDoc = Server.CreateObject("Microsoft.XMLDOM")
     xmlDoc.async = False
     xmlDoc.setProperty "ServerHTTPRequest", true

     ' load the samlValidate response into XMLDOM object
     If Not(xmlDoc.LoadXML(objReq.responseText)) Then
       Set objReq = Nothing
       Set postSamlValidate = Nothing
     End If
     Set objReq = Nothing

     ' setup for accessing document's contents using xpath
     xmlDoc.setProperty "SelectionLanguage", "XPath"
     xmlDoc.setProperty "SelectionNamespaces", "xmlns:s='http://schemas.xmlsoap.org/soap/envelope/' " & _
                                            "xmlns:nsR='urn:oasis:names:tc:SAML:1.0:protocol' " & _
                                            "xmlns:nsA='urn:oasis:names:tc:SAML:1.0:assertion'"

    ' Response.write xmlDoc.xml
    ' Response.end

     ' SAMLResponseDoc is public property of classicAspCasSaml class
     Set SAMLResponseDoc = xmlDoc
     'Set xmlDoc = Nothing

   End Sub


  ' evaluate samlValidate response and populate a dictionary based on what exists
  ' The object should look something like this:
  ' {
  '   "status"         => [0 => "Y" | "N"],
  '   "debug"          => [0 => "text description of status"],
  '   "NameIdentifier" => [0 => "username returned from CAS"],
  '   "otherAttrs"     => [0 => "otherAttrVal1",
  '                        1 => "otherAttrVal2"]
  ' }

  ' a little excessive to give everything a dictionary but
  ' its flexible and avoids lots of ReDim calls.
  ' probably a better way

  Public Function parseSamlValidate(attrList)

   Dim objDict, objTmp
   Set objDict = Server.CreateObject("Scripting.Dictionary")


   If SAMLResponseDoc Is Nothing Then
     Set objTmp = Server.CreateObject("Scripting.Dictionary")
     objTmp.Add 0, "N"
     objDict.Add "status", objTmp

     Set objTmp = Server.CreateObject("Scripting.Dictionary")
     objTmp.Add 0, "No SAML XML"
     objDict.Add "debug", objTmp
     parseSamlValidate = objDict
   Else

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     ' check the XML documentElement for response status values using
     ' 'With' to limit SAMLResponseDoc.documentElement evaluation

     With SAMLResponseDoc.documentElement
       Dim objNode
       Dim sStatusCode, sStatusMessage, sDebug


       ' status code - interpret it as Y or N
       Set objNode = .selectSingleNode("/s:Envelope/s:Body/nsR:Response/nsR:Status/nsR:StatusCode")
       If Not(objNode Is Nothing) Then
         sStatusCode = objNode.getAttribute("Value")
         If sStatusCode = "samlp:Success" Then
           sStatusCode = "Y"
         End If
       End If

       ' Try to get some good debugging information if status wasn't valid
       sDebug = ""
       If Not(sStatusCode = "Y") Then
         sStatusCode = "N"
         Set objNode = .selectSingleNode("/s:Envelope/s:Body/nsR:Response/nsR:Status/nsR:StatusMessage")
         If Not(objNode Is Nothing) Then
           sDebug = objNode.Text
         End If
       ElseIf sStatusCode = "Y" Then
         sDebug = "Success"
       Else
         sDebug = "Status check failed."
       End If

       Set objTmp = Server.CreateObject("Scripting.Dictionary")
       objTmp.Add 0, sStatusCode
       objDict.Add "status", objTmp


       Set objTmp = Server.CreateObject("Scripting.Dictionary")
       objTmp.Add 0, sDebug
       objDict.Add "debug", objTmp


       ' Get NameIdentifier
       Set objNode = .selectSingleNode("/s:Envelope/s:Body/nsR:Response/nsA:Assertion/nsA:AuthenticationStatement/nsA:Subject/nsA:NameIdentifier")
       If Not(objNode Is Nothing) Then
         Set objTmp = Server.CreateObject("Scripting.Dictionary")
         objTmp.Add 0, objNode.Text
         objDict.Add "NameIdentifier", objTmp
       End If

       Dim i, attrName, oAttrValue

       ' Walk attrList for elements with matching AttributeName attributes
       ' add to dictionary when appropriate
       For Each attrName in attrList
         Set objNode = .selectSingleNode("/s:Envelope/s:Body/nsR:Response/nsA:Assertion" & _
                     "/nsA:AttributeStatement/nsA:Attribute[@AttributeName='" & attrName & "']")
         If Not(objNode Is Nothing) Then
           Set oAttrValue = objNode.SelectNodes("descendant::nsA:AttributeValue")
           If Not(oAttrValue Is Nothing) Then
              Set objTmp = Server.CreateObject("Scripting.Dictionary")
              For i = 0 to oAttrValue.length - 1
                 objTmp.Add i, oAttrValue.Item(i).Text
              Next
                objDict.add attrName, objTmp
              End If
           End If
       Next

     End With


     ' clean up objects
     Set objNode = Nothing
     Set objTmp = Nothing

   End If

   Set parseSamlValidate = objDict
   Set objDict = Nothing
  End Function

End Class


%>