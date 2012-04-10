<% 
Option Explicit 

' (c) 2012 Keene State College
' refer to README.txt for details

%>

<!--#include file="classicAspCasSaml.asp"-->

<%
' define URL for Service (client application) and CAS 
Dim serviceURL, casURL

' url to go to after authentication
serviceURL = "https://myapp.domain.com/path/demo.asp"

' the url of CAS
casURL = "https://auth.domain.edu/cas/"


' no ticket? redirect to CAS login
If Request.QueryString("ticket") = "" Then
  Response.Redirect casURL + "login/?service=" + serviceURL
Else
  ' build soap request for ticket validation, send samlValidation, parse the response

  Dim oCASSaml, validateURL, oSAMLDoc
  Set oCASSaml = new classicAspCasSaml
  validateURL = casURL & "samlValidate" & "?TARGET=" & serviceURL

  oCASSaml.SAMLValidateURL = validateURL
  oCASSaml.ValidateTicket()

  ' oCASSaml.ValidateTicket() attempts to validate the ticket. 
  ' The remainder of this code demonstrates SAML response parsing and use of 
  ' a helper function to simplify access to the response by using a dictionary 
  
  response.ContentType = "text/plain"
  Response.write "--------------------------------------------------------------" & vbCrLf
  Response.write "Retrieve one item from CAS's saml response xml: " & vbCrLf
  Dim oTmpNode
  Set oTmpNode = oCASSaml.SAMLResponseDoc.selectSingleNode("/s:Envelope/s:Body/nsR:Response/nsA:Assertion/nsA:AuthenticationStatement/nsA:Subject/nsA:NameIdentifier")
  If Not(oTmpNode Is Nothing) Then
    Response.write "NameIdentifier: " & oTmpNode.Text & vbCrLf
    Set oTmpNode = Nothing
  End If

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  ' For easier access to SAML elements, parseSamlValidate accepts a variant 
  ' of attribute identifiers and returns a dictionary of items from samlValidate. 
  ' In addition to specific attributes,  this function should return 
  ' status, debug, and NameIdentifier depending on success

  Dim oSAMLDictionary, myAttributes(3)
  myAttributes(0) = "mail"
  myAttributes(1) = "memberOf"
  myAttributes(2) = "employeeID"
  myAttributes(3) = "displayName"
  Set oSAMLDictionary = oCASSaml.parseSamlValidate(myAttributes)


  ' access one item from the parsed response object
  Response.write "-------------------------------------------------------------" & vbCrLf
  Response.write "Retrieve one item from the parsed response object: " & vbCrLf
  Response.write "Status: " & oSamlDictionary.Item("status").Item(0) & vbCrLF


  Response.write "-------------------------------------------------------------" & vbCrLf
  Response.write "walk parsed response..." & vbCrLf
  Dim attrKey, valKey
  For Each attrKey in oSAMLDictionary.Keys
      Response.write attrKey & ":" & vbCrLf
      With oSAMLDictionary.Item(attrKey)
        For each valKey in oSAMLDictionary.Item(attrKey)
           Response.write vbTab & .Item(valKey) & vbCrLf
        Next
      End With
  Next

  Set oSAMLDictionary = Nothing
  Set oCASSaml = Nothing
End If





%>