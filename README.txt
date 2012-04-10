classicAspCasSaml is an unofficial CAS client to enable CAS authentication for 
classic ASP (VB) applications that require attribute release. Attribute 
release is made possible by posting a SAML document (over SOAP over HTTP) to 
CAS's samlValidate for ticket validation. classicAspCasSaml was developed by 
Keene State College.    


QUICKSTART
-------------------------------------------------------------------------------
* Put this directory on a test machine in a Classic ASP (VB) environment in a
space where IIS serves web pages (e.g., C:\inetpub\wwwtest\classicAspCasSaml). 
* Update the serviceURL and casURL variable in demo.asp to reflect your 
environment. 
* Open your CAS server's services manager.
* Add the URL of the demo.asp page as an authorized service. 
* Select attributes for CAS to release.
* Save changes.
* Log out of CAS.
* Open the URL of the demo.asp through your browser. 

Following a successful login and ticket validation, the demo.asp page will 
display information that was returned by CAS. 



 
TROUBLESHOOTING
-------------------------------------------------------------------------------
Tail your CAS log file and review debug.asp for more info.

 

LICENSE
-------------------------------------------------------------------------------
See LICENSE.txt for details.


HISTORY
-------------------------------------------------------------------------------
2012-03-20

version 0.1 - alpha 


See ChangeLog.txt for details.