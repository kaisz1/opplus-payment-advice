## Introduction ##
This repository contains an extension for the add-ons [Continia Document Output](https://www.continia.com/solutions/continia-document-output/) and [OPplus](https://www.gbedv.de/en/payment-import).
- Document Output is a solution for distributing Business Central/Dynamics NAV documents by e-mail, print, XML (Peppol, XRechnung, etc.).
- OPplus is a solution for handling incoming and outgoing payments

This add-on enables the users to send custom formatted e-mails with the OPplus payment advices as attachment to their customers or vendors.

The code will be unregularly maintained or updated. 

## Remark ##
You can use this code as it is, without any warranty or support by me, [Continia Software A/S](https://www.continia.com "Continia Software"). 
You can use the extension on your own risk. 

**If you find issues in the code, please report these in the issues list here on Github.**

## Current Features ##
1. We add an action to the payment proposal 
2. Find a field value by searching for a caption in the current position (sames procedure like on header fields - but only processed in the current position range)
3. Find a field value by the column heading (same procedure/training) like on default line fields. This enables the user to capture values based on the column that are not in the default line of the current position
4. Define a substitution value. If it was not able to find a value for a field you can define another field, who's value will be set as value for the field with the empty value.
5. Copy a value from the previous line. If it was not able to find a value for a field the system will use the value of this field from the previous line. 

## Documentation ##
There is a basic english documentation available in the repositories Documentation directories:
- [Microsoft Word](https://github.com/document-capture/advanced-line-recognition/blob/master/Documentation/Advanced%20line%20recognition%20-%20ENU.docx)
- [PDF Format](https://github.com/document-capture/advanced-line-recognition/blob/master/Documentation/Advanced%20line%20recognition%20-%20ENU.pdf)


## Available versions ##
The code will be updated unregularly for new versions of Document Capture. 
Going forward we will focus to enhance this module for Business Central al (application language), only. 
For a limited period we will also support the BC14 code to enable partners and customers to use the module on legacy versions of Business Central/NAV/Navision.
