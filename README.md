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

## Currently included in the extension ##
1. New Action at the payment proposals to send advices by Document Output queue (not directly)
![alt text](https://github.com/document-output/opplus-payment-advice/raw/main/media/Payment_Proposals_-_Dynamics_365_Business_Central.png)
2. Included two Document Output templates that can be used to align it with the end customers corporate identity
3. Included a customized report 62040 as 1:1 copy the original OPplus payment advice report with a minor customization that prevents the mails to run into an error

## How to setup and use the extension ##
1. Upload and publish the extension to your Business Central instance
2. Change the default payment advice report on page "Pmt. Export Setup" to 62040
3. Import the two templates to Document Output (one template is for vendor advices, the other for customer advices)
4. Customize the templates according customer requirements
5. Test with customers (test) payment advices and test recipients

## Available versions ##
The code will be updated unregularly for new versions of Document Capture. 
Going forward we will focus to enhance this module for Business Central al (application language), only. 
For a limited period we will also support the BC14 code to enable partners and customers to use the module on legacy versions of Business Central/NAV/Navision.
