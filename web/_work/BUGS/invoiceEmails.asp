<%
strHead = "<table width=""600"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center""><tr><td><img src=""http://system.newzapp.co.uk/images/nzEmailHeader.gif"" width=""600"" height=""77"" alt=""NewZapp Email Marketing""></td></tr></table><table width=""600"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center""><tr><td style=""font-family:Arial, Helvetica, sans-serif; color:#000000; font-size:11px; padding:10px;"">"

strFoot = "</td></tr><tr><td><img src=""http://system.newzapp.co.uk/images/nztrial/NewZappFooter.gif"" width=""600"" height=""100""></td></tr><tr><td style=""font-size:11px; font-family:Arial, Helvetica, sans-serif; padding:10px;""> <strong>New</strong>Zapp is a product of DestiNet Ltd. <br><br>DestiNet Limited is Registered in England with Company No. 3679291. Registered   Address: Bradley House, 7 Park Five Business Centre, Harrier Way, Exeter, EX2   7HU. This message, and any associated files, is intended only for the use of the   individual or entity to which it is addressed and may contain information that   is confidential, subject to copyright or constitutes a commercial secret. If you   are not the intended recipient you are hereby notified that any dissemination,   copying or distribution of this message, or files associated with the message,   is strictly prohibited. If you are not the intended recipient, please notify us immediately.</td></tr></table>"

strInvoice1="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Order</span><br><br>Thank you for choosing <strong>New</strong>Zapp for your email marketing.<br>Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.<p>If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544<br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts</p>"

strInvoice2="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Order</span><br><br>Thank you for choosing <strong>New</strong>Zapp for your email marketing.  We haven't received payment for your recent order as per the attached invoice.<br><br>Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.<br>You can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card.<p>If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544<br><br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts</p>"

strInvoice3="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Order</span><br><br>Unpaid order as attached.<br><br><strong>New</strong>Zapp Accounts"

strInvoice4="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>Your New</strong>Zapp Licence has Expired </span><br><br>Your  annual <strong>New</strong>Zapp  licence has expired and  we have not received payment for your renewal.  In order to continue to use the <strong>New</strong>Zapp email marketing services you will need to pay your renewal invoice attached. If you do not renew your licence within 30 days ALL your data will be removed from our systems.<br><br>Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.<br>You can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card.  When you log in to your <strong>New</strong>Zapp account the only option available will be to view and pay your unpaid invoices. <p>If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544<br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br><br>This invoice is due for payment immediately.  In order for your new licence to be added to your account you must pay this invoice. <br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts</p>"

strInvoice5="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Licence due for Renewal </span><br><br>Our records show that your next annual <strong>New</strong>Zapp payment is due shortly.<br /><br />Payment of your annual licence  is due for payment by return.  Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.  If you normally pay by credit card this invoice will be charged to your card. You can pay this invoice using a different credit card by logging in to your account.<br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br><br>Please ensure your payment reaches us by return to prevent any interruption to your <strong>New</strong>Zapp email marketing services.<br /><br />You can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card. If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544 <br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts"

strInvoice6="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Licence due for Renewal </span><br>    <br>  Our records show that your next quarterly <strong>New</strong>Zapp payment is due shortly.<br /><br />Quarterly payment of your annual licence  is due for payment by return.  Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.  If you normally pay by credit card this invoice will be charged to your card. You can pay this invoice using a different credit card by logging in to your account.<br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br /><br />Please ensure your payment reaches us by return to prevent any interruption to your <strong>New</strong>Zapp email marketing services.<br /><br />You can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card. If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544<br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts"

strInvoice7="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Licence due for Renewal </span><br><br>Our records show that your next monthly <strong>New</strong>Zapp payment is due shortly.<br /><br />Monthly payment of your annual licence  is due for payment by return.  Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.  If you normally pay by credit card this invoice will be charged to your card. You can pay this invoice using a different credit card by logging in to your account.<br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br><br>Please ensure your payment reaches us by return to prevent any interruption to your NewZapp email marketing services.<br /><br />You can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card. If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544 <br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts"

strInvoice8="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Licence due for Renewal </span><br><br>Our records show that your <strong>New</strong>Zapp licence is due for renewal and payment is due shortly.<br /><br />Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.  If you normally pay by credit card this invoice will be charged to your card. You can pay this invoice using a different credit card by logging in to your account.<br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br><br>Please ensure your payment reaches us by return to prevent any interruption to your <strong>New</strong>Zapp email marketing services.<br /><br />You can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card. If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544 <br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts"

strInvoice9="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Order</span><br><br>Thank you for choosing <strong>New</strong>Zapp for your email marketing.  We haven't received payment for your recent order as per the attached invoice.<br><br>Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.<br>You can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card.<p>If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544<br><br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts</p>"

strInvoice10="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#cc0000; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp  Final Reminder </span><br><br>Thank you for choosing <strong>New</strong>Zapp for your email marketing.  We haven't received payment for your recent order as per the attached invoice.<br><br>Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.<br>You can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card.<p>If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544<br><br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts</p>"

strInvoice11="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#cc0000; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp  Final Reminder </span><br><br>Thank you for choosing <strong>New</strong>Zapp for your email marketing.  We haven't received payment for your recent order as per the attached invoice.<br><br>Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.<br>You can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card.<p>If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544<br><br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts</p>"

strInvoice12="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#cc0000; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Licence due for Renewal: Final Reminder </span><br><br>Our records show that your next  <strong>New</strong>Zapp payment is due.<br /><br />Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.  If you normally pay by credit card this invoice will be charged to your card. You can pay this invoice using a different credit card by logging in to your account.<br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br /><br />Please ensure your payment reaches us by return to prevent any interruption to your <strong>New</strong>Zapp email marketing services.<br /><br />You can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card. If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544 <br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts"

strInvoice13="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#cc0000; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Licence: Content Deletion Notice </span><br><br>Our records show that your   <strong>New</strong>Zapp licence expired  more than 30 days ago. <br>Any content and subscriber data you had uploaded to <strong>New</strong>Zapp has been deleted. <br /><br /> If you have any queries  please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544 <br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts"

strInvoice14="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Licence: Please Review </span><br><br>There appears to be a problem with this invoice. <br /><br /> If you have any queries  service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544 <br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts"

strInvoice15="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Order</span><br><br>Thank you for choosing <strong>New</strong>Zapp for your email marketing.  <br><br>Please find your  <strong>New</strong>Zapp invoice attached to this email as a PDF document.<br>If you haven't already, you can pay this invoice by BACS, cheque or online via your <strong>New</strong>Zapp account using your debit or credit card.<p>We are unable to progress your order until payment has been   received.</p><p>If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544<br><br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts</p>"

strInvoice16="16 <strong>New</strong>Zapp Account"

strInvoice17="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;"">Online<strong>New</strong>Zapp Order</span><br><br>An order has been taken online.<br>Please find the <strong>New</strong>Zapp invoice attached to this email as a PDF document.<br><br>Kind Regards<br><br><strong>New</strong>Zapp Development" 

strInvoice18="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;"">Online<strong>New</strong>Zapp Order</span><br><br>An order has been taken online.<br>Please find the <strong>New</strong>Zapp invoice attached to this email as a PDF document.<br><br>Kind Regards<br><br><strong>New</strong>Zapp Development"

strInvoice19="<p><span style=""font-family:Arial, Helvetica, sans-serif; color:#006699; font-size:20px;padding:10px; padding-left:0px;""><strong>New</strong>Zapp Renewal </span><br><br>Dear Valued Customer<br><br>Your <strong>New</strong>Zapp renewal is due   shortly.&nbsp; Please find attached an invoice for the renewal of your annual <STRONG>New</STRONG>Zapp licence.<p>Why do we invoice your renewal   before your renewal is due?<br><br>It   gives you the opportunity to discuss with our expert marketing advisors your   email marketing needs for the coming year.&nbsp; Your needs may have changed and you   may require a&nbsp;different licence or additional services.&nbsp; <STRONG>New</STRONG>Zapp licences require payment in advance which helps   us to keep our prices as low as possible.&nbsp;&nbsp;Invoicing earlier than the renewal   date ensures you have plenty of time to pay for your renewal before your licence  is due.&nbsp; <br><br>We hope that you continue to use <STRONG>New</STRONG>Zapp for your email   marketing but in the event that you decide not to renew our early invoicing   allows you time to give us the 1 month notice that we require for   cancellation.<br><br>We look forward to working with you   for another year and thank you for choosing <STRONG>New</STRONG>Zapp.<br><br>Please find your <STRONG>New</STRONG>Zapp invoice attached to this   email as a PDF document.<br><br>If you have any queries relating to this invoice or your <strong>New</strong>Zapp service please contact us, Monday to Friday 9.00am - 5.00pm on 0845 612 5544<br><br><br><a href=""https://system.newzapp.co.uk/loginsystem/login/loginnewzapppaynow.asp?orderID=@orderID@""><img src=""http://system.newzapp.co.uk/images/loginButtonPayNow.gif"" alt=""Login"" width=""236"" height=""47"" border=""0""></a><br><br>Kind Regards<br><br><strong>New</strong>Zapp Accounts</p>"

strSubject1 = "1) Thank you for your order [EVERYONE GET THIS THANKYOU IF THEY PAID WITH ORDER]"
strSubject2 = "2) NewZapp Invoice [Standard Reminder]"
strSubject3 = "3) Invoice for review [Internal for Accounts]"
strSubject4 = "4) NewZapp Licence Expired [Account Expired]"
strSubject5 = "5) NewZapp Annual Renewal [Annual Renewal Payment Reminder]"
strSubject6 = "6) NewZapp Quarterly Renewal [Quarterly Renewal Payment Reminder]"
strSubject7 = "7) NewZapp Monthly Renewal [Monthly Renewal Payment Reminder]"
strSubject8 = "8) NewZapp Renewal [Unspecified Renewal Payment Reminder]"
strSubject9 = "9) NewZapp Invoice [Standard reminder that includes a licence but is not a renewal]"
strSubject10 = "10) NewZapp Invoice - Due Immediately [Final Reminder]"
strSubject11 = "11) NewZapp Invoice - Due Immediately [Final that includes a licence but is not a renewal]"
strSubject12 = "12) NewZapp Renewal - Due Immediately [Final Reminder to Renew]"
strSubject13 = "13) NewZapp Content Deletion Notice"
strSubject14 = "14) Problem with Order [Internal goes to tech dept]"
strSubject15 = "15) Thank you for your order [EVERYONE GET THIS THANKYOU IF THEY HAVE YET TO PAY]"
strSubject16 = "16) "
strSubject17 = "17) NewZapp Online Order Placed [Internal Order Notification]"
strSubject18 = "18) NewZapp Credit Note"
strSubject19 = "19) NewZapp Annual Renewal [Annual renewal 1st email]"


%>
<p>INVOICE EMAILS </p>
<p>Everybody gets a 1 or a 15</p>
<p>[Negative numbers indicate payment is overdue i.e should have been paid X days ago]<br>
  <br>
  Then<br>
  <br>
  <strong>' Standard Invoice </strong><br>
  Number of days from due date -7 they get an email of type 2<br>
  Number of days from due date -14 they get an email of type 2 <br>
  Number of days from due date -21 they get an email of type 2 <br>
  Number of days from due date -28 they get an email of type 10 <br>
  Number of days from due date -28 they get an email of type 3 <br>
  <br>
  <strong>' Annual Renewal</strong><br>
  Number of days from due date 14 they get an email of type 5<br>
  Number of days from due date 7 they get an email of type 5 <br>
  Number of days from due date 0 they get an email of type 12<br>
  Number of days from due date -1 they get an email of type 4<br>
  Number of days from due date -1 they get an email of type 3<br>
  <br>
  <strong>' Quartly Renewal</strong><br>
  Number of days from due date 14 they get an email of type 6<br>
  Number of days from due date 7 they get an email of type 6 <br>
  Number of days from due date 0 they get an email of type 12<br>
  Number of days from due date -1 they get an email of type 4<br>
  Number of days from due date -1 they get an email of type 3<br>
  <br>
  <strong>' Monthly Renewal</strong><br>
  Number of days from due date 2 they get an email of type 7 <br>
  Number of days from due date 0 they get an email of type 12<br>
  Number of days from due date -1 they get an email of type 4<br>
  Number of days from due date -1 they get an email of type 3<br>
  <br>
  <strong>' Renewal Unspecified Period</strong><br>
Number of days from due date 0 they get an email of type 3</p>
<p><strong>' Reminder Containing a Licence </strong><br>
  Number of days from due date -7 they get an email of type 9<br>
  Number of days from due date -14 they get an email of type 9 <br>
  Number of days from due date -21 they get an email of type 9 <br>
  Number of days from due date -28 they get an email of type 11 <br>
Number of days from due date -28 they get an email of type 3</p>
<p></p>
<%



If 1=2 Then
Response.Write strSubject1 & "<br><br>" & strHead & strInvoice1 & strFoot & "<br><br>"
Response.Write strSubject2 & "<br><br>" & strHead & strInvoice2 & strFoot & "<br><br>"
Response.Write strSubject3 & "<br><br>" & strHead & strInvoice3 & strFoot & "<br><br>"
Response.Write strSubject4 & "<br><br>" & strHead & strInvoice4 & strFoot & "<br><br>"
Response.Write strSubject5 & "<br><br>" & strHead & strInvoice5 & strFoot & "<br><br>"
Response.Write strSubject6 & "<br><br>" & strHead & strInvoice6 & strFoot & "<br><br>"
Response.Write strSubject7 & "<br><br>" & strHead & strInvoice7 & strFoot & "<br><br>"
Response.Write strSubject8 & "<br><br>" & strHead & strInvoice8 & strFoot & "<br><br>"
Response.Write strSubject9 & "<br><br>" & strHead & strInvoice9 & strFoot & "<br><br>"
Response.Write strSubject10 & "<br><br>" & strHead & strInvoice10 & strFoot & "<br><br>"
Response.Write strSubject11 & "<br><br>" & strHead & strInvoice11 & strFoot & "<br><br>"
Response.Write strSubject12 & "<br><br>" & strHead & strInvoice12 & strFoot & "<br><br>"
'Response.Write strSubject13 & "<br><br>" & strHead & strInvoice13 & strFoot & "<br><br>"
Response.Write strSubject14 & "<br><br>" & strHead & strInvoice14 & strFoot & "<br><br>"
Response.Write strSubject15 & "<br><br>" & strHead & strInvoice15 & strFoot & "<br><br>"
'Response.Write strSubject16 & "<br><br>" & strHead & strInvoice16 & strFoot & "<br><br>"
Response.Write strSubject17 & "<br><br>" & strHead & strInvoice17 & strFoot & "<br><br>"
'Response.Write strSubject18 & "<br><br>" & strHead & strInvoice18 & strFoot & "<br><br>"
End iF
Response.Write strSubject19 & "<br><br>" & strHead & strInvoice19 & strFoot & "<br><br>"


%>