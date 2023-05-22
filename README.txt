# NetworkPrinterAudit
This script polls all specified systems, read the currently logged on user, then identifies the printers currently installed


         File Name : NetworkPrinterAudit.ps1
   Original Author : Kenneth C. Mazie (kcmjr AT kcmjr DOT com)
                   :
       Description : This script polls all specified systems, read the currently logged
                   : on user, then identifies the printers currently installed. If the
                   : associated print server matches a "bad server" list The results are
                   : noted on screen and/or in an Excel spreadsheet. Output can be
                   : "server/printer" (full) or just "server" (brief).
                   :
         Arguments : Named command line parameters: (all are optional)
                   :
             Notes : This script was originally used during print server migration.
                   : The intent was to identify users who still had mapped printers
                   : pointing to the older print servers, hence the output only showing
                   : "bad" servers. Because the logged on user environment is volatile and
                   : may change at logoff you cannot read remote systems to determine printers.
                   : This is a best effort attempt to do just that, read the volatile session
                   : to collect installed printers.
                   :
          Warnings : None
                   :
             Legal : Public Domain. Modify and redistribute freely. No rights reserved.
                   : SCRIPT PROVIDED "AS IS" WITHOUT WARRANTIES OR GUARANTEES OF
                   : ANY KIND. USE AT YOUR OWN RISK. NO TECHNICAL SUPPORT PROVIDED.
                   :
           Credits : Code snippets and/or ideas came from many sources around the web.
                   :
    Last Update by : Kenneth C. Mazie
   Version History : v1.00 - 06-04-14 - Original
    Change History : v1.01 - 00-00-00 -
                   :
