# Cisco IPCC / UCCX Full Screen Wallboard for Windows

*AKA Tenox Wallboard*

There are several different wallboard applications for Cisco IPCC/UCCX, this one however has a number of unique features. Namely, instead of being an ASP web page, it's a stand alone Windows executable application that connects to a remote Contact Center server (or servers) via ODBC. It doesn't require .NET, Java or a web browser. It will run on a very old PC. It displays only one specific queue in a full screen mode. It's main purpose is to be run on an airport style display per each team. And it's free. 

The wallboard comes in with a configuration file in where you can specify the queue name, odbc source, database credentials, colors, fonts, etc. Included are also installation instructions. Also possible to configure over threshold blinking fields, beep sound alert and an option of arbitrary position on the screen to run several instances on the same LCD panel. It possible to configure 12h or 24h time format and customize panel captions so you can localize them to your needs (eg a non-English language).

**Important: Versions 2.x are for UCCX 7.x ONLY!** If you are suing UCCX 8.x and above use use Wallboard 3.x

Wallboard version 3.7.3 is now available. It supports UCCX 8.x and above through the IBM Informix ODBC driver. The latest version fixes Multiqueue calculation issues. 

**Important info for 64bit Windows:** Wallboard is a 32bit application and therefore uses a 32bit ODBC connection. You need to configure the DSN under 32bit ODBC. To do that launch *c:\windows\syswow64\odbcad32.exe*.

Support note: if you are having issues setting it up please download SqlDbx Personal and connect to the same ODBC source as you defined in the wallboard config file. Make sure to use the same login credentials. If you are able to connect and browse RtCSQsSummary table then the problem is with Wallboard. If you cannot connect with SqlDbx then you have not configured UCCX or ODBC Driver correctly.


![alt text](https://raw.githubusercontent.com/tenox7/wallboard/master/wallboard.gif "Wallboard Screenshot")
