# Cisco IPCC / UCCX Full Screen Wallboard for Windows
*AKA Tenox Wallboard*

![alt text](https://raw.githubusercontent.com/tenox7/wallboard/master/wallboard.gif "Wallboard Screenshot")


## Intro
There are several different wallboard applications for Cisco IPCC/UCCX, this one however has a number of unique features. Namely, instead of being an ASP web page, it's a stand alone Windows executable application that connects to a remote Contact Center server (or servers) via ODBC. It doesn't require .NET, Java or a web browser. It will run on a very old PC. It displays only one specific queue in a full screen mode. It's main purpose is to be run on an airport style display per each team. And it's free. 

The wallboard comes in with a configuration file in where you can specify the queue name, odbc source, database credentials, colors, fonts, etc. Included are also installation instructions. Also possible to configure over threshold blinking fields, beep sound alert and an option of arbitrary position on the screen to run several instances on the same LCD panel. It possible to configure 12h or 24h time format and customize panel captions so you can localize them to your needs (eg a non-English language).


## Version Information
**Important:**

If you are using UCCX 7.x with MS-SQL you need to use Wallboard 2.x

If you are suing UCCX 8.x and above with Informix you need to use Wallboard 3.x

## Installation Instruction
Note these are for UCCX 7.x / Wallboard 2.x

1. You need to enable hybrid / mixed mode on the UCCX MS-SQL server. Yes it is supported by Cisco, however if you use the eporting tool you will need to set `AUTH=1` in `hrcConfig.ini` file.

2. You need to create a wallboard user on the SQL server and assign it `db_wallboard_read` role.
   
3. In the UCCX configuration page go to `Tools->Real Time Snapshot Config`. Enable data writing at 5 second interval and tick in both CSQs Summary and CCX System Summary. Don't bother putting Wallboard System server or user id / password. Leave them empty.
   
4. On the actual machine where wallboard exe will be running you need to configure a ODBC source pointing to the server using the SQL server authentication (second option). **Important info for 64bit Windows:** Wallboard is a 32bit application and therefore uses a 32bit ODBC connection. You need to configure the DSN under 32bit ODBC. To do that launch *c:\windows\syswow64\odbcad32.exe*. Test the connection with username and password you created in point 2. You can also use [SqlDbx](http://www.sqldbx.com/personal_edition.htm) to test if the ODBC connection can be made and data queried.
   
5. Configure `wallboard.cfg` with appropriate queue name, ODBC DSN and auth parameters (username/password). If you have two UCCX servers in a cluster you can have two ODBC DSNs created. Wallboard will fail over automatically.
   
6. Make sure that system clock is synchronized and Wallboard auto starts.

## Support 
If you are having issues setting it up please download [SqlDbx](http://www.sqldbx.com/personal_edition.htm) and connect to the same ODBC source as you defined in the wallboard config file. Make sure to use the same login credentials. If you are able to connect and browse RtCSQsSummary table then the problem is with Wallboard. If you cannot connect with SqlDbx then you have not configured UCCX or ODBC Driver correctly.

A more detailed installation instruction document is available [here](https://github.com/tenox7/wallboard/raw/master/wallboard-2x-install.doc).


