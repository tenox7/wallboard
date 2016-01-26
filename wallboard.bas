'
' Cisco UCCX Wallboard 3.5
' Copyright (c) 2009-2011 by Antoni Sawicki <as@tenoware.com>
'

$APPTYPE GUI
$TYPECHECK ON
$RESOURCE 0 as "icon.ico"

declare sub db1_worker() std
declare sub db2_worker() std
declare sub http_worker() std
'declare sub screen_paint() std
declare sub DoBlink()
declare sub keypress
declare sub ShowCursor lib "user32" (bShow as long)
declare sub logevent(myevent as string)
declare function GetSystemMetrics lib "user32" (nIndex as long) as long
declare function tval std (lpstr as long) as long

dim db1 as sqldata
dim db2 as sqldata
dim d as date
dim logfile as file
defstr LOGFILENAME
deflng ret
defstr cfg,f1,f2,t
defint ivl=5
defint mqu=false
defint windowed=false
defint logguievents=false
defstr ORG="Cisco UCCX"
defstr q, QNAME, SQLQNAME, ODBC_DSN1, ODBC_DSN2, ODBC_USERNAME, ODBC_PASSWORD, DSNUSED, QUERYCMD, last_update_str
defdword last_update=0
defint xres=GetSystemMetrics(0)
defint yres=GetSystemMetrics(1)
defint xpos,ypos
defint spc=1
defint fclose=0
defint showmousecursor=false
defint showdsn=false
defint w,h, w4, h9
defint n, inblink
defint wt,wtnew,cw,lc,al
defstr wt_str, wt_strnew, stbc, ar, ta, dummy
defint pfs=0
defint tfs=0
defint sla_waitqueue=0
defint sla_waittime=0
defint spacercolor=&HFFFFFF
defint labelfgcolor=&HFFFFFF
defint labelbgcolor=&H800000
defint statbgcolor=&HFF0000
defint statfgcolor=&HFFFFFF
defint blinkcolor=&H0000FF
defstr pnlfont="Arial Narrow"
defstr ttlfont="Arial Narrow"
deflng sock, res 
defint HTTPPORT
defstr buff, inbound_calls, outbound_calls, HTTPCMD, HTTPHOST, HTTPFILE
dim time0 as REAL10
dim peer as SOCKET

'
' read config file
'
if not fileexists("wallboard.cfg") then 
  showmessage "ERROR: Unable to open configuration file!"
  app.terminate
end if

cfg.loadfromfile "wallboard.cfg"
for n=0 to cfg.itemcount
  f1=field$(cfg.item(n), "=", 1):f2=field$(cfg.item(n), "=", 2)
  select case f1
    case "queue_name": if len(f2) > 1 then QNAME=f2
    case "odbc_dsn1": if len(f2) > 1 then ODBC_DSN1=f2
    case "odbc_dsn2": if len(f2) > 1 then ODBC_DSN2=f2
    case "odbc_username": if len(f2) > 1 then ODBC_USERNAME=f2
    case "odbc_password": if len(f2) > 1 then ODBC_PASSWORD=f2
    case "interval": if VAL(f2) > 5 then ivl=VAL(f2)
    case "spacer_size": if VAL(f2) > 0 then spc=VAL(f2)
    case "custom_xres": if VAL(f2) > 0 then xres=VAL(f2)
    case "custom_yres": if VAL(f2) > 0 then yres=VAL(f2)
    case "custom_xpos": if VAL(f2) > 0 then xpos=VAL(f2)
    case "custom_ypos": if VAL(f2) > 0 then ypos=VAL(f2)
    case "multiqueue": if f2="yes" then mqu=true
    case "showdsn": if f2="yes" then showdsn=true
    case "panel_font": if len(f2) > 1 then pnlfont=f2
    case "title_font": if len(f2) > 1 then ttlfont=f2
    case "sla_waitqueue": if VAL(f2) > 0 then sla_waitqueue=VAL(f2)
    case "sla_waittime": if VAL(f2) > 0 then sla_waittime=VAL(f2)
    case "panel_fontsize": if VAL(f2) > 0 then pfs=VAL(f2)
    case "title_fontsize": if VAL(f2) > 0 then tfs=VAL(f2)
    case "org_name": if len(f2) > 1 then ORG=f2
    case "showmousecursor": if f2="yes" then showmousecursor=true
    case "logguievents": if f2="yes" then logguievents=true
    case "windowed": if f2="yes" then windowed=true
    case "http_host": if len(f2) > 1 then HTTPHOST=f2
    case "http_port": if VAL(f2) > 1 then HTTPPORT=VAL(f2)
    case "http_file": if len(f2) > 1 then HTTPFILE=f2
    case "logfile": if len(f2) > 1 then LOGFILENAME=f2
    case "spacer_color": if len(f2) = 6 then 
        t.clear:t.append f2[5],f2[6],f2[3],f2[4],f2[1],f2[2]:spacercolor=HEX2DW(t)
      end if
    case "panel_fgcolor": if len(f2) = 6 then
        t.clear:t.append f2[5],f2[6],f2[3],f2[4],f2[1],f2[2]:labelfgcolor=HEX2DW(t)
      end if
    case "panel_bgcolor": if len(f2) = 6 then
        t.clear:t.append f2[5],f2[6],f2[3],f2[4],f2[1],f2[2]:labelbgcolor=HEX2DW(t)
      end if
    case "status_fgcolor": if len(f2) = 6 then 
        t.clear:t.append f2[5],f2[6],f2[3],f2[4],f2[1],f2[2]:statfgcolor=HEX2DW(t)
      end if
    case "status_bgcolor": if len(f2) = 6 then 
        t.clear:t.append f2[5],f2[6],f2[3],f2[4],f2[1],f2[2]:statbgcolor=HEX2DW(t)
      end if
    case "blink_fgcolor": if len(f2) = 6 then 
        t.clear:t.append f2[5],f2[6],f2[3],f2[4],f2[1],f2[2]:blinkcolor=HEX2DW(t)
      end if
  end select
next n

if len(QNAME) < 2 then
  showmessage "ERROR: queue_name not defined"
  app.terminate
end if 

if len(ODBC_DSN1) < 2 then
  showmessage "ERROR: odbc_dsn1 not defined"
  app.terminate
end if

if ODBC_DSN1=ODBC_DSN2 then
  showmessage "ERROR: dsn1 is same as dsn2"
  app.terminate
end if

if len(ODBC_USERNAME) < 2 or ODBC_USERNAME="null" then ODBC_USERNAME=null
if len(ODBC_PASSWORD) < 2 or ODBC_PASSWORD="null" then ODBC_PASSWORD=null

if mqu=true then
  SQLQNAME=QNAME + "%"
else
  SQLQNAME=QNAME
end if

QUERYCMD="select timestamp = DATEDIFF(s, '19700101', endDateTime), " + _
	   "callsWaiting,convOldestContact,availableAgents,talkingAgents,callsAbandoned,OldestContact " + _
         "from RtCSQsSummary where CSQName like '" + SQLQNAME + "';")

HTTPCMD="GET " + HTTPFILE + " HTTP/1.0" + CRLF

if len(HTTPHOST) < 6 then
  showmessage "ERROR: http_host not defined"
  app.terminate
end if

if HTTPPORT < 2 then
  showmessage "ERROR: http_port not defined"
  app.terminate
end if

if len(HTTPFILE) < 4 then
  showmessage "ERROR: http_file not defined"
  app.terminate
end if

logevent("Wallboard starting....")
logevent("HWCAPS: XRES="+str$(xres)+" YRES="+str$(yres)+" NCPU="+str$(CPUCOUNT))
logevent("QNAME="+QNAME+" DSN1="+ODBC_DSN1+" DSN2="+ODBC_DSN2)
logevent("HTTP=http://"+HTTPHOST+":"+str$(HTTPPORT)+""+HTTPFILE)

' panel resolution
w=xres+(4*spc)
w4=int(w/4)
h=yres+(6*spc)
h9=int(h/9)

' font size
if pfs=0 then pfs=2.8*h9
if tfs=0 then tfs=h9/2

create blinker as timer
  repeated=1
  interval=510
  ontimer=DoBlink
  enabled=0
end create

' enable blinker based on sla settings from config file
if sla_waitqueue>0 or sla_waittime>0 then blinker.enabled=1

create bigfont as font
  name=pnlfont
  height=pfs
  weight=1000
end create

create notbigfont as font
  name=pnlfont
  height=2*pfs/3
  weight=1000
end create

create titlefont as font
  name=ttlfont
  height=tfs
  weight=700
end create


create f as form
  width=xres:height=yres
  color=spacercolor
  caption=" UCCX Wallboard"
  onkeydown=keypress


  '
  ' statusbar
  '
  create statbar as label
    top=0:left=0:width=w-spc:height=h9-spc
    color=statbgcolor:textcolor=statfgcolor
    font=titlefont:style=statbar.style or &H1
    onmouseup=cleanup
    caption=ORG + " Wallboard 3.5"
  end create

  '
  ' Panel Map:
  ' A B B C
  ' D E F G
  '
  ' pt for oanel title, pc for panel content/cell
  '

  '
  ' top row
  '
  create pt_a as label
    top=h9:left=0:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=pt_a.style or &H1
    caption="Wait Queue"
  end create

  create pc_a as label
    top=2*h9:left=0:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=pc_a.style or &H1
  end create

  create pt_b as label
    top=h9:left=w4:width=2*w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=pt_b.style or &H1
    caption="Oldest Caller Wait Time"
  end create

  create pc_b as label
    top=2*h9:left=w4:width=(2*w4)-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=pc_b.style or &H1
  end create

  create pt_c as label
    top=h9:left=3*w4:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=pt_c.style or &H1
    caption="Lost Calls"
  end create

  create pc_c as label
    top=2*h9:left=3*w4:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=pc_c.style or &H1
  end create

  '
  ' bottom row
  '
  create pt_d as label
    top=5*h9:left=0:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=pt_d.style or &H1
    caption="Ready Agents"
  end create

  create pc_d as label
    top=6*h9:left=0:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=pc_d.style or &H1
  end create

  create pt_e as label
    top=5*h9:left=w4:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=pt_e.style or &H1
    caption="Talking Agents"
  end create

  create pc_e as label
    top=6*h9:left=w4:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=pc_e.style or &H1
  end create

  create pt_f as label
    top=5*h9:left=2*w4:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=pt_f.style or &H1
    caption="Outbound Calls"
  end create

  create pc_f as label
    top=6*h9:left=2*w4:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=pc_f.style or &H1
  end create

  create pt_g as label
    top=5*h9:left=3*w4:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=pt_g.style or &H1
    caption="Inbound Calls"
  end create

  create pc_g as label
    top=6*h9:left=3*w4:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=pc_g.style or &H1
  end create

end create

dim ds as splash
dim df as form



' main
if showmousecursor=false then ShowCursor(0)
if logguievents then f.onmessage=msgcapture
if windowed=false then 
  f.style=ds.style
  f.exstyle=ds.exstyle
end if
if xpos>0 and ypos>0 then
  f.top=ypos-1:f.left=xpos-1
else
  f.center
end if
'ret=createthread(codeptr(screen_paint),0)
ret=createthread(codeptr(http_worker),0)
ret=createthread(codeptr(db1_worker),0)
if len(ODBC_DSN2) > 1 then
  sleep 0.1
  ret=createthread(codeptr(db2_worker),0)
end if
logevent(hex$(f.exstyle))
f.show
do
'  begin thread
'    logevent("SCRN: DSN="+DSNUSED+" WQ="+str$(cw)+" WT="+wt_str+" AR="+str$(ar)+ _
'             " AT="+str$(ta)+" IC="+inbound_calls+" OC="+outbound_calls+" LC="+str$(lc))
    if showdsn=true then
      stbc=ORG + "  " + QNAME + "  " + left$(last_update_str, 5) + "  " + DSNUSED
    else
      stbc=ORG + "  " + QNAME + "  " + left$(last_update_str, 5)
    end if

   ' update only on a change to prevent flicker
    if stbc <> statbar.caption then statbar.caption=stbc
    if str$(cw) <> pc_a.caption then pc_a.caption=str$(cw)
    if wt_str <> pc_b.caption then pc_b.caption=right$(wt_str,5)
    if ar <> pc_d.caption then pc_d.caption=ar
  '  if str$(al) <> al2.caption then al2.caption=str$(al) not used
    if ta <> pc_e.caption then pc_e.caption=ta
    if str$(lc) <> pc_c.caption then pc_c.caption=str$(lc)

    if inbound_calls <> pc_g.caption then 
      pc_g.caption=inbound_calls
      if len(pc_g.caption) > 2 then 
        pc_g.font=notbigfont
      else 
        pc_g.font=bigfont
      end if
    end if

    if outbound_calls <> pc_f.caption then 
      pc_f.caption=outbound_calls
      if len(pc_f.caption) > 2 then 
        pc_f.font=notbigfont
      else 
        pc_f.font=bigfont
      end if
    end if

'  end thread
  doevents
  sleep 0.5
loop

cleanup:
  fclose=1
  logevent("Cleanup invoked, exiting...")
  if logfile.handle then: logfile.close: end if
  app.terminate
end


sub db1_worker()
  defdword db1_time=0
  defstr cline
  do
    if fclose=1 then exit sub

    db1.command(QUERYCMD)

    if db1.error=0 and db1.fieldcount=7 and db1.row<>100 then
      cline=db1.rowvalue(1,1): db1_time=tval(@cline)
      logevent("DB_1: OK Time=" + str$(db1_time) + " err=" + str$(db1.error))

      if db1_time > last_update then
          logevent("DB_1: Updating values... new=" + str$(db1_time) + " old=" + str$(last_update))
          last_update=db1_time
          last_update_str=left$(TIME$, 5)
          DSNUSED=ODBC_DSN1
          cline=db1.rowvalue(2,1): cw=tval(@cline)
          wt_str=db1.rowvalue(3,1) 
          ar=db1.rowvalue(4,1)
          ta=db1.rowvalue(5,1)
          cline=db1.rowvalue(6,1): lc=tval(@cline)
          cline=db1.rowvalue(7,1): wt=tval(@cline)
          wtnew=0
          wt_strnew=""

          ' if multiqueue there should be more data!
          while db1.row<>100
            cline=db1.rowvalue(2,1): cw=cw+tval(@cline)
            wt_strnew=db1.rowvalue(3,1)
            dummy=db1.rowvalue(4,1)
            dummy=db1.rowvalue(5,1)
            cline=db1.rowvalue(6,1): lc=lc+tval(@cline)
            cline=db1.rowvalue(7,1): wtnew=tval(@cline)
            if wtnew > wt then 
              wt=wtnew
              wt_str=wt_strnew
            end if
          end while
      end if
    else
      logevent("DB_1: Reconnecting... err=" + str$(db1.error))
      db1.connect(ODBC_DSN1,ODBC_USERNAME,ODBC_PASSWORD)
    end if

    db1.freememory
    doevents
    sleep ivl
  loop
end sub

sub db2_worker()
  defdword db2_time=0
  defstr cline
  do
    if fclose=1 then exit sub

    db2.command(QUERYCMD)

    if db2.error=0 and db2.fieldcount=7 and db2.row<>100 then
      cline=db2.rowvalue(1,1): db2_time=tval(@cline)
      logevent("DB_2: OK Time=" + str$(db2_time) + " err=" + str$(db2.error))

      if db2_time > last_update then
          logevent("DB_2: Updating values... new=" + str$(db2_time) + " old=" + str$(last_update))
          last_update=db2_time
          last_update_str=left$(TIME$, 5)
          DSNUSED=ODBC_DSN1
          cline=db2.rowvalue(2,1): cw=tval(@cline)
          wt_str=db2.rowvalue(3,1) 
          ar=db2.rowvalue(4,1)
          ta=db2.rowvalue(5,1)
          cline=db2.rowvalue(6,1): lc=tval(@cline)
          cline=db2.rowvalue(7,1): wt=tval(@cline)
          wtnew=0
          wt_strnew=""

          ' if multiqueue there should be more data!
          while db2.row<>100
            cline=db2.rowvalue(2,1): cw=cw+tval(@cline)
            wt_strnew=db2.rowvalue(3,1)
            dummy=db2.rowvalue(4,1)
            dummy=db2.rowvalue(5,1)
            cline=db2.rowvalue(6,1): lc=lc+tval(@cline)
            cline=db2.rowvalue(7,1): wtnew=tval(@cline)
            if wtnew > wt then 
              wt=wtnew
              wt_str=wt_strnew
            end if
          end while
      end if
    else
      logevent("DB_2: Reconnecting... err=" + str$(db2.error))
      db2.connect(ODBC_DSN2,ODBC_USERNAME,ODBC_PASSWORD)
    end if

    db2.freememory
    doevents
    sleep ivl
  loop
end sub



sub http_worker()
  do
    if fclose=1 then exit sub
    sock=peer.s
    if sock>=0 then 
      res=peer.connect(sock,HTTPHOST,HTTPPORT)
      if res>=0 then
        res=peer.writeline(sock,HTTPCMD)
        time0=timer
        while timer-time0<3 
          'print timer-time0
          sleep 0.1
          if peer.isserverready(sock) then
            buff=peer.read(sock,1024)
            inbound_calls=field$(buff.item(buff.itemcount), ",", 2)
            outbound_calls=field$(buff.item(buff.itemcount), ",", 3)
            logevent("HTTP: ic=" + inbound_calls + " oc=" + outbound_calls)
           exit while
          end if
        end while
      end if
    end if
    res=peer.shutdown(sock,2)
    res=peer.close(sock)
    doevents
    sleep ivl
  loop
end sub


sub DoBlink()
  defint lcw, lwt
  begin thread
    lcw=cw
    lwt=wt
  end thread

  if sla_waitqueue>0 and lcw>=sla_waitqueue then
    if pc_a.textcolor=labelfgcolor then
      pc_a.textcolor=blinkcolor
    else
      pc_a.textcolor=labelfgcolor
    end if
    pc_a.repaint
  else
    if pc_a.textcolor<>labelfgcolor then
      pc_a.textcolor=labelfgcolor
      pc_a.repaint
    end if
  end if

  if sla_waittime>0 and lwt>=(sla_waittime*1000) then
    if pc_b.textcolor=labelfgcolor then
      pc_b.textcolor=blinkcolor
    else
      pc_b.textcolor=labelfgcolor
    end if
    pc_b.repaint
  else 
    if pc_b.textcolor<>labelfgcolor then
      pc_b.textcolor=labelfgcolor
      pc_b.repaint
    end if
  end if
  doevents
end sub

sub logevent(myevent as string)
  begin thread
  if LOGFILENAME="console" then
    begin runonce: showconsole: end runonce
    print DATE$ + " " + TIME$ + " : " + myevent
  elseif len(LOGFILENAME) > 4 then
    if not logfile.handle then
      if fileexists(LOGFILENAME) then 
        delete LOGFILENAME+".old"
        rename LOGFILENAME, LOGFILENAME+".old"
      end if
      logfile.open(LOGFILENAME,1)
    end if 
    logfile.writeline(DATE$ + " " + TIME$ + " : " + myevent)
  end if
  end thread
end sub

sub keypress
  logevent("Key Pressed="+str$(wParam))
  if wParam=27 or wParam=32 or wParam=81 then goto cleanup
end sub

function tval(lpstr as long)
  begin thread
    defstr v$=byref$(lpstr)
    result=VAL(v$)
  end thread
end function

msgcapture:
  if uMsg<>&H20 then
    logevent(HEX$(hWnd)+space+HEX$(uMsg)+space+HEX$(wParam)+space+HEX$(lParam) )
  end if
  retval zero
return


PROP.FILEVERSION 3,5,0,0
PROP.PRODUCTVERSION 0,0,0,0
PROP.FILEFLAGSMASK 0x0000003FL
PROP.FILEFLAGS 0x0000000BL
PROP.FILEOS 0x00010001L
PROP.FILETYPE 0x00000001L
PROP.FILESUBTYPE 0x00000000L
PROP.BEGIN
PROP.StringFileInfo
PROP.BEGIN
PROP.BLOCK "040904E4"
PROP.BEGIN
PROP.VALUE "Author","Antoni Sawicki"
PROP.VALUE "FileDescription", "Cisco UCCX Wallboard 3.5"
PROP.VALUE "FileVersion", "3.5.0.0" 
PROP.END  
PROP.END  
PROP.END  
