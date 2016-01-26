'
' Cisco UCCX Wallboard 3.4
' Copyright (c) 2009 by Antoni Sawicki <as@tenoware.com>
'

$APPTYPE GUI
$TYPECHECK ON
$RESOURCE 0 as "icon.ico"

declare sub DoStuff
declare sub DoBlink
declare sub ShowCursor lib "user32" (bShow as long)
declare function GetSystemMetrics lib "user32" (nIndex as long) as long

dim db1 as sqldata
dim db2 as sqldata
dim d as date
deflng ret
defstr cfg,f1,f2,t
defint ivl=5
defint mqu=false
defstr ORG="Cisco UCCX"
defstr q, QNAME, SQLQNAME, ODBC_DSN1, ODBC_DSN2, ODBC_USERNAME, ODBC_PASSWORD, DSNUSED, QUERYCMD
deflng db1_time, db2_time
defint xres=GetSystemMetrics(0)
defint yres=GetSystemMetrics(1)
defint xpos,ypos
defint spc=1
defint showmousecursor=false
defint showdsn=false
defint w,h, w4, h9
defint n, dummy, inblink
defint wt,wtnew,cw,lc,ar,al,ta
defstr wt_str, wt_strnew, stbc
defint pfs=0
defint tfs=0
defint sleepat=0
defint wakeupat=0
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
    case "sleep_hour": if VAL(f2) > 0 then sleepat=VAL(f2)
    case "wakeup_hour": if VAL(f2) > 0 then wakeupat=VAL(f2)
    case "sla_waitqueue": if VAL(f2) > 0 then sla_waitqueue=VAL(f2)
    case "sla_waittime": if VAL(f2) > 0 then sla_waittime=VAL(f2)
    case "panel_fontsize": if VAL(f2) > 0 then pfs=VAL(f2)
    case "title_fontsize": if VAL(f2) > 0 then tfs=VAL(f2)
    case "org_name": if len(f2) > 1 then ORG=f2
    case "showmousecursor": if f2="yes" then showmousecursor=true
    case "http_host": if len(f2) > 1 then HTTPHOST=f2
    case "http_port": if VAL(f2) > 1 then HTTPPORT=VAL(f2)
    case "http_file": if len(f2) > 1 then HTTPFILE=f2
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

'
' Connect to SQL database
'
create b as splash
  width=200:height=50:center:caption=" UCCX Wallboard Loader":onkeydown=cleanup
  create msg as label
    top=0:left=10:width=b.width:height=b.height:caption="UCCX Wallboard: Trying ODBC..."
  end create
end create
b.show

db1.connect(ODBC_DSN1,ODBC_USERNAME,ODBC_PASSWORD)
db2.connect(ODBC_DSN2,ODBC_USERNAME,ODBC_PASSWORD)

' this NEVER happens
if db1.error > 1 and db2.error > 1 then 
    showmessage "Both ODBC data sources returned error 1="+hex$(db1.error)+" 2="+hex$(db2.error)
    goto cleanup
end if
msg.caption="  Connected..."


' panel resolution
w=xres+(4*spc)
w4=int(w/4)
h=yres+(6*spc)
h9=int(h/9)

' font size
if pfs=0 then pfs=2.8*h9
if tfs=0 then tfs=h9/2

' if this didn't have to run on Windows 98 we would have a thread
create nothreads as timer
  repeated=1
  interval=ivl*1000
  ontimer=DoStuff
  enabled=1
end create

create blinker as timer
  repeated=1
  interval=510
  ontimer=DoBlink
  enabled=0
end create

' enable blinker based on sla settings from config file
if sla_waitqueue>0 or sla_waittime>0 then blinker.enabled=1

create f as splash
  width=xres:height=yres
  color=spacercolor
  caption=" UCCX Wallboard"
  onkeydown=cleanup

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

  '
  ' statusbar
  '
  create statbar as label
    top=0:left=0:width=w-spc:height=h9-spc
    color=statbgcolor:textcolor=statfgcolor
    font=titlefont:style=statbar.style or &H1
    onmouseup=cleanup
    caption=ORG + " Wallboard 3.4"
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
    caption="??"
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
    caption="??:??"
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
    caption="??"
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
    caption="??"
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
    caption="??"
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
    caption="??"
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
    caption="??"
  end create

end create

' main
if showmousecursor=false then ShowCursor(0)
if xpos>0 and ypos>0 then
  f.top=ypos-1:f.left=xpos-1
else
  f.center
end if
b.visible=0
showconsole
f.showmodal

'
' query the database and update screen
'
sub DoStuff
  print TIME$
  db1.command(QUERYCMD)
  db2.command(QUERYCMD)

  ' obtain time of last update
  ' error checking is useless but row<>100 is required
  if not db1.error and db1.fieldcount=7 and db1.row<>100 then
    db1_time=VAL(db1.rowvalue(1,1))
    print "DB1=" + STR$(db1_time)
  end if

  if not db2.error and db2.fieldcount=7 and db2.row<>100 then
    db2_time=VAL(db2.rowvalue(1,1))
    print "DB2=" + STR$(db2_time)
  end if

  ' used data from the most up to date node
  if db1_time>db2_time then
    DSNUSED=ODBC_DSN1
    cw=VAL(db1.rowvalue(2,1))
    wt_str=right$(db1.rowvalue(3,1),5)
    ar=VAL(db1.rowvalue(4,1))
    ta=VAL(db1.rowvalue(5,1))
    lc=VAL(db1.rowvalue(6,1))
    wt=VAL(db1.rowvalue(7,1))
    wtnew=0
    wt_strnew=""

    ' if multiqueue there should be more data!
    while db1.row<>100
      cw=cw+VAL(db1.rowvalue(2,1))
      wt_strnew=right$(db1.rowvalue(3,1),5)    
      dummy=VAL(db1.rowvalue(4,1))
      dummy=VAL(db1.rowvalue(5,1))
      lc=lc+VAL(db1.rowvalue(6,1))
      wtnew=VAL(db1.rowvalue(7,1))
      if wtnew > wt then 
        wt=wtnew
        wt_str=wt_strnew
      end if
    end while

  elseif db2_time>db1_time then
    DSNUSED=ODBC_DSN2
    cw=VAL(db2.rowvalue(2,1))
    wt_str=right$(db2.rowvalue(3,1),5)
    ar=VAL(db2.rowvalue(4,1))
    ta=VAL(db2.rowvalue(5,1))
    lc=VAL(db2.rowvalue(6,1))
    wt=VAL(db2.rowvalue(7,1))
    wtnew=0
    wt_strnew=""

    ' if multiqueue there should be more data!
    while db2.row<>100
      cw=cw+VAL(db2.rowvalue(2,1))
      wt_strnew=right$(db2.rowvalue(3,1),5)    
      dummy=VAL(db2.rowvalue(4,1))
      dummy=VAL(db2.rowvalue(5,1))
      lc=lc+VAL(db2.rowvalue(6,1))
      wtnew=VAL(db2.rowvalue(7,1))
      if wtnew > wt then 
        wt=wtnew
        wt_str=wt_strnew
      end if
    end while
  else
    DSNUSED="ERROR"
    cw=-1
    wt_str="--"
    ar=-1
    ta=-1
    lc=-1
    wt=-1
    wtnew=0
    wt_strnew=""
  end if

  if showdsn=true or DSNUSED="ERROR" then
    stbc=ORG + "  " + QNAME + "  " + left$(TIME$, 5) + "  " + DSNUSED
  else
    stbc=ORG + "  " + QNAME + "  " + left$(TIME$, 5)
  end if


  ' dirty little hack to get direct calls from the billing server
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
         exit while
        end if
      end while
    end if
  end if
  res=peer.shutdown(sock,2)
  res=peer.close(sock)

  'cw=4

  ' update only on a change to prevent flicker
  if stbc <> statbar.caption then statbar.caption=stbc
  if str$(cw) <> pc_a.caption then pc_a.caption=str$(cw)
  if wt_str <> pc_b.caption then pc_b.caption=wt_str
  if str$(ar) <> pc_d.caption then pc_d.caption=str$(ar)
'  if str$(al) <> al2.caption then al2.caption=str$(al) not used
  if str$(ta) <> pc_e.caption then pc_e.caption=str$(ta)
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


  if wakeupat<>0 and sleepat<>0 and wakeupat<>sleepat then
    d.update
    if d.hour=wakeupat and d.minute=0 then ret=sendmessage(65535,274,61808,-1) 
    if d.hour=sleepat and d.minute=0 then  ret=sendmessage(65535,274,61808,2)  
  end if

  db1.freememory
  db2.freememory
  doevents
end sub

sub DoBlink
  if sla_waitqueue>0 and cw>=sla_waitqueue then
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

  if sla_waittime>0 and wt>=(sla_waittime*1000) then
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


cleanup:
ret=sendmessage(65535,274,61808,-1) 
db2.close
db2.close
app.terminate

PROP.FILEVERSION 3,4,0,0
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
PROP.VALUE "FileDescription", "Cisco UCCX Wallboard 3.4"
PROP.VALUE "FileVersion", "3.4.0.0" 
PROP.END  
PROP.END  
PROP.END  
