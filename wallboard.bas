'
' Cisco UCCX Wallboard 3.3
' Copyright (c) 2009 by Antoni Sawicki <as@tenoware.com>
'

$APPTYPE GUI
$TYPECHECK ON
$RESOURCE 0 as "icon.ico"

declare sub DoStuff
declare sub DoBlink
declare sub ShowCursor lib "user32" (bShow as long)
declare function GetSystemMetrics lib "user32" (nIndex as long) as long

dim db as sqldata
dim db2 as sqldata
dim d as date
deflng ret
defstr cfg,f1,f2,t
defint ivl=5
defint mqu=false
defstr ORG="Cisco UCCX"
defstr q, QNAME, SQLQNAME, ODBC_DSN1, ODBC_DSN2, ODBC_USERNAME, ODBC_PASSWORD, DSNUSED, QUERYCMD
defint xres=GetSystemMetrics(0)
defint yres=GetSystemMetrics(1)
defint xpos,ypos
defint spc=1
defint showmousecursor=false
defint showdsn=false
defint w,h, w4, h9
defint n, dummy, inblink
defint wt,wtnew,cw,lc,ar,al,ta
defstr wt3, wt3new, stbc
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

if len(ODBC_USERNAME) < 2 or ODBC_USERNAME="null" then ODBC_USERNAME=null
if len(ODBC_PASSWORD) < 2 or ODBC_PASSWORD="null" then ODBC_PASSWORD=null

if mqu=true then
  SQLQNAME=QNAME + "%"
else
  SQLQNAME=QNAME
end if

QUERYCMD="select callsWaiting,convOldestContact,availableAgents,talkingAgents,callsAbandoned,OldestContact" + _
         " from RtCSQsSummary where CSQName like '" + SQLQNAME + "';")

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

create b as splash
  width=200:height=50:center:caption=" UCCX Wallboard Loader":onkeydown=cleanup
  create msg as label
    top=0:left=10:width=b.width:height=b.height:caption="UCCX Wallboard 3.3: Trying " + ODBC_DSN1 + "..."
  end create
end create
b.show

db.connect(ODBC_DSN1,ODBC_USERNAME,ODBC_PASSWORD)
DSNUSED=ODBC_DSN1
' this NEVER happens
if db.error > 1 then 
  if len(ODBC_DSN2) > 1 then
    msg.caption="Failed. Trying " + ODBC_DSN2 + "...":msg.focus
    db.close
    db.connect(ODBC_DSN2,ODBC_USERNAME,ODBC_PASSWORD)
    DSNUSED=ODBC_DSN2
    if db.error > 1 then 
      showmessage "Secondary ODBC connect error="+hex$(db.error)
      goto cleanup
    end if
  else
    showmessage "Primary ODBC connect failure and no secondary defined. Error="+hex$(db.error)
    goto cleanup
  end if
end if
msg.caption="  Connected..."

w=xres+(4*spc)
w4=int(w/4)
h=yres+(6*spc)
h9=int(h/9)

if pfs=0 then pfs=2.8*h9
if tfs=0 then tfs=h9/2

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
    caption=ORG + " Wallboard 3.3"
  end create

  '
  ' top row
  '
  create cw1 as label
    top=h9:left=0:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=cw1.style or &H1
    caption="Wait Queue"
  end create

  create cw2 as label
    top=2*h9:left=0:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=cw2.style or &H1
    caption="??"
  end create

  create wt1 as label
    top=h9:left=w4:width=2*w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=wt1.style or &H1
    caption="Oldest Caller Wait Time"
  end create

  create wt2 as label
    top=2*h9:left=w4:width=(2*w4)-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=wt2.style or &H1
    caption="??:??"
  end create

  create lc1 as label
    top=h9:left=3*w4:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=wt1.style or &H1
    caption="Lost Calls"
  end create

  create lc2 as label
    top=2*h9:left=3*w4:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=wt2.style or &H1
    caption="??"
  end create

  '
  ' bottom row
  '
  create ar1 as label
    top=5*h9:left=0:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=ar1.style or &H1
    caption="Ready Agents"
  end create

  create ar2 as label
    top=6*h9:left=0:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=ar2.style or &H1
    caption="??"
  end create

  create ta1 as label
    top=5*h9:left=w4:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=ta1.style or &H1
    caption="Talking Agents"
  end create

  create ta2 as label
    top=6*h9:left=w4:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=ta2.style or &H1
    caption="??"
  end create

  create oc1 as label
    top=5*h9:left=2*w4:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=oc1.style or &H1
    caption="Outbound Calls"
  end create

  create oc2 as label
    top=6*h9:left=2*w4:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=oc2.style or &H1
    caption="??"
  end create

  create ic1 as label
    top=5*h9:left=3*w4:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=ic1.style or &H1
    caption="Inbound Calls"
  end create

  create ic2 as label
    top=6*h9:left=3*w4:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=ic2.style or &H1
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
f.showmodal

'
' query the database and update screen
'
sub DoStuff
  db.command(QUERYCMD)

  if db.error then
    db.freememory
    db.close
      if DSNUSED=ODBC_DSN1 then
        if len(ODBC_DSN2) > 1 then
          db.connect(ODBC_DSN2,ODBC_USERNAME,ODBC_PASSWORD)
          DSNUSED=ODBC_DSN2
        else
          showmessage "Primary ODBC DSN failure and no secondary present. Error="+hex$(db.error)
          goto cleanup
        end if
      else 
        ' if we were on secondary, the primary must be defined..
        db.connect(ODBC_DSN1,ODBC_USERNAME,ODBC_PASSWORD)
        DSNUSED=ODBC_DSN1
      end if
      db.command(QUERYCMD)
      if db.error then
        showmessage "Both ODBC DSNs query error="+hex$(db.error)
        goto cleanup
      end if
    end if

  if db.row=100 then
    showmessage "ODBC query returned no data"
    goto cleanup
  end if
  
  if db.fieldcount<>6 then 
    showmessage "ODBC query returned wrong number of columns"
    goto cleanup
  end if

  if showdsn=true then
    stbc=ORG + "  " + QNAME + "  " + left$(TIME$, 5) + "  " + DSNUSED
  else
    stbc=ORG + "  " + QNAME + "  " + left$(TIME$, 5)
  end if

  cw=VAL(db.rowvalue(1,1))
  wt3=right$(db.rowvalue(2,1),5)
  ar=VAL(db.rowvalue(3,1))
  ta=VAL(db.rowvalue(4,1))
  lc=VAL(db.rowvalue(5,1))
  wt=val(db.rowvalue(6,1))
  wtnew=0
  wt3new=""

  ' if multiqueue there should be more data!
  while db.row<>100
    cw=cw+VAL(db.rowvalue(1,1))
    wt3new=right$(db.rowvalue(2,1),5)    
    dummy=VAL(db.rowvalue(3,1))
    dummy=VAL(db.rowvalue(4,1))
    lc=lc+VAL(db.rowvalue(5,1))
    wtnew=VAL(db.rowvalue(6,1))
    if wtnew > wt then 
      wt=wtnew
      wt3=wt3new
    end if
  end while

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
  if str$(cw) <> cw2.caption then cw2.caption=str$(cw)
  if wt3 <> wt2.caption then wt2.caption=wt3
  if str$(ar) <> ar2.caption then ar2.caption=str$(ar)
'  if str$(al) <> al2.caption then al2.caption=str$(al) not used
  if str$(ta) <> ta2.caption then ta2.caption=str$(ta)
  if str$(lc) <> lc2.caption then lc2.caption=str$(lc)

  if inbound_calls <> ic2.caption then 
    ic2.caption=inbound_calls
    if len(ic2.caption) > 2 then 
      ic2.font=notbigfont
    else 
      ic2.font=bigfont
    end if
  end if

  if outbound_calls <> oc2.caption then 
    oc2.caption=outbound_calls
    if len(oc2.caption) > 2 then 
      oc2.font=notbigfont
    else 
      oc2.font=bigfont
    end if
  end if


  if wakeupat<>0 and sleepat<>0 and wakeupat<>sleepat then
    d.update
    if d.hour=wakeupat and d.minute=0 then ret=sendmessage(65535,274,61808,-1) 
    if d.hour=sleepat and d.minute=0 then  ret=sendmessage(65535,274,61808,2)  
  end if

  db.freememory
  doevents
end sub

sub DoBlink
  if sla_waitqueue>0 and cw>=sla_waitqueue then
    if cw2.textcolor=labelfgcolor then
      cw2.textcolor=blinkcolor
    else
      cw2.textcolor=labelfgcolor
    end if
    cw2.repaint
  else
    if cw2.textcolor<>labelfgcolor then
      cw2.textcolor=labelfgcolor
      cw2.repaint
    end if
  end if

  if sla_waittime>0 and wt>=(sla_waittime*1000) then
    if wt2.textcolor=labelfgcolor then
      wt2.textcolor=blinkcolor
    else
      wt2.textcolor=labelfgcolor
    end if
    wt2.repaint
  else 
    if wt2.textcolor<>labelfgcolor then
      wt2.textcolor=labelfgcolor
      wt2.repaint
    end if
  end if
  doevents
end sub


cleanup:
ret=sendmessage(65535,274,61808,-1) 
db.close
app.terminate

PROP.FILEVERSION 3,3,0,0
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
PROP.VALUE "FileDescription", "Cisco UCCX Wallboard 3.3"
PROP.VALUE "FileVersion", "3.2.0.0" 
PROP.END  
PROP.END  
PROP.END  
