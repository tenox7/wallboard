'
' Cisco UCCX Wallboard 3.7.3
' Copyright (c) 2009-2016 by Antoni Sawicki <as@tenoware.com>
'

$APPTYPE GUI
$TYPECHECK ON
$RESOURCE 0 as "icon.ico"

declare sub db1_worker() std
declare sub db2_worker() std
declare sub http_worker() std
declare sub DoBlink()
declare sub print_screen()
declare sub keypress
declare sub ShowCursor lib "user32" (bShow as long)
declare sub logevent(myevent as string)
declare function GetSystemMetrics lib "user32" (nIndex as long) as long

defstr WBVER="3.7.3"
dim db1 as sqldata
dim db2 as sqldata
dim d as date
dim logfile as file
deflng ret, sock, res 
defstr LOGFILENAME,cfg,f1,f2,t
defint ivl=5
defint mqu=false : defint windowed=false : defint logguievents=false
defstr ORG="Cisco UCCX"
defstr q, QNAME, SQLQNAME, SQLDISPTIMEFORMAT, ODBC_DSN1, ODBC_DSN2, ODBC_USERNAME, ODBC_PASSWORD, DSNUSED, QUERYCMD, last_update_str, stbc
defint FIELDCNT
defstr last_update, db1_time, db2_time, db1_hhmm, db2_hhmm
defint xres=GetSystemMetrics(0) : defint yres=GetSystemMetrics(1)
defint xpos,ypos
defint fclose=0
defint spc=1
defint disptimeformat=24
defint showmousecursor=false : defint showdsn=false
defint w,h, w4, h9, n, inblink, ptr
defint db1_wt,db1_wtnew,db1_cw,db1_cw_m,db1_lc,db1_lc_m,db1_al,db1_tc,db1_tc_m
defstr db1_wt_str, db1_wt_strnew, db1_ar, db1_ta, db1_oa, db1_dummy
defint db2_wt,db2_wtnew,db2_cw,db2_cw_m,db2_lc,db2_lc_m,db2_al,db2_tc,db2_tc_m
defstr db2_wt_str, db2_wt_strnew, db2_ar, db2_ta,db2_oa,db2_dummy
defint scr_wt,scr_cw,scr_lc,scr_al,scr_tc
defstr scr_wt_str, scr_ar, scr_ta,scr_oa
defint pfs=0 : defint tfs=0
defint sla_waitqueue=0 : defint sla_waittime=0
defint spacercolor=&HFFFFFF : defint labelfgcolor=&HFFFFFF : defint labelbgcolor=&H800000 
defint statbgcolor=&HFF0000 : defint statfgcolor=&HFFFFFF :defint blinkcolor=&H0000FF
defstr pnlfont="Arial Narrow" : defstr ttlfont="Arial Narrow"
defstr buff
defstr caption_a="Wait Queue"
defstr caption_b="Oldest Caller Wait Time"
defstr caption_c="Lost Calls"
defstr caption_d="Ready Agents"
defstr caption_e="Talking Agents"
defstr caption_f="Online Agents"
defstr caption_g="Total Calls"

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
    case "caption_a": if len(f2) > 1 then caption_a=f2
    case "caption_b": if len(f2) > 1 then caption_b=f2
    case "caption_c": if len(f2) > 1 then caption_c=f2
    case "caption_d": if len(f2) > 1 then caption_d=f2
    case "caption_e": if len(f2) > 1 then caption_e=f2
    case "caption_f": if len(f2) > 1 then caption_f=f2
    case "caption_g": if len(f2) > 1 then caption_g=f2
    case "multiqueue": if f2="yes" then mqu=true
    case "showdsn": if f2="yes" then showdsn=true
    case "disptimeformat": if f2="12h" then disptimeformat=12
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

if disptimeformat=12 then
  SQLDISPTIMEFORMAT="%I:%M"
else
  SQLDISPTIMEFORMAT="%H:%M"
end if

QUERYCMD="select TO_CHAR(enddatetime, '%y%m%d%H%M%S'), " + _
	   "callsWaiting,convOldestContact,availableAgents,talkingAgents,callsAbandoned,OldestContact,loggedInAgents,totalCalls, " + _
         "TO_CHAR(enddatetime, '" + SQLDISPTIMEFORMAT + "') " + _
         "from RtCSQsSummary where CSQName like '" + SQLQNAME + "';")
FIELDCNT=10

logevent("Tenox Wallboard " + WBVER + " starting...")
logevent("HWCAPS: XRES="+str$(xres)+" YRES="+str$(yres)+" NCPU="+str$(CPUCOUNT))
logevent("QNAME="+QNAME+" DSN1="+ODBC_DSN1+" DSN2="+ODBC_DSN2)

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

create scrprint as timer
  repeated=1
  interval=ivl*1000
  ontimer=print_screen
  enabled=1
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
    caption=ORG + " Wallboard " + WBVER
  end create

  '
  ' Panel Map:
  ' A B B C
  ' D E F G
  '
  ' pt for panel title, pc for panel content/cell
  '

  '
  ' top row
  '
  create pt_a as label
    top=h9:left=0:width=w4-spc:height=h9-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=titlefont:style=pt_a.style or &H1
    caption=caption_a
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
    caption=caption_b
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
    caption=caption_c
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
    caption=caption_d
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
    caption=caption_e
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
    caption=caption_f
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
    caption=caption_g
  end create

  create pc_g as label
    top=6*h9:left=3*w4:width=w4-spc:height=(3*h9)-spc
    color=labelbgcolor:textcolor=labelfgcolor
    font=bigfont:style=pc_g.style or &H1
    caption="??"
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
logevent("Starting DB1 Worker Thread for ODBC_DSN1=" + ODBC_DSN1)
ret=createthread(codeptr(db1_worker),0)
if len(ODBC_DSN2) > 1 then
  sleep 0.1
  logevent("Starting DB2 Worker Thread for ODBC_DSN2=" + ODBC_DSN2)
  ret=createthread(codeptr(db2_worker),0)
end if
f.show
do
    if val(db1_time) >= val(db2_time) then
      last_update=db1_hhmm
      DSNUSED=ODBC_DSN1
      scr_cw=db1_cw
      scr_wt_str=right$(db1_wt_str,5)
      scr_ar=db1_ar
      scr_ta=db1_ta
      scr_lc=db1_lc
      scr_oa=db1_oa
      scr_tc=db1_tc
    else
      last_update=db2_hhmm
      DSNUSED=ODBC_DSN2
      scr_cw=db2_cw
      scr_wt_str=right$(db2_wt_str,5)
      scr_ar=db2_ar
      scr_ta=db2_ta
      scr_lc=db2_lc
      scr_oa=db2_oa
      scr_tc=db2_tc
    end if

    if showdsn=true then
      stbc=ORG + "  " + QNAME + "  " + last_update + "  " + DSNUSED
    else
      stbc=ORG + "  " + QNAME + "  " + last_update + "  " 
    end if


    ' update only on a change to prevent flicker
    if stbc <> statbar.caption then statbar.caption=stbc
    if str$(scr_cw) <> pc_a.caption then pc_a.caption=str$(scr_cw)
    if scr_wt_str <> pc_b.caption then pc_b.caption=scr_wt_str
    if scr_ar <> pc_d.caption then pc_d.caption=scr_ar
    if scr_ta <> pc_e.caption then pc_e.caption=scr_ta
    if str$(scr_lc) <> pc_c.caption then pc_c.caption=str$(scr_lc)
    if scr_oa <> pc_f.caption then pc_f.caption=scr_oa
    if str$(scr_tc) <> pc_g.caption then pc_g.caption=str$(scr_tc)

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
  defstr cline
  defint ptr
  do
    if fclose=1 then exit sub
    db1.command(QUERYCMD)
    if db1.error=0 and db1.fieldcount=FIELDCNT and db1.row<>100 then
     'begin thread
      'cline=db1.rowvalue(1,1): ptr=@cline: push ptr: gosub tval: pop db1_time
      'logevent("DB1: RV=" + db1.rowvalue(1,1) + " ET=" + str$(db1_time))
      db1_time=db1.rowvalue(1,1)
      cline=db1.rowvalue(2,1): ptr=@cline: push ptr: gosub tval: pop db1_cw
      db1_wt_str=db1.rowvalue(3,1) 
      db1_ar=db1.rowvalue(4,1)
      db1_ta=db1.rowvalue(5,1)
      cline=db1.rowvalue(6,1): ptr=@cline: push ptr: gosub tval: pop db1_lc
      cline=db1.rowvalue(7,1): ptr=@cline: push ptr: gosub tval: pop db1_wt
      db1_oa=db1.rowvalue(8,1) 
      cline=db1.rowvalue(9,1): ptr=@cline: push ptr: gosub tval: pop db1_tc
      db1_hhmm=db1.rowvalue(10,1)
      db1_wtnew=0
      db1_wt_strnew=""
      ' if multiqueue there should be more data!
      while db1.row<>100
        cline=db1.rowvalue(2,1): ptr=@cline: push ptr: gosub tval: pop db1_cw_m: db1_cw=db1_cw+db1_cw_m
        db1_wt_strnew=db1.rowvalue(3,1)
        cline=db1.rowvalue(6,1): ptr=@cline: push ptr: gosub tval: pop db1_lc_m: db1_lc=db1_lc+db1_lc_m
        cline=db1.rowvalue(7,1): ptr=@cline: push ptr: gosub tval: pop db1_wtnew
        cline=db1.rowvalue(9,1): ptr=@cline: push ptr: gosub tval: pop db1_tc_m: db1_tc=db1_tc+db1_tc_m
        if db1_wtnew > db1_wt then 
          db1_wt=db1_wtnew
          db1_wt_str=db1_wt_strnew
        end if
      end while
      logevent("DB1: OK Time=" + db1_time + " err=" + str$(db1.error))
    else
      logevent("DB1: ERR Reconnecting... ERR=" + str$(db1.error) + " FC=" + str$(db1.fieldcount))
      db1.connect(ODBC_DSN1,ODBC_USERNAME,ODBC_PASSWORD)
    end if
    db1.freememory: 
    doevents: sleep ivl
  loop
end sub

sub db2_worker()
  defstr cline
  defint ptr
  do
    if fclose=1 then exit sub
    db2.command(QUERYCMD)
    if db2.error=0 and db2.fieldcount=FIELDCNT and db2.row<>100 then
      'cline=db2.rowvalue(1,1): ptr=@cline: push ptr: gosub tval: pop db2_time
      db2_time=db2.rowvalue(1,1)
      cline=db2.rowvalue(2,1): ptr=@cline: push ptr: gosub tval: pop db2_cw
      db2_wt_str=db2.rowvalue(3,1) 
      db2_ar=db2.rowvalue(4,1)
      db2_ta=db2.rowvalue(5,1)
      cline=db2.rowvalue(6,1): ptr=@cline: push ptr: gosub tval: pop db2_lc
      cline=db2.rowvalue(7,1): ptr=@cline: push ptr: gosub tval: pop db2_wt
      db2_oa=db2.rowvalue(8,1) 
      cline=db2.rowvalue(9,1): ptr=@cline: push ptr: gosub tval: pop db2_tc
      db2_hhmm=db2.rowvalue(10,1)
      db2_wtnew=0
      db2_wt_strnew=""
      ' if multiqueue there should be more data!
      while db2.row<>100
        cline=db2.rowvalue(2,1): ptr=@cline: push ptr: gosub tval: pop db2_cw_m: db2_cw=db2_cw+db2_cw_m
        db2_wt_strnew=db2.rowvalue(3,1)
        cline=db2.rowvalue(6,1): ptr=@cline: push ptr: gosub tval: pop db2_lc_m: db2_lc=db2_lc+db2_lc_m
        cline=db2.rowvalue(7,1): ptr=@cline: push ptr: gosub tval: pop db2_wtnew
        cline=db2.rowvalue(9,1): ptr=@cline: push ptr: gosub tval: pop db2_tc_m: db2_tc=db2_tc+db2_tc_m
        if db2_wtnew > db2_wt then 
          db2_wt=db2_wtnew
          db2_wt_str=db2_wt_strnew
        end if
      end while
      logevent("DB2: OK Time=" + db2_time + " err=" + str$(db2.error))
    else
      logevent("DB2: ERR Reconnecting... ERR=" + str$(db2.error) + " FC=" + str$(db2.fieldcount))
      db2.connect(ODBC_DSN2,ODBC_USERNAME,ODBC_PASSWORD)
    end if
    db2.freememory: 
    doevents: sleep ivl
  loop
end sub




sub DoBlink()
  if sla_waitqueue>0 and scr_cw>=sla_waitqueue then
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

  if sla_waittime>0 and scr_wt>=(sla_waittime*1000) then
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

sub print_screen()
    logevent("SCR: DSN="+DSNUSED+" TIME="+last_update+" WQ="+str$(scr_cw)+" WT="+scr_wt_str+" AR="+scr_ar+ _
             " AT="+scr_ta+" LC="+str$(scr_lc)+" OA="+scr_oa+" TC="+str$(scr_tc))
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

tval:
  begin runonce: defstr v$: end runonce
  begin thread
    v$=byref$(stack(1))
    stack(1)=val(v$)
  end thread
return

msgcapture:
  if uMsg<>&H20 then
    logevent(HEX$(hWnd)+space+HEX$(uMsg)+space+HEX$(wParam)+space+HEX$(lParam) )
  end if
  retval zero
return


PROP.FILEVERSION 3,7,3,0
PROP.PRODUCTVERSION 3,7,3,0
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
PROP.VALUE "Author","Antoni Sawicki <as@tenoware.com>"
PROP.VALUE "FileDescription", "Cisco UCCX Wallboard 3.7.3"
PROP.VALUE "FileVersion", "3.7.3.0" 
PROP.VALUE "LegalCopyright", "(c) 2009-2016 by Antoni Sawicki"
PROP.VALUE "Additional Credits", "Aaron Harrison <aaro.harrison@ipcommute.co.uk>, Greg Markiewicz <gregm@bootstrap.ie>"
PROP.END  
PROP.END  
PROP.END  
