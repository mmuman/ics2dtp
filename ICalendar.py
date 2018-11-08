###!/usr/bin/python
# HelloWorld python script for the scripting framework

import uno
import sys
import time
#import traceback
import urllib.request
from datetime import datetime
from datetime import date

from com.sun.star.text.ControlCharacter import LINE_BREAK
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK

from com.sun.star.lang import XMain

# references:
# https://wiki.openoffice.org/wiki/Python_as_a_macro_language
# http://stackoverflow.com/questions/21413664/how-to-run-python-macros-in-libreoffice
# https://tmtlakmal.wordpress.com/2013/08/11/a-simple-python-macro-in-libreoffice-4-0/
# http://hydrogeotools.blogspot.fr/2014/03/libreoffice-and-python-macros.html
# http://openoffice3.web.fc2.com/Python_Macro_General_No6.html


def OpenICalendar():
    """Prints the string 'Hello World(in Python)' into the current document"""
    #get the doc from the scripting context which is made available to all scripts
    try:
        # icalendar module is not always installed
        import icalendar
        import pytz
        from pytz import timezone
    except:
        status.setText('Ex: %s' % (str(sys.exc_info())))
        print("Except:%s\n" % (str(sys.exc_info())))
        raise


    model = XSCRIPTCONTEXT.getDocument()
    controller = model.getCurrentController()
    status = controller.getFrame().createStatusIndicator()
    statusDone = 0
    statusMax = 100
    status.reset()
    events = []
    url = None
    try:
        ctx = uno.getComponentContext()
        serviceManager = ctx.ServiceManager
        filePicker = serviceManager.createInstance('com.sun.star.ui.dialogs.FilePicker')
        filePicker.appendFilter("iCalendar Files (*.ics)", "*.ics")
        status.setText('XXX: %s' % url)
        if url is not None:
            oAccept = 1
        else:
            oAccept = filePicker.execute()
        if oAccept == 1:
            if url is not None:
                oFiles = [url]
            else:
                oFiles = filePicker.getFiles()
            #statusMax = 100 * len(oFiles)
            status.start('Opening', statusMax)
            status.setValue(0)
            for url in oFiles:
                status.setText('Processing: ' + url )
                status.setValue(statusDone)
                tz = None
                print(url)
                with urllib.request.urlopen(url) as f:
                    #print(f)
                    data = f.read()
                    status.setValue(statusDone+50/2*len(oFiles))
                    cal = icalendar.Calendar.from_ical(data)
                    for comp in cal.walk():
                        print(comp.name)
                        if comp.name == 'VTIMEZONE':
                            if 'TZID' in comp:
                                tz = timezone(comp['TZID'])
                            print(comp)
                        elif comp.name == 'VEVENT':
                            start = comp.decoded('DTSTART')
                            end = comp.decoded('DTEND')
                            print("Start: %s %s" % (start, type(start)))
                            if hasattr(start, 'tzinfo'):
                                if start.tzinfo == pytz.UTC:
                                    start = start.astimezone(tz)
                                print(type(start.tzinfo))
                            elif isinstance(start, datetime):
                                start = tz.localize(start)
                                print("START:%s" % type(start))
                            else:
                                start = tz.localize(datetime.combine(start, datetime.min.time()))
                            print("End: %s" % end)
                            if hasattr(end, 'tzinfo'):
                                print(end.tzinfo or None)
                                if end.tzinfo == pytz.UTC:
                                    end = end.astimezone(tz)
                            elif isinstance(end, datetime):
                                end = tz.localize(end)
                            else:
                                end = tz.localize(datetime.combine(end, datetime.min.time()))
                            #if hasattr(comp, 'start') or hasattr(comp, 'end'):
                            #    print("HHHHHHHHHHHHHH: %s" % comp)
                            comp.start = start
                            comp.end = end
                            events.append(comp)

                            # cursor.CharWeight = FontWeight.BOLD
                            # text.insertString( cursor, "[%s - %s]" % (start, end), 0 )
                            # cursor.CharWeight = FontWeight.NORMAL
                            # cursor.CharPosture = FontSlant.ITALIC
                            # summary = ""
                            # if 'SUMMARY' in comp:
                            #     summary = comp['SUMMARY']
                            # text.insertString( cursor, " %s\n" % summary, 0 )
                            # cursor.CharPosture = FontSlant.NONE
                            # if 'DESCRIPTION' in comp:
                            #     text.insertString( cursor, "%s\n" % comp['DESCRIPTION'], 0 )
                        else:
                            print(comp.name)
                    #try:
                    #    for comp in cal.walk():
                    #        print comp.name
                    #except:
                    #    raise
                    #    print component
                    print(cal.__class__)
                    #print(cal.property_items())
                    statusDone += 100/len(oFiles)
                    status.setValue(statusDone)

            # sort events in place
            # events.sort(key=lambda e: e.start)
            # undos.enterUndoContext( 'Insert iCalendar' ); 
            # for comp in events:
            #     start = comp.start
            #     end = comp.end
            #     cursor.CharWeight = FontWeight.BOLD
            #     text.insertString( cursor, "[%s - %s]" % (start, end), 0 )
            #     cursor.CharWeight = FontWeight.NORMAL
            #     cursor.CharPosture = FontSlant.ITALIC
            #     summary = ""
            #     if 'SUMMARY' in comp:
            #         summary = comp['SUMMARY']
            #         text.insertString( cursor, " %s\n" % summary, 0 )
            #         cursor.CharPosture = FontSlant.NONE
            #     if 'DESCRIPTION' in comp:
            #         text.insertString( cursor, "%s\n" % comp['DESCRIPTION'], 0 )
            #     statusDone += 50/len(events)
            #     status.setValue(statusDone)
            # undos.leaveUndoContext()
            status.end()
        else:
            print("cancelled!")
    except:
        status.setText('XXX: %s' % (str(sys.exc_info())))
        print("Except:%s\n" % (str(sys.exc_info())))
        status.setText('Aborted')
        status.setValue(-1)
        status.end()
        raise
    #text.insertString( cursor, "Fine %s\n" % sys.version, 0 )
    #print("Fine %s\n" % sys.version)
    return events

def InsertICalendar( ):
    #return
    class FontSlant():
        from com.sun.star.awt.FontSlant import (NONE, ITALIC,)
    class FontWeight():
        from com.sun.star.awt.FontWeight import (NORMAL, BOLD,)
    
    events = OpenICalendar()
    
    model = XSCRIPTCONTEXT.getDocument()
    controller = model.getCurrentController()
    status = controller.getFrame().createStatusIndicator()
    statusDone = 0
    statusMax = 100
    status.reset()
    status.start('Inserting', statusMax)
    status.setValue(0)
    # this puts the text at start of document
    #text = model.Text
    #cursor = text.createTextCursor()
    # this puts the text at current insertion point
    text = controller.getViewCursor().getText()
    cursor = text.createTextCursorByRange(controller.getViewCursor().getStart())
    undos = model.getUndoManager()
    
    # sort events in place
    events.sort(key=lambda e: e.start)
    undos.enterUndoContext( 'Insert iCalendar' ); 
    for comp in events:
        start = comp.start
        end = comp.end
        cursor.CharWeight = FontWeight.BOLD
        text.insertString( cursor, "[%s - %s]" % (start, end), 0 )
        cursor.CharWeight = FontWeight.NORMAL
        cursor.CharPosture = FontSlant.ITALIC
        summary = ""
        if 'SUMMARY' in comp:
            summary = comp['SUMMARY']
            text.insertString( cursor, " %s\n" % summary, 0 )
            cursor.CharPosture = FontSlant.NONE
        if 'DESCRIPTION' in comp:
            text.insertString( cursor, "%s\n" % comp['DESCRIPTION'], 0 )
        statusDone += 50/len(events)
        status.setValue(statusDone)
    undos.leaveUndoContext()
    status.end()

def InsertICalendarTimeTable( ):
    class FontSlant():
        from com.sun.star.awt.FontSlant import (NONE, ITALIC,)
    class FontWeight():
        from com.sun.star.awt.FontWeight import (NORMAL, BOLD,)

    events = OpenICalendar()

    model = XSCRIPTCONTEXT.getDocument()
    controller = model.getCurrentController()
    status = controller.getFrame().createStatusIndicator()
    statusDone = 0
    statusMax = 100
    status.reset()
    status.start('Inserting', statusMax)
    status.setValue(0)
    # this puts the text at start of document
    #text = model.Text
    #cursor = text.createTextCursor()
    # this puts the text at current insertion point
    text = controller.getViewCursor().getText()
    cursor = text.createTextCursorByRange(controller.getViewCursor().getStart())
    oldStyle = cursor.ParaStyleName
    print(oldStyle)
    undos = model.getUndoManager()
    
    # sort events in place
    events.sort(key=lambda e: e.start)
    undos.enterUndoContext( 'Insert iCalendar' );
    last_date = date.fromordinal(1)
    cursor.CharWeight = FontWeight.NORMAL
    cursor.CharPosture = FontSlant.NONE
    for comp in events:
        start = comp.start
        end = comp.end

        if start.date().month != last_date.month:
            #cursor.CharWeight = FontWeight.BOLD
            #text.insertControlCharacter(cursor, LINE_BREAK, 0)
            #cursor.collapseToEnd()
            cursor.CharWeight = FontWeight.NORMAL
            cursor.CharPosture = FontSlant.NONE
            cursor.ParaStyleName = "Heading 1"
            #cursor.collapseToEnd()
            text.insertString( cursor, "%s\r" % start.date().strftime("%B").capitalize(), 0 )
            #print("HEADING:%s" % start.date().strftime("%B"))
            #text.insertControlCharacter(cursor, PARAGRAPH_BREAK, 0)
            cursor.collapseToEnd()
            cursor.ParaStyleName = oldStyle
            #text.insertString( cursor, "\r", 0 )
            #cursor.CharWeight = FontWeight.NORMAL

        if start.date() != last_date:
            last_date = start.date()
            #print("DATE:%s" % last_date.strftime("%A"))
            cursor.CharWeight = FontWeight.BOLD
            text.insertString( cursor, "%s %d\n" % (last_date.strftime("%A"), last_date.day), 0 )
            cursor.CharWeight = FontWeight.NORMAL
        cursor.CharWeight = FontWeight.BOLD
        text.insertString( cursor, "\t%s" % start.time().strftime("%H:%M"), 0 )
        cursor.CharWeight = FontWeight.NORMAL
        cursor.CharPosture = FontSlant.ITALIC
        summary = ""
        if 'SUMMARY' in comp:
            summary = comp['SUMMARY']
            text.insertString( cursor, " %s\r" % summary, 0 )
            cursor.CharPosture = FontSlant.NONE
        #if 'DESCRIPTION' in comp:
        #    text.insertString( cursor, "%s\n" % comp['DESCRIPTION'], 0 )
        statusDone += 50/len(events)
        status.setValue(statusDone)
    undos.leaveUndoContext()
    status.end()

g_exportedScripts = InsertICalendar, InsertICalendarTimeTable,

