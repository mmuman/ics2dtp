#!/usr/bin/env python3

import re
import sys
import time
#import traceback
import urllib.request
from datetime import datetime
from datetime import date

#import locale
#import gettext
import tempfile
import markdown
import configparser

config = configparser.ConfigParser()
config.sections()
#TODO: Get settings path from OS
config.read('ics2dtp.ini')
print(config.sections())

# XXX: modules/scripts should not do this:
#use current locale setting
#locale.setlocale(locale.LC_ALL, None)
#TODO: test for Win stuff
#print(locale.getlocale(locale.LC_MESSAGES))

# TODO: use gettext? maybe a local DictTranslations class to avoid installing mo files
# For now we'll just wrap around English messages.
def _(m):
    return m

# try:
#     import uno
#     from com.sun.star.text.ControlCharacter import LINE_BREAK
#     from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
#     from com.sun.star.lang import XMain
# except ImportError as err:
#     print (_('Cannot import the LibreOffice scripting interface.'))
#     # TODO

try:
    import scribus
except ImportError as err:
    print (_('This Python script is written for the Scribus scripting interface.'))
    print (_('It can only be run from within Scribus.'))
    sys.exit(1)

# LibreOffice scripting references:
# https://wiki.openoffice.org/wiki/Python_as_a_macro_language
# http://stackoverflow.com/questions/21413664/how-to-run-python-macros-in-libreoffice
# https://tmtlakmal.wordpress.com/2013/08/11/a-simple-python-macro-in-libreoffice-4-0/
# http://hydrogeotools.blogspot.fr/2014/03/libreoffice-and-python-macros.html
# http://openoffice3.web.fc2.com/Python_Macro_General_No6.html

# Scribus scripting references:
# https://wiki.scribus.net/canvas/Category:Scripts
# https://wiki.scribus.net/canvas/Beginners_Scripts
# https://wiki.scribus.net/canvas/Automatic_Scripter_Commands_list
# https://scribus-scripter.readthedocs.io/en/latest/

agenda_text_block = "Agenda"
desc_text_block = "Description"


# https://ask.libreoffice.org/t/paste-html-content-using-api/89749/4
def InsertHtml2Odt_3(sHTML, doc):
    oStream = ctx.ServiceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", ctx)
    oStream.initialize((uno.ByteSequence(sHTML.encode()),))

    prop1 = PropertyValue()
    prop1.Name  = "FilterName"
    prop1.Value = "HTML (StarWriter)"
    prop2 = PropertyValue()
    prop2.Name = "InputStream" 
    prop2.Value = oStream
    
    doc.Text.createTextCursor().insertDocumentFromURL("private:stream", (prop1, prop2))


def OpenICalendar():
    """Prints the string 'Hello World(in Python)' into the current document"""
    #get the doc from the scripting context which is made available to all scripts
    try:
        # icalendar module is not always installed
        import icalendar
        import pytz
        from pytz import timezone
    except:
        #status.setText('Ex: %s' % (str(sys.exc_info())))
        scribus.statusMessage('Ex: %s' % (str(sys.exc_info())))
        print("Except:%s\n" % (str(sys.exc_info())))
        raise

    #model = XSCRIPTCONTEXT.getDocument()
    #controller = model.getCurrentController()
    #status = controller.getFrame().createStatusIndicator()
    statusDone = 0
    statusMax = 100
    #status.reset()
    scribus.progressReset()
    events = []
    url = None
    try:
        #ctx = uno.getComponentContext()
        #serviceManager = ctx.ServiceManager
        #filePicker = serviceManager.createInstance('com.sun.star.ui.dialogs.FilePicker')
        #filePicker.appendFilter("iCalendar Files (*.ics)", "*.ics")
        #url = "https://..."
        url = config['source']['url']
        #status.setText('XXX: %s' % url)
        scribus.statusMessage('Requesting URL: %s' % url)
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
            #status.start('Opening', statusMax)
            #status.setValue(0)
            scribus.statusMessage(_('Opening...'))
            scribus.progressTotal(100)
            scribus.progressSet(0)
            for url in oFiles:
                #status.setText('Processing: ' + url )
                #status.setValue(statusDone)
                scribus.statusMessage('Processing: ' + url )
                scribus.progressSet(statusDone)
                tz = None
                print(url)
                with urllib.request.urlopen(url) as f:
                    #print(f)
                    data = f.read()
                    #status.setValue(statusDone+50/2*len(oFiles))
                    scribus.progressSet(int(statusDone+50/2*len(oFiles)))
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
                    #status.setValue(statusDone)
                    scribus.progressSet(int(statusDone))

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
            #status.end()
            scribus.progressReset()
        else:
            print("cancelled!")
    except:
        #status.setText('XXX: %s' % (str(sys.exc_info())))
        scribus.statusMessage('XXX: %s' % (str(sys.exc_info())))
        print("Except:%s\n" % (str(sys.exc_info())))
        #status.setText(_('Aborted'))
        #status.setValue(-1)
        #status.end()
        scribus.statusMessage(_('Aborted'))
        scribus.progressReset()
        raise
    #text.insertString( cursor, "Fine %s\n" % sys.version, 0 )
    #print("Fine %s\n" % sys.version)
    return events



def InsertICalendar( ):
    #return
    
    # class FontSlant():
    #     from com.sun.star.awt.FontSlant import (NONE, ITALIC,)
    # class FontWeight():
    #     from com.sun.star.awt.FontWeight import (NORMAL, BOLD,)

    #frame = agenda_text_block
    frame = "Description_test" # XXX:test

    events = OpenICalendar()
    
    #model = XSCRIPTCONTEXT.getDocument()
    #controller = model.getCurrentController()
    #status = controller.getFrame().createStatusIndicator()
    statusDone = 0
    statusMax = 100
    #status.reset()
    #status.start('Inserting', statusMax)
    #status.setValue(0)
    scribus.progressReset()
    scribus.progressTotal(100)
    scribus.progressSet(0)
    scribus.statusMessage(_('Inserting'))
    # this puts the text at start of document
    #text = model.Text
    #cursor = text.createTextCursor()
    # this puts the text at current insertion point
    #text = controller.getViewCursor().getText()
    #cursor = text.createTextCursorByRange(controller.getViewCursor().getStart())
    #undos = model.getUndoManager()
    
    # sort events in place
    events.sort(key=lambda e: e.start)
    #undos.enterUndoContext( 'Insert iCalendar' ); 
    md = ""
    for comp in events:
        start = comp.start
        end = comp.end
        # TODO: check Windows: %- -> %# ?
        # https://strftime.org/
        dfmt = "%A %-d %B"
        tfmt = "%-Hh%-M"
        #cursor.CharWeight = FontWeight.BOLD
        #text.insertString( cursor, "[%s - %s]" % (start, end), 0 )
        #scribus.insertText("[%s - %s]" % (start, end), -1, frame)
        #cursor.CharWeight = FontWeight.NORMAL
        #cursor.CharPosture = FontSlant.ITALIC
        summary = ""
        if 'SUMMARY' in comp:
            summary = comp['SUMMARY']
            #text.insertString( cursor, " %s\n" % summary, 0 )
            md += "## %s\n" % summary
            #scribus.insertText(" %s\n" % summary, -1, frame)
            # TODO: markdown.markdown()
            #scribus.insertHtmlText("/home/revol/68k_news.html", frame)
            #cursor.CharPosture = FontSlant.NONE
        md += "**%s**" % start.strftime(dfmt)
        md += " %s" % start.strftime(tfmt)
        md += "\u2192%s\n" % end.strftime(tfmt)
        if 'DESCRIPTION' in comp:
            #text.insertString( cursor, "%s\n" % comp['DESCRIPTION'], 0 )
            #scribus.insertText("%s\n" % comp['DESCRIPTION'], -1, frame)
            md += "\n%s" % comp['DESCRIPTION']
            
        statusDone += 50/len(events)
        #status.setValue(statusDone)
        print(statusDone)
        md += "\n"

    #scribus.progressSet(min(int(statusDone),99))
    # TODO: check Win reopen
    # https://docs.python.org/3/library/tempfile.html#tempfile.NamedTemporaryFile
    with tempfile.NamedTemporaryFile(suffix=".html") as f:
        print(md)
        # TODO: h1,h2 -> span class= ?
        #f.write(("<html><head><meta encoding='UTF-8'></head><body>%s</body></html>" % str(markdown.markdown(md))))
        html = markdown.markdown(md)
        #h2_style = "04 TITRES CATÉGORIES BLANC SUR COULEUR"
        h2_style = "Edito"
        html = re.sub('<h2>(.*)</h2>', '<p style="%s">\\1</p>' % h2_style, html)
        html = '<?xml version="1.0" encoding="utf-8"?><html><head></head><body>%s</body></html>\n' % html
        print(html)
        f.write(html.encode())
        f.flush()
        scribus.insertHtmlText(f.name, frame)
        #scribus.insertText(" %s\n" % markdown.markdown(md), -1, frame)
    #undos.leaveUndoContext()
    #status.end()

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
    status.start(_('Inserting'), statusMax)
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
    undos.enterUndoContext( _('Insert iCalendar') );
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
        #cursor.CharWeight = FontWeight.BOLD
        #startTime = start.time().strftime("%Hh%M").replace('h00','h')
        #text.insertString( cursor, "\t%s" % startTime, 0 )
        cursor.CharWeight = FontWeight.NORMAL
        cursor.CharPosture = FontSlant.NONE
        #cursor.CharPosture = FontSlant.ITALIC
        summary = ""
        subtitle = ""
        if 'SUMMARY' in comp:
            summary = comp['SUMMARY']
            text.insertString( cursor, "%s" % summary, 0 )
            cursor.CharPosture = FontSlant.NONE
        cursor.CharWeight = FontWeight.BOLD
        startTime = start.time().strftime("%Hh%M").replace('h00','h')
        text.insertString( cursor, " - %s" % startTime, 0 )
        cursor.CharWeight = FontWeight.NORMAL
        text.insertString( cursor, "\r", 0 )
        #if 'DESCRIPTION' in comp:
        #    text.insertString( cursor, "%s\n" % comp['DESCRIPTION'], 0 )
        statusDone += 50/len(events)
        status.setValue(statusDone)
    undos.leaveUndoContext()
    status.end()

#g_exportedScripts = InsertICalendar, InsertICalendarTimeTable,

def main(argv):
    """Application initialization, font checks and initial setup."""
    #initialisation()
    #f = scribus.createText(x, y, w, h)
    aName = "Agenda"
    dName = "Description"
    #scribus.insertText("foobar", -1, aName)
    #scribus.setFont(font, f)
    #scribus.setFontSize(fontSize, f)
    #scribus.setLineSpacing(lineSpace, f)
    #scribus.setTextAlignment(0, f)
    InsertICalendar()


def main_wrapper(argv):
    """The main_wrapper() function disables redrawing, sets a sensible generic
    status bar message, and optionally sets up the progress bar. It then runs
    the main() function. Once everything finishes it cleans up after the main()
    function, making sure everything is sane before the script terminates."""
    try:
        scribus.statusMessage(_('Running script...'))
        scribus.progressReset()
        main(argv)
    finally:
        # Exit neatly even if the script terminated with an exception,
        # so we leave the progress bar and status bar blank and make sure
        # drawing is enabled.
        if scribus.haveDoc() > 0:
            scribus.setRedraw(True)
        scribus.statusMessage('')
        scribus.progressReset()

# This code detects if the script is being run as a script, or imported as a module.
# It only runs main() if being run as a script. This permits you to import your script
# and control it manually for debugging.
if __name__ == '__main__':
    main_wrapper(sys.argv)

# vim: set shiftwidth=4 softtabstop=4 expandtab:
