#!/usr/bin/env python3

import os
import re
import sys
import time
#import traceback
import urllib.request
from datetime import datetime
from datetime import date

import calendar
import locale
#import gettext
import tempfile
import markdown
import configparser

# TODO: handle argv[0] to select action from ini file
#print(sys.argv)

# sft0c allows referencing the strftime OS-specific modifier to remove leading 0
config_defaults = {'sft0c': '-' if os.sep == '/' else '#'}
config = configparser.ConfigParser(
    defaults = config_defaults,
    interpolation = configparser.ExtendedInterpolation()
    )

#TODO: Get settings path from Windows
# Maybe use https://pypi.org/project/config-path/ ? not packaged on Debian.
for loc in os.curdir, os.environ.get("XDG_CONFIG_HOME"), os.path.join(os.path.expanduser("~"), ".config"):
    if loc is None:
        continue
    try:
        with open(os.path.join(loc,"ics2dtp.ini")) as source:
            print("Found ini file in %s" % loc)
            config.read_file( source )
            break
    except IOError:
        pass
# DEBUG:
#print(config.sections())
#print(config['templates']['foo'])
#print(list(config['categories'].keys()).remove("sft0c"))
#print(list(config['categories'].keys()))
#filter out defaults: list(filter(lambda x: x not in config_defaults.keys(), k))
#print(config['categories']['DIY'])
#print(config['JEUX']['preamble'])

#print(getmodule())

# XXX: modules/scripts should not do this:
#use current locale setting
#locale.setlocale(locale.LC_ALL, None)
#TODO: test for Win stuff
#print("LC:%s" % str(locale.getlocale(locale.LC_MESSAGES)))

# TODO: use gettext? maybe a local DictTranslations class to avoid installing mo files
# For now we'll just wrap around English messages.
def _(m):
    return m

#print(f'{calendar.month_name[2]=}')
#print(f'{calendar.day_name[0]=}')

def _monthName(i):
    # TODO: check on Windows
    return calendar.month_name[i]
    #return locale.nl_langinfo(locale.MON_1+i)

def _dayName(i):
    # TODO: check on Windows
    return calendar.day_name[i]
    #return locale.nl_langinfo(locale.DAY_1+(0+1)%7)


try:
    import uno
    from com.sun.star.text.ControlCharacter import LINE_BREAK
    from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
    from com.sun.star.beans import PropertyValue
    from com.sun.star.lang import XMain
    from com.sun.star.script.provider import XScript
    #import screen_io as ui
except (ImportError,NameError) as err:
    #print("No LO API")
    pass

# base class for DTP scripting APIs
class DTPInterface:
    #def __init__:
    pass

# LibreOffice scripting references:
# https://wiki.openoffice.org/wiki/Python_as_a_macro_language
# http://stackoverflow.com/questions/21413664/how-to-run-python-macros-in-libreoffice
# https://tmtlakmal.wordpress.com/2013/08/11/a-simple-python-macro-in-libreoffice-4-0/
# http://hydrogeotools.blogspot.fr/2014/03/libreoffice-and-python-macros.html
# http://openoffice3.web.fc2.com/Python_Macro_General_No6.html

class LibreOfficeInterface(DTPInterface):
    def __init__(self):
        self.model = XSCRIPTCONTEXT.getDocument()
        self.controller = self.model.getCurrentController()
        self.undos = self.model.getUndoManager()
        self.status = self.controller.getFrame().createStatusIndicator()
        #print("status = %s" % str(self.status))
        self.lastStatus = ""
    pass

    # unused
    def _getScript(self, script: str, library='_Basic', module='devTools') -> XScript:
        from com.sun.star.script.provider import XScript
        sm = uno.getComponentContext().ServiceManager
        mspf = sm.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", uno.getComponentContext())
        scriptPro = mspf.createScriptProvider("")
        scriptName = "vnd.sun.star.script:"+library+"."+module+"."+script+"?language=Basic&location=application"
        xScript = scriptPro.getScript(scriptName)
        return xScript

    # taken from https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python
    # I'd think there was a ready-made version but no.
    def _inputbox(self, message, title="", default="", x=None, y=None):
        """ Shows dialog with input box.
            @param message message to show on the dialog
            @param title window title
            @param default default value
            @param x dialog positio in twips, pass y also
            @param y dialog position in twips, pass y also
            @return string if OK button pushed, otherwise zero length string
        """
        from com.sun.star.awt.MessageBoxType import \
            MESSAGEBOX, \
            INFOBOX, \
            WARNINGBOX, \
            ERRORBOX, \
            QUERYBOX
        from com.sun.star.awt.MessageBoxButtons import \
            BUTTONS_OK, \
            BUTTONS_OK_CANCEL, \
            BUTTONS_YES_NO, \
            BUTTONS_YES_NO_CANCEL, \
            BUTTONS_RETRY_CANCEL, \
            BUTTONS_ABORT_IGNORE_RETRY, \
            DEFAULT_BUTTON_OK, \
            DEFAULT_BUTTON_CANCEL, \
            DEFAULT_BUTTON_RETRY, \
            DEFAULT_BUTTON_YES, \
            DEFAULT_BUTTON_NO, \
            DEFAULT_BUTTON_IGNORE
        from com.sun.star.awt.MessageBoxResults import \
            CANCEL, OK, YES, NO, RETRY, IGNORE

        from com.sun.star.awt.PosSize import POS, SIZE, POSSIZE
        from com.sun.star.awt.PushButtonType import OK, CANCEL
        from com.sun.star.util.MeasureUnit import TWIP
        WIDTH = 600
        HORI_MARGIN = VERT_MARGIN = 8
        BUTTON_WIDTH = 100
        BUTTON_HEIGHT = 26
        HORI_SEP = VERT_SEP = 8
        LABEL_HEIGHT = BUTTON_HEIGHT * 2 + 5
        EDIT_HEIGHT = 24
        HEIGHT = VERT_MARGIN * 2 + LABEL_HEIGHT + VERT_SEP + EDIT_HEIGHT
        ctx = uno.getComponentContext()
        def create(name):
            return ctx.getServiceManager().createInstanceWithContext(name, ctx)
        dialog = create("com.sun.star.awt.UnoControlDialog")
        dialog_model = create("com.sun.star.awt.UnoControlDialogModel")
        dialog.setModel(dialog_model)
        dialog.setVisible(False)
        dialog.setTitle(title)
        dialog.setPosSize(0, 0, WIDTH, HEIGHT, SIZE)
        def add(name, type, x_, y_, width_, height_, props):
            model = dialog_model.createInstance("com.sun.star.awt.UnoControl" + type + "Model")
            dialog_model.insertByName(name, model)
            control = dialog.getControl(name)
            control.setPosSize(x_, y_, width_, height_, POSSIZE)
            for key, value in props.items():
                setattr(model, key, value)
        label_width = WIDTH - BUTTON_WIDTH - HORI_SEP - HORI_MARGIN * 2
        add("label", "FixedText", HORI_MARGIN, VERT_MARGIN, label_width, LABEL_HEIGHT, 
            {"Label": str(message), "NoLabel": True})
        add("btn_ok", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN, 
                BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": OK, "DefaultButton": True})
        add("btn_cancel", "Button", HORI_MARGIN + label_width + HORI_SEP, VERT_MARGIN + BUTTON_HEIGHT + 5, 
                BUTTON_WIDTH, BUTTON_HEIGHT, {"PushButtonType": CANCEL})
        add("edit", "Edit", HORI_MARGIN, LABEL_HEIGHT + VERT_MARGIN + VERT_SEP, 
                WIDTH - HORI_MARGIN * 2, EDIT_HEIGHT, {"Text": str(default)})
        frame = create("com.sun.star.frame.Desktop").getCurrentFrame()
        window = frame.getContainerWindow() if frame else None
        dialog.createPeer(create("com.sun.star.awt.Toolkit"), window)
        if not x is None and not y is None:
            ps = dialog.convertSizeToPixel(uno.createUnoStruct("com.sun.star.awt.Size", x, y), TWIP)
            _x, _y = ps.Width, ps.Height
        elif window:
            ps = window.getPosSize()
            _x = ps.Width / 2 - WIDTH / 2
            _y = ps.Height / 2 - HEIGHT / 2
        dialog.setPosSize(_x, _y, 0, 0, POS)
        edit = dialog.getControl("edit")
        edit.setSelection(uno.createUnoStruct("com.sun.star.awt.Selection", 0, len(str(default))))
        edit.setFocus()
        ret = edit.getModel().Text if dialog.execute() else ""
        dialog.dispose()
        return ret

    # TODO
    def InsertText(self, t):
        pass

    def insertHtmlText(self, html, frame):
        from com.sun.star.beans import PropertyValue
        # https://ask.libreoffice.org/t/paste-html-content-using-api/89749/4
        #def InsertHtml2Odt_3(sHTML, doc):
        local = uno.getComponentContext()
        ctx = local # not sure we need to resolve external context
        oStream = ctx.ServiceManager.createInstanceWithContext("com.sun.star.io.SequenceInputStream", ctx)
        oStream.initialize((uno.ByteSequence(html.encode()),))

        prop1 = PropertyValue()
        prop1.Name  = "FilterName"
        prop1.Value = "HTML (StarWriter)"
        prop2 = PropertyValue()
        prop2.Name = "InputStream"
        prop2.Value = oStream

        self.model.Text.createTextCursor().insertDocumentFromURL("private:stream", (prop1, prop2))


    def statusMessage(self, s):
        self.lastStatus = s
        self.status.setText(s)

    def progressReset(self):
        self.status.reset()
    def progressTotal(self, t):
        self.status.start(self.lastStatus, t-1) # XXX
    def progressSet(self, p):
        self.status.setValue(p)
    def progressEnd(self):
        self.status.end()

    def enterUndoContext(self, name):
        self.undos.enterUndoContext(name)
    def leaveUndoContext(self):
        self.undos.leaveUndoContext()

    def valueDialog(self, caption: str, message='LibreOffice', defaultvalue = '') -> str:
        #import screen_io as ui
        #import msgbox as ui
        #reply = ui.InputBox(message, title=caption, defaultValue=defaultValue)
        reply = self._inputbox(message,caption,defaultvalue)
        #xScript = self._getScript("_InputBox")
        #res = xScript.invoke((prompt,title,defaultValue), (), ())
        #return res[0]
        return reply



# Scribus scripting references:
# https://wiki.scribus.net/canvas/Category:Scripts
# https://wiki.scribus.net/canvas/Beginners_Scripts
# https://wiki.scribus.net/canvas/Automatic_Scripter_Commands_list
# https://scribus-scripter.readthedocs.io/en/latest/

# frame jump seems to be 1a 1b 0a
# the line breaks inside same paragraph is:
# U+2028	e2 80 a8	LINE SEPARATOR

class ScribusInterface(DTPInterface):
    def __init__(self):
        self.scribus = scribus
        print(f'{scribus.getGuiLanguage()=}')

    # TODO
    def InsertText(self, t):
        pass

    def insertHtmlText(self, html, frame):
        # TODO: check Win reopen
        # https://docs.python.org/3/library/tempfile.html#tempfile.NamedTemporaryFile
        with tempfile.NamedTemporaryFile(suffix=".html") as f:
            f.write(html.encode())
            f.flush()
            self.scribus.insertHtmlText(f.name, frame)

    def statusMessage(self, s):
        scribus.statusMessage(s)

    def progressReset(self):
        scribus.progressReset()
    def progressTotal(self, t):
        scribus.progressTotal(t)
    def progressSet(self, p):
        scribus.progressSet(int(p))
    def progressEnd(self):
        scribus.progressReset()

    def enterUndoContext(self, name):
        #TODO: unsupported yet in Scribus
        pass
    def leaveUndoContext(self):
        #TODO: unsupported yet in Scribus
        scribus.docChanged(True)
        pass

    def valueDialog(self, caption: str, message='Scribus', defaultvalue = '') -> str:
            return scribus.valueDialog(caption, message, defaultvalue)

dtp = None

try:
    import scribus
    dtp = ScribusInterface()
except ImportError as err:
    #print(_('This Python script is written for the Scribus scripting interface.'))
    #print(_('It can only be run from within Scribus.'))
    print(_('Cannot access the Scribus scripting interface.'))
    #sys.exit(1)
    try:
        #import uno
        #from com.sun.star.text.ControlCharacter import LINE_BREAK
        #from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
        #from com.sun.star.beans import PropertyValue
        #from com.sun.star.lang import XMain
        #from com.sun.star.script.provider import XScript
        #import screen_io as ui
        # We can actually import all these even from outside LibreOffice,
        # so try to access the scripting entry point
        XSCRIPTCONTEXT.getDocument()
        dtp = LibreOfficeInterface()
    except (ImportError,NameError) as err:
        #pass
        print(str(err))
        print(_('Cannot access the LibreOffice scripting interface.'))


if dtp is None:
    print(_('This Python script is written for the LibreOffice or Scribus scripting interface.'))
    print(_('It can only be run from within either of these programs.'))
    sys.exit(1)

try:
    # icalendar module is not always installed
    import icalendar
    import recurring_ical_events
    import pytz
    from pytz import timezone
except ImportError as err:
    dtp.statusMessage(': %s' % (str(err)))
    print("Except:%s\n" % (str(sys.exc_info())))
    raise


# TODO: move to ini file
agenda_text_block = "Agenda"
desc_text_block = "Description"


def OpenICalendar():
    """Prints the string 'Hello World(in Python)' into the current document"""
    #get the doc from the scripting context which is made available to all scripts

    statusDone = 0
    statusMax = 100
    dtp.progressReset()
    events = []
    url = None
    try:
        #ctx = uno.getComponentContext()
        #serviceManager = ctx.ServiceManager
        #filePicker = serviceManager.createInstance('com.sun.star.ui.dialogs.FilePicker')
        #filePicker.appendFilter("iCalendar Files (*.ics)", "*.ics")
        #url = "https://..."
        url = config['source']['url']
        dtp.statusMessage('Requesting URL: %s' % url)
        if url is not None:
            oAccept = 1
        else:
            oAccept = filePicker.execute()
        if oAccept == 1:
            if url is not None:
                oFiles = [url]
            else:
                oFiles = filePicker.getFiles()
            statusMax = 100 * len(oFiles)
            dtp.statusMessage(_('Opening...'))
            dtp.progressTotal(statusMax+1)
            dtp.progressSet(0)
            for url in oFiles:
                dtp.statusMessage('Processing: ' + url )
                dtp.progressSet(statusDone)
                tz = None
                print(f'{url=}')
                with urllib.request.urlopen(url) as f:
                    #print(f)
                    data = f.read()
                    dtp.progressSet(int(statusDone+50/2*len(oFiles)))
                    cal = icalendar.Calendar.from_ical(data)
                    import json
                    #print(f'{str(cal.walk())=}')
                    #recurring_ical_events.of(calendar).
                    for comp in cal.walk():
                        print(f'{comp.name=}')
                        if comp.name == 'VTIMEZONE':
                            if 'TZID' in comp:
                                tz = timezone(comp['TZID'])
                            print(f'{comp=}')
                        elif comp.name == 'VEVENT':
                            start = comp.decoded('DTSTART')
                            end = comp.decoded('DTEND')
                            print("Start: %s %s" % (start, type(start)))
                            if hasattr(start, 'tzinfo'):
                                if start.tzinfo == pytz.UTC:
                                    start = start.astimezone(tz)
                                print(f'{type(start.tzinfo)=}')
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
                            print("Skipping %s object" % comp.name)
                    #try:
                    #    for comp in cal.walk():
                    #        print comp.name
                    #except:
                    #    raise
                    #    print component
                    #print(cal.__class__)
                    #print(cal.property_items())
                    statusDone += 100/len(oFiles)
                    #status.setValue(statusDone)
                    dtp.progressSet(int(statusDone))

            # sort events in place
            # events.sort(key=lambda e: e.start)
            # dtp.enterUndoContext( 'Insert iCalendar' )
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
            # dtp.leaveUndoContext()
            dtp.progressEnd()
        else:
            print("cancelled!")
    except:
        dtp.statusMessage('XXX: %s' % (str(sys.exc_info())))
        print("Except:%s\n" % (str(sys.exc_info())))
        #status.setValue(-1)
        #status.end()
        dtp.statusMessage(_('Aborted'))
        dtp.progressEnd()
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

    frame = agenda_text_block
    #frame = "Description_test" # XXX:test

    #ret = dtp.valueDialog("Date range", "which date range?", "2023-09-01:2023-09-29")
    #print("DR: %s" % ret)

    events = OpenICalendar()
    
    #model = XSCRIPTCONTEXT.getDocument()
    #controller = model.getCurrentController()
    #status = controller.getFrame().createStatusIndicator()
    statusDone = 0
    statusMax = 100
    dtp.progressReset()
    dtp.progressTotal(statusMax+1)
    dtp.progressSet(0)
    dtp.statusMessage(_('Inserting'))
    # this puts the text at start of document
    #text = model.Text
    #cursor = text.createTextCursor()
    # this puts the text at current insertion point
    #text = controller.getViewCursor().getText()
    #cursor = text.createTextCursorByRange(controller.getViewCursor().getStart())
    
    # sort events in place
    events.sort(key=lambda e: e.start)
    dtp.enterUndoContext( _('Insert iCalendar') )
    md = ""
    for comp in events:
        start = comp.start
        end = comp.end
        # TODO: check Windows: %- -> %# ?
        # https://strftime.org/
        dfmt = "%A %-d %B"
        tfmt = "%-Hh%-M"
        print("XXX:{event.start:%A %-d %B} {event.start:%-Hh%M %-X}//{event.end.minute:d}\u2192{event.end:%-Hh%-0M}".format(event=comp))
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
        md += "\u2192%s  \n" % end.strftime(tfmt)
        md += "Test\n"
        if 'DESCRIPTION' in comp:
            #text.insertString( cursor, "%s\n" % comp['DESCRIPTION'], 0 )
            #scribus.insertText("%s\n" % comp['DESCRIPTION'], -1, frame)
            md += "\n%s" % comp['DESCRIPTION']
            
        statusDone += 50/len(events)
        dtp.progressSet(int(statusDone))
        print(statusDone)
        md += "\n"

    #dtp.progressSet(min(int(statusDone),99))
    print(md)
    # TODO: h1,h2 -> span class= ?
    #f.write(("<html><head><meta encoding='UTF-8'></head><body>%s</body></html>" % str(markdown.markdown(md))))
    html = markdown.markdown(md)
    #h2_style = "04 TITRES CATÉGORIES BLANC SUR COULEUR"
    h2_style = "Edito"
    html = re.sub('<h2>(.*)</h2>', '<p style="%s">\\1</p>' % h2_style, html)
    html = '<?xml version="1.0" encoding="utf-8"?><html><head></head><body>%s</body></html>\n' % html
    print(html)
    dtp.insertHtmlText(html, frame)
    #print(scribus.getPosition())
    #scribus.insertText("Foo\nbar\n\ntoto", -1, frame)
    #print(scribus.getPosition())
    dtp.leaveUndoContext()
    dtp.progressEnd()

# TODO: FIXME
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
    undos.enterUndoContext( _('Insert iCalendar') )
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

# LibreOffice: export only these to the UI

g_exportedScripts = InsertICalendar, InsertICalendarTimeTable,

# Scribus: handle script launch

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
        dtp.statusMessage(_('Running script...'))
        dtp.progressReset()
        main(argv)
    finally:
        # Exit neatly even if the script terminated with an exception,
        # so we leave the progress bar and status bar blank and make sure
        # drawing is enabled.
        if scribus.haveDoc() > 0:
            scribus.setRedraw(True)
        dtp.statusMessage('')
        dtp.progressReset()

# This code detects if the script is being run as a script, or imported as a module.
# It only runs main() if being run as a script. This permits you to import your script
# and control it manually for debugging.
if __name__ == '__main__':
    main_wrapper(sys.argv)

# vim: set shiftwidth=4 softtabstop=4 expandtab:
