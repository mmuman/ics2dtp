#!/usr/bin/env python3

import os
import re
import sys
import time
#import traceback
import urllib.request
from urllib.error import URLError
from datetime import datetime
from datetime import date

import calendar
import locale
#import gettext
import tempfile
import markdown
import configparser
import itertools

# TODO: handle argv[0] to select action from ini file
#print(sys.argv)

# fake this until we use gettext
def _(s):
    return s

#print(f'{calendar.month_name[2]=}')
#print(f'{calendar.day_name[0]=}')

def _monthName(i):
    # TODO: check on Windows
    # XXX: depends on default locale, not ini!
    return calendar.month_name[i]
    #return locale.nl_langinfo(locale.MON_1+i)

def _dayName(i):
    # TODO: check on Windows
    # XXX: depends on default locale, not ini!
    return calendar.day_name[i]
    #return locale.nl_langinfo(locale.DAY_1+(0+1)%7)


# Work around https://bugs.python.org/issue27400
def strptime(s, fmt):
    try:
        return datetime.strptime(s, fmt)
    except TypeError:
        import time
        return datetime.fromtimestamp(time.mktime(time.strptime(s, fmt)))

try:
    import uno
    from com.sun.star.text.ControlCharacter import LINE_BREAK
    from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
    from com.sun.star.beans import PropertyValue
    from com.sun.star.lang import XMain
    from com.sun.star.script.provider import XScript
    from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
    #import screen_io as ui
except (ImportError,NameError) as err:
    #print("No LO API")
    pass

# base class for DTP scripting APIs
class DTPInterface:
    #def __init__:
    pass

    def insertRestyledHtmlText(self, html, frame, styles):
        print("REST")
        if styles is None:
            return self.insertHtmlText(html, frame)
        #print(html)
        # l = scribus.getTextLength(frame)
        # scribus.selectText(l, scribus.getTextLength(frame)-l, frame)
        # #self.setParagraphStyle("04 TITRES CATÉGORIES BLANC SUR COULEUR", frame)
        # scribus.setStyle("04 TITRES CATÉGORIES BLANC SUR COULEUR", frame)
        # self.InsertText("Ceci est un test\n", frame)
        # scribus.selectText(l, scribus.getTextLength(frame)-l, frame)
        # scribus.setStyle("03 TEXTE NOIR", frame)
        # #self.setParagraphStyle("03 TEXTE NOIR", frame)
        # self.InsertText("Ceci est un test aussi\n", frame)
        # return

        for tline in html.split("<"):
            if len(tline) < 1:
                continue
            [token, text] = tline.split(">")
            print(f"T: {token}, R: {text}")

            if token[0] is '/' and len(text):
                print(f"BAD!!!! T: {token}, R: {text}")

            if token == 'br /':
                self.InsertText(text, frame)
            elif token == 'p':
                if token in styles:
                    self.setParagraphStyle(styles[token], frame)
                self.InsertText(text, frame)
            elif token == '/p':
                self.InsertText(text, frame)

            elif token == 'i':
                self.InsertText(text, frame)
                if token in styles:
                    self.setCharacterStyle(styles[token], frame)
            elif token == '/i':
                if token in styles:
                    self.setCharacterStyle(styles[token], frame)
                else:
                    self.setCharacterStyle("", frame)
                self.InsertText(text, frame)

            elif token in ['b', 'strong']:
                self.InsertText(text, frame)
            elif token in ['/b', '/strong']:
                self.InsertText(text, frame)

            elif token in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                l = token[1]
                if token in styles:
                    self.setParagraphStyle(styles[token], frame)
                self.InsertText(text, frame)
            elif token in ['/h1', '/h2', '/h3', '/h4', '/h5', '/h6']:
                l = token[2]
                if token in styles:
                    self.setParagraphStyle(styles[token], frame)
                elif 'p' in styles:
                    self.setParagraphStyle(styles['p'], frame)
                else:
                    self.setParagraphStyle('', frame)
                self.InsertText(text, frame)
            else:
                print(f'UNHANDLED T: {token}, R: {text}')
                self.InsertText(text, frame)


# LibreOffice scripting references:
# https://wiki.openoffice.org/wiki/Python_as_a_macro_language
# http://stackoverflow.com/questions/21413664/how-to-run-python-macros-in-libreoffice
# https://tmtlakmal.wordpress.com/2013/08/11/a-simple-python-macro-in-libreoffice-4-0/
# http://hydrogeotools.blogspot.fr/2014/03/libreoffice-and-python-macros.html
# http://openoffice3.web.fc2.com/Python_Macro_General_No6.html
# https://wiki.documentfoundation.org/Macros/Python_Guide/Useful_functions

class LibreOfficeInterface(DTPInterface):
    def __init__(self):
        self.model = XSCRIPTCONTEXT.getDocument()
        self.controller = self.model.getCurrentController()
        self.undos = self.model.getUndoManager()
        self.status = self.controller.getFrame().createStatusIndicator()
        #print("status = %s" % str(self.status))
        self.lastStatus = ""
        self.frame_jump = "\n" # FIXME
        self.text = self.controller.getViewCursor().getText()
        self.cursor = self.text.createTextCursorByRange(self.controller.getViewCursor().getStart())
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

    def msgbox(self, message, title, buttons=MSG_BUTTONS.BUTTONS_OK, type_msg='infobox'):
        """ Create message box
            type_msg: infobox, warningbox, errorbox, querybox, messbox
            https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1awt_1_1XMessageBoxFactory.html
        """
        def create(name):
            return ctx.getServiceManager().createInstanceWithContext(name, ctx)
        ctx = uno.getComponentContext()
        toolkit = create('com.sun.star.awt.Toolkit')
        parent = toolkit.getDesktopWindow()
        mb = toolkit.createMessageBox(parent, type_msg, buttons, title, str(message))
        return mb.execute()

    def setCharacterStyle(self, style, frame):
        pass

    def setParagraphStyle(self, style, frame):
        pass

    def InsertText(self, t, frame):
        self.text.insertString( self.cursor, t, 0 )
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

    def messageBox(self, message: str, caption='LibreOffice'):
        return self.msgbox(message, caption)

    def valueDialog(self, caption: str, message='LibreOffice', defaultvalue = '') -> str:
        #import screen_io as ui
        #import msgbox as ui
        #reply = ui.InputBox(message, title=caption, defaultValue=defaultValue)
        reply = self._inputbox(message,caption,defaultvalue)
        #xScript = self._getScript("_InputBox")
        #res = xScript.invoke((prompt,title,defaultValue), (), ())
        #return res[0]
        return reply

    # TODO: handle styles

    # class FontSlant():
    #     from com.sun.star.awt.FontSlant import (NONE, ITALIC,)
    # class FontWeight():
    #     from com.sun.star.awt.FontWeight import (NORMAL, BOLD,)


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
        self.frame_jump = "\x1a\x1b\n"

    def setCharacterStyle(self, style, frame):
        scribus.setCharacterStyle(style, frame)

    def setParagraphStyle(self, style, frame):
        scribus.setStyle(style, frame)

    def InsertText(self, t, frame):
        scribus.insertText(t, -1, frame)
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

    def messageBox(self, message: str, caption='Scribus'):
        return scribus.messageBox(caption, message)

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

config_defaults = {
    # allows referencing the strftime OS-specific modifier to remove leading 0
    'sft0c': '-' if os.sep == '/' else '#',
    'dtp_fj': dtp.frame_jump
    }
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
        # TODO
        pass
# TODO: warn if no config file / propose a default one?

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
        urls = [config['source']['url']]
        dtp.statusMessage('Requesting URL: %s' % url)
        if urls is not None:
            statusMax = 100 * len(urls)
            dtp.statusMessage(_('Opening...'))
            dtp.progressTotal(statusMax+1)
            dtp.progressSet(0)
            for url in urls:
                dtp.statusMessage('Processing: ' + url )
                dtp.progressSet(statusDone)
                tz = None
                data = None
                print(f'{url=}')
                try:
                    with urllib.request.urlopen(url) as f:
                        #print(f)
                        data = f.read()
                        dtp.progressSet(int(statusDone+50/2*len(urls)))
                except URLError as e:
                    dtp.statusMessage(_('Error fetching: ') + url)
                    dtp.messageBox(_('Error fetching: ') + url)
                    continue
                if data is not None:
                    cal = icalendar.Calendar.from_ical(data)
                    # We must walk it first to get the timezone,
                    # since recurring_ical_events eats it, and passes multi-days events without time
                    for comp in cal.walk():
                        print(comp)
                        print(f'{comp.name=}')
                        if comp.name == 'VTIMEZONE':
                            if 'TZID' in comp:
                                tz = timezone(comp['TZID'])
                            print(f'{comp=}')
                    # FIXME: pass start,end as args
                    #print(str(cal))
                    if ('period' in config['general']) and not config['general'].getboolean('confirm_period'):
                        period = config['general']['period'].split(" ")
                    else:
                        period = dtp.valueDialog(_("Period"), _("Period to load"), config.get('general', 'period', fallback="2024-01-01 2024-01-07"))
                        if len(period) < 1:
                            dtp.statusMessage(_('Cancelled'))
                            continue
                        period = period.split(" ")
                    period = [datetime.fromisoformat(p) for p in period]
                    cal = recurring_ical_events.of(cal, keep_recurrence_attributes=True).between(period[0], period[1])
                    # we get a list, not an enumerator
                    #cal = enumerate(cal)
                    for comp in cal:#cal.walk():
                    #for comp in cal.walk():
                        print(f'{comp.name=}')
                        if comp.name == 'VTIMEZONE':
                            if 'TZID' in comp:
                                tz = timezone(comp['TZID'])
                            print(f'{comp=}')
                        elif comp.name == 'VEVENT':
                            start = comp.decoded('DTSTART')
                            end = comp.decoded('DTEND')
                            start_dt = start
                            end_dt = end
                            print(comp.get('SUMMARY', 'NO SUMMARY'))
                            print("Start: %s %s" % (start, type(start)))
                            if not isinstance(start, datetime):
                                print("ISDATE")
                                start_dt = tz.localize(datetime.combine(start, datetime.min.time()))
                                print("StartDT: %s %s" % (start_dt, type(start_dt)))
                                print(comp)
                            print("End: %s" % end)
                            comp.start_dt = start_dt
                            comp.start = start
                            if not isinstance(end, datetime):
                                end_dt = tz.localize(datetime.combine(end, datetime.min.time()))
                            comp.end = end
                            comp.end_dt = end_dt

                            if 'categories_summary_match' not in config['source']:
                                # TODO
                                print("TODO: get categories from categories field")

                            if 'SUMMARY' in comp:
                                # TODO: check this!
                                if 'categories_summary_match' in config['source']:
                                    #print(config['source']['categories_summary_match'])
                                    m = re.search(config['source']['categories_summary_match'], comp['SUMMARY'])
                                    if m and len(m.group("title")):
                                        print(m.groups())
                                        # XXX: abuse this field or some other?
                                        comp['CATEGORY'] = m.group("category")
                                        comp['SUMMARY'] = m.group("title")
                                    else:
                                        print(comp)
                                        dtp.messageBox(_("{0:%c}:\nEvent '{1}' without category, skipping.").format(comp.start_dt, comp['SUMMARY']))
                                        continue
                                comp.title = comp['SUMMARY']

                                if comp['CATEGORY'] not in config['categories']:
                                    print(comp)
                                    dtp.messageBox(_("{0:%c}:\nEvent '{1}' in unknown category '{2}'").format(comp.start_dt, comp['SUMMARY'], comp['CATEGORY']))
                                    #break
                                    continue

                                if comp['CATEGORY'] in config and config[comp['CATEGORY']].getboolean('skip'):
                                    dtp.statusMessage(_('{0:%c}:\nSkipping event: {1}').format(comp.start_dt, comp['SUMMARY']))
                                    continue
                            else:
                                print("NO SUMM")
                                dtp.messageBox(_('{0:%c}:\nEvent without summary').format(comp.start_dt))
                                continue

                            comp.subtitle = ""
                            comp.paragraphs = ""
                            if 'DESCRIPTION' in comp:
                                paragraphs = comp['DESCRIPTION'].split('\n')
                                # TODO: filter out:
                                # TODO: filtrer "(sur inscription.*" vers l'agenda!
                                if len(paragraphs):
                                    # First line is taken as subtitle
                                    comp.subtitle = paragraphs[0]
                                    comp.paragraphs = paragraphs[1:]

                            events.append(comp)

                        else:
                            print("Skipping %s object" % comp.name)

                    statusDone += 100/len(urls)
                    dtp.progressSet(int(statusDone))

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

    # sort events in place
    events.sort(key=lambda e: e.start_dt)

    return events

class ActionHandler:
    def __init__(self, action):
        self.action = action
        self.groups = None
        if 'groups' in config[self.action]:
            self.groups = config[self.action]['groups'].split(" ") or []

    def ApplyTransforms(self, group, s):
        k = "%s_transforms" % group
        print("TRANS:%s" % k)
        if k not in config[self.action]:
            return s
        transforms = config[self.action][k].split(" ")

        if "capitalize" in transforms:
            # does not work with ** markdown before name
            s = s.capitalize()
        if "title" in transforms:
            s = s.title()
        if "newline" in transforms:
            s += "  \n"
        if "newparagraph" in transforms:
            s += "\n\n"
        return s

    def HandleGroup(self, events, level):
        md = ""
        print("HandleGroup(%d)" % level)
        if level >= len(self.groups):
            print("level %d > groups!" % level)
            return
        group = self.groups[level]

        if group == "items":
            # sort events in place
            events.sort(key=lambda e: e.start_dt)
            for event in events:
                print()
                print(event)
                print(event.__dict__)
                if not isinstance(event.start, datetime):
                    s = config[self.action]['items_multi_model'].format(event=event)
                else:
                    s = config[self.action]['items_model'].format(event=event)
                md += self.ApplyTransforms(group, s)
            #sys.exit(0)
        elif group == "descriptions":
            # sort events in place
            events.sort(key=lambda e: e.start_dt)
            # find common strings first
            subtitles = [x.subtitle for x in events]
            subtitles_done = False
            print(subtitles)
            if len(subtitles) == subtitles.count(subtitles[0]):
                if len(subtitles[0]):
                    s = config[self.action]['descriptions_subtitle_model'].format(subtitle=subtitles[0])
                    md += self.ApplyTransforms('descriptions_subtitle', s)
                subtitles_done = True
            all_paragraphs = [x.paragraphs for x in events]
            common_paragraphs = []
            for i in range(0,len(all_paragraphs[0])):
                nth_paragraphs = ["" if i >= len(x) else x[i] for x in all_paragraphs]
                print(nth_paragraphs)
                if len(nth_paragraphs[0]) < 1:
                    continue
                # They are all the same
                if len(nth_paragraphs) == nth_paragraphs.count(nth_paragraphs[0]):
                    paragraph = nth_paragraphs[0]
                    common_paragraphs.append(paragraph)
                    s = config[self.action]['descriptions_paragraph_model'].format(paragraph=paragraph)
                    md += self.ApplyTransforms('descriptions_paragraph', s)


            for event in events:
                print()
                print(event)
                print(event.__dict__)
                if not subtitles_done:
                    subtitle = event.subtitle
                    s = config[self.action]['descriptions_subtitle_model'].format(subtitle=subtitle)
                    md += self.ApplyTransforms('descriptions_subtitle', s)
                for paragraph in event.paragraphs:
                    if paragraph in common_paragraphs:
                        # Skip these
                        continue
                    s = config[self.action]['descriptions_paragraph_model'].format(paragraph=paragraph)
                    md += self.ApplyTransforms('descriptions_paragraph', s)

                if not isinstance(event.start, datetime):
                    s = config[self.action]['descriptions_date_multi_model'].format(event=event)
                else:
                    s = config[self.action]['descriptions_date_model'].format(event=event)
                md += self.ApplyTransforms('descriptions_date', s)
            #sys.exit(0)
        elif group == "title":
            # sort events in place
            events.sort(key=lambda e: e.title)
            grouped = itertools.groupby(events, lambda x: x.title)
            for k, g in grouped:
                gevents = list(g)
                title=gevents[0].title
                s = config[self.action]['title_header'].format(title=title)
                md += self.ApplyTransforms(group, s)
                md += self.HandleGroup(gevents, level+1)
            #sys.exit(0)
        elif group == "category":
            categories = config['categories'].keys()
            # eliminate default config keys which appear in all sections
            categories = filter(lambda x: x not in config_defaults.keys(), categories)
            if config['source']['categories_uppercase'] == "true": #XXX
                categories = [str.upper(x) for x in categories]
                print(categories)

            for category in categories:
                if category in config and config[category].getboolean('skip'):
                    continue
                gevents = list(filter(lambda x: x['CATEGORY'] == category, events))
                if len(gevents) or category in config and config[category].getboolean('force'):
                    preamble=''
                    if category in config and 'preamble' in config[category]:
                        preamble = config[category]['preamble']
                    category_pretty = category
                    if category in config['categories']:
                        category_pretty = config['categories'][category]
                    s = config[self.action]['category_header'].format(category=category, category_pretty=category_pretty, preamble=preamble)
                    md += self.ApplyTransforms(group, s)
                    md += self.HandleGroup(gevents, level+1)

            return md
            #     md += "# %s\n" % config['categories'][category]
            #     if category in config and "preamble" in config[category]:
            #         md += "%s\n" % config[category]["preamble"]

            # sort events in place
            events.sort(key=lambda e: e.start_dt)
            grouped = itertools.groupby(events, lambda x: categories.index(x['CATEGORY']))
            for k, g in grouped:
                gevents = list(g)
                category = gevents[0]['CATEGORY']
                category_pretty = category
                if category in config and config[category].getboolean('skip'):
                    continue
                if category in config['categories']:
                    category_pretty = config['categories'][category]
                preamble=''
                if 'preamble' in config[category]:
                    preamble = config[category]["preamble"]

                s = config[self.action]['category_header'].format(category=category, category_pretty=category_pretty, preamble=preamble)
                md += self.ApplyTransforms(group, s)
                md += self.HandleGroup(gevents, level+1)

            #return ""
        elif group == "month":
            # sort events in place
            events.sort(key=lambda e: e.start_dt)
            grouped = itertools.groupby(events, lambda x: x.start_dt.replace(day=1).date())
            for k, g in grouped:
                gevents = list(g)
                start_dt = gevents[0].start_dt.date()
                # can't get config parser to keep the ending newline
                #print("'%s'" % (bytes(config[self.action]['month_header'], "utf-8").decode('string_escape')))
                s = config[self.action]['month_header'].format(start_dt=start_dt)
                md += self.ApplyTransforms(group, s)
                md += self.HandleGroup(gevents, level+1)
        elif group == "day":
            # TODO: config: day_multi_first
            # sort events in place
            print(len(events))
            print(events[0].__dict__)
            events.sort(key=lambda e: e.start_dt)
            print(len(events))
            print(events[0].__dict__)
            # day_multi_first ? -1 if not isinstance(x.start, datetime) else
            grouped = itertools.groupby(events, lambda x: x.start_dt.date())
            for k, g in grouped:
                gevents = list(g)
                start_dt = gevents[0].start_dt.date()
                print(start_dt)
                # can't get config parser to keep the ending newline
                #print("'%s'" % (bytes(config[self.action]['month_header'], "utf-8").decode('string_escape')))
                s = config[self.action]['day_header'].format(start_dt=start_dt)
                md += self.ApplyTransforms(group, s)
                md += self.HandleGroup(gevents, level+1)
        else:
            print("Unknown group type %s" % group)
            return ""

        return md

    def Handle(self, events):
        print("Handle()")

        if 'actions' in config[self.action]:
            actions = config[self.action]['actions'].split(" ")
            for action in actions:
                print("Calling handler for action '%s'" % action)
                handler = ActionHandler(action)
                handler.Handle(events)

        if 'groups' not in config[self.action]:
            return

        # Leaf action
        print("Handle() leaf %s" % self.action)
        frame = config[self.action]['frame']
        print(f"{frame=}")
        md = ""

        statusDone = 0
        statusMax = 100
        dtp.progressReset()
        dtp.progressTotal(statusMax+1)
        dtp.progressSet(0)
        dtp.statusMessage(_('Inserting'))

        dtp.enterUndoContext( _('Insert iCalendar') )

        md = self.HandleGroup(events, 0)

        #dtp.progressSet(min(int(statusDone),99))
        print(md)
        # Apply post filter TODO: configure this in .ini
        md = md.replace("h00","h")
        md = md.replace("*0h\u21920h*","")

        html = markdown.markdown(md)
        html = '<?xml version="1.0" encoding="utf-8"?><html><head></head><body>%s</body></html>\n' % html

        styles = None
        if 'styles' in config[self.action]:
            keys = config[config[self.action]['styles']].keys()
            # eliminate default config keys which appear in all sections
            keys = filter(lambda x: x not in config_defaults.keys(), keys)
            styles = {key:config[config[self.action]['styles']][key] for key in keys}
            print(f"{styles=}")
            #sys.exit(1)


        #dtp.insertHtmlText(html, frame)
        dtp.insertRestyledHtmlText(html, frame, styles)

        #print(scribus.getPosition())
        #scribus.insertText("Foo\nbar\n\ntoto", -1, frame)
        #print(scribus.getPosition())
        dtp.leaveUndoContext()
        dtp.progressEnd()
        return


def InsertICalendar( ):
    events = OpenICalendar()


    handler = ActionHandler('general')
    handler.Handle(events)



def InsertICalendarOld( ):

    frame = agenda_text_block

    events = OpenICalendar()
    
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
    
    dtp.enterUndoContext( _('Insert iCalendar') )
    md = ""

    # TODO: sort into categories
    categories = config['categories'].keys()
    # eliminate default config keys which appear in all sections
    categories = filter(lambda x: x not in config_defaults.keys(), categories)
    if config['source']['categories_uppercase'] == "true": #XXX
        categories = [str.upper(x) for x in categories]
        print(categories)
    for category in categories:
        if category in config and config[category].getboolean('skip'):
            continue
        md += "# %s\n" % config['categories'][category]
        if category in config and "preamble" in config[category]:
            md += "%s\n" % config[category]["preamble"]

    for comp in events:
        start = comp.start_dt
        end = comp.end
        # TODO: check Windows: %- -> %# ?
        # https://strftime.org/
        dfmt = "%A %-d %B"
        tfmt = "%-Hh%M"
        print("XXX:{event.start:%A %-d %B} {event.start:%-Hh%M %-X}//{event.end_dt.minute:d}\u2192{event.end:%-Hh%-0M}".format(event=comp))
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
    # Apply post filter TODO: configure this in .ini
    md = md.replace("h00","h")
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

# TODO: check if we can get the name of the calling function
# maybe sys.exc_info()[2] ?
# cf. help("TRACEBACKS")

# Scribus: handle script launch

def main(argv):
    """Application initialization, font checks and initial setup."""
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
