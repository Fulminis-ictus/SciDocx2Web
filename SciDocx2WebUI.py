"""
Convert scientific papers in DOCX format to HTML. See this project's GitHub page for more info: https://github.com/Fulminis-ictus/SciDocx2Web

This module displays the GUI and calls SciDocx2WebConversion to start the actual conversion process.

Documentation last updated: 2023.05.16\n
Author: Tim Reichert\n
Version: 1.0 (first public release)

Uses and is dependent on Mammoth: https://github.com/mwilliamson/python-mammoth\n
Makes use of dwasyl's added page break detection functionailty: https://github.com/dwasyl/python-mammoth/commit/38777ee623b60e6b8b313e1e63f12dafd82b63a4
"""

### IMPORTS ###
# Extraction and Conversion modules
import mammoth # Convert docx to html
from lxml import etree # XML and XPath
import SciDocx2WebConversion as SciConvert # Handles this tool's conversion

# Path
import os.path

# GUI
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import scrolledtext

# Saving settings to and loading them from .ini
from configparser import ConfigParser

### GUI ###
# input/output variables
inputPath = None
outputPath = None

## FUNCTIONS
# input and output functions
def inputPathFunc():
    '''"Browse" button:

    Prompts the user to choose the file that shall be converted and replaces the input path field text with the path of the chosen file.'''

    global inputPath

    # delete current text in input field
    inputPathEntry.config(state="normal")
    inputPathEntry.delete(0, "end")
    inputPathEntry.config(state="disabled")
    
    # get input path
    inputPath = filedialog.askopenfile(mode="r", filetypes=((".docx file", "*.docx"), ("All files", "*.*")))

    # insert input path into field
    if (inputPath != None) and (inputPath != ""):
        inputPathEntry.config(state="normal")
        inputPath = os.path.abspath(inputPath.name) # get only stem of file
        inputPathEntry.insert(0, inputPath)
        inputPathEntry.xview("end") # scroll view to the right-most part of the path text
        inputPathEntry.config(state="disabled")

    return

def submitFunc():
    '''"Convert" button:

    Prompts the user to choose the output path. Starts the conversion process by calling "convertAndExport()".\n
    Throws an error if no input file has been chosen yet.'''

    global outputPath
    
    # get output path
    outputPath = filedialog.asksaveasfilename(defaultextension=".html", filetypes=((".html file", "*.html"), ("All files", "*.*")))

    # start conversion process
    if inputPath != None:
        if (outputPath != None) and (outputPath != ""):
            convertAndExport()
    else:
        messagebox.showerror("No input file given", "Choose an input file.")

    return

# enable/disable fields depending on what other fields are enabled or disabled and reset variables if necessary
def ablePageTitleAndCssAndJavascript():
    '''Disables the "Page title" entry, "Add suggested css?" checkbox and "Add javascript to highlight navigation while scrolling?" checkbox if "Only export the body?" is checked. Does the opposite if it's unchecked.'''
    # "Only export the body?" is checked: disable
    if bodyCheckVar.get():
        pageTitleEntry.config(state="disabled")
        cssCheck.config(state="disabled")
        cssCheckVar.set(False)
        javascriptCheck.config(state="disabled")
        javascriptCheckVar.set(False)

    # enable
    else:
        pageTitleEntry.config(state="normal")
        cssCheck.config(state="normal")
        javascriptCheck.config(state="normal")

    return

def ableNavigation():
    '''Disables the "Create navigation?" checkbox and the "Paragraph" and "Button" radio buttons if "Add IDs to headings?" is unchecked. Does the opposite if it's checked.'''

    # "Add IDs to headings?" is checked: enable
    if headingsIDVar.get():
        navigationCheck.config(state="normal")

    # disable
    else:
        navigationCheck.config(state="disabled")
        navigationPar.config(state="disabled")
        navigationBut.config(state="disabled")
        navigationVar.set(False)

    return

def ableNavigationElement():
    '''Disables the "Paragraph" and "Button" radio buttons if "Create navigation?" is unchecked. Does the opposite if it's checked.'''

    # "Create navigation?" is checked: enable
    if navigationVar.get():
        navigationPar.config(state="normal")
        navigationBut.config(state="normal")

    # disable
    else:
        navigationPar.config(state="disabled")
        navigationBut.config(state="disabled")

    return

def ableIgnorePNum():
    '''Disables the "Detect paragraphs that should not be numbered..." input field if "Number the paragraphs?" is unchecked. Does the opposite if it's checked.'''

    # "Number the paragraphs?" is checked: enable
    if paragraphNumberCheckVar.get():
        detectIgnorePNumEntry.config(state="normal")

    # disable
    else:
        detectIgnorePNumEntry.config(state="disabled")

    return

def ablePageNum():
    '''Disables the "Which docx page should be counted..." input field if "Insert page numbers?" is unchecked. Does the opposite if it's checked.'''

    # "Insert page numbers?" is checked: enable
    if pageNumberCheckVar.get():
        pageNumberStartCheckEntry.config(state="normal")

    # disable
    else:
        pageNumberStartCheckEntry.config(state="disabled")

    return

def ableAbbreviateTooltips():
    '''Disables the "Abbreviate tooltips after how many symbols?" input field if "Add tooltips to footnotes?" is unchecked. Does the opposite if it's checked.'''

    # "Add tooltips to footnotes?" is checked: enable
    if tooltipsCheckVar.get():
        abbreviateTooltipsEntry.config(state="normal")

    # disable
    else:
        abbreviateTooltipsEntry.config(state="disabled")

    return

def saveOptions():
    '''"Save options" button:
    
    Writes current settings to the INI file and display a message stating that settings has been saved successfully.'''

    # read .ini file
    config.read(iniLocation)

    # set to new values
    config.set("Body and head", "bodyCheckVar", str(bodyCheckVar.get()))
    config.set('Body and head', 'csscheckvar', str(cssCheckVar.get()))
    config.set('Body and head', 'javascriptcheckvar', str(javascriptCheckVar.get()))
    config.set("Body and head", "pagetitleentrytext", pageTitleEntry.get())
    config.set("Heading IDs and nav", "headingsidvar", str(headingsIDVar.get()))
    config.set("Heading IDs and nav", "navigationvar", str(navigationVar.get()))
    config.set("Heading IDs and nav", "navigationtypevar", navigationTypeVar.get())
    config.set('Format templates', 'detectheadingsentry1', detectHeadingsEntry1.get())
    config.set('Format templates', 'detectheadingsentry2', detectHeadingsEntry2.get())
    config.set('Format templates', 'detectheadingsentry3', detectHeadingsEntry3.get())
    config.set('Format templates', 'detectimagesentry', detectImagesEntry.get())
    config.set('Format templates', 'imagesdimensionsentry', imagesDimensionsEntry.get())
    config.set('Format templates', 'detectvideosentry', detectVideosEntry.get())
    config.set('Format templates', 'videosdimensionsentry', videosDimensionsEntry.get())
    config.set('Format templates', 'detectaudioentry', detectAudioEntry.get())
    config.set('Format templates', 'detectMediaentry', detectMediaEntry.get())
    config.set('Format templates', 'detectblockquotesentry', detectBlockquotesEntry.get())
    config.set('Format templates', 'detecttablecaptionsentry', detectTableCaptionsEntry.get())
    config.set('Format templates', 'detectBibliographyentry', detectBibliographyEntry.get())
    config.set('Format templates', 'detectignorepnumentry', detectIgnorePNumEntry.get())
    config.set('Format templates', 'detectcodeentry', detectCodeEntry.get())
    config.set('Format templates', 'customstylemap', customStyleMapEntry.get('1.0', 'end'))
    config.set('Tooltips', 'tooltipscheckvar', str(tooltipsCheckVar.get()))
    config.set('Tooltips', 'abbreviatetooltipsentry', abbreviateTooltipsEntry.get())
    config.set('Citability', 'paragraphnumbercheckvar', str(paragraphNumberCheckVar.get()))
    config.set('Citability', 'pagenumbercheckvar', str(pageNumberCheckVar.get()))
    config.set('Citability', 'pagenumberstartcheckvar', str(pageNumberStartCheckEntry.get()))

    # write .ini file
    with open(iniLocation, "w") as configFile:
        config.write(configFile)

    messagebox.showinfo("Saved", "Settings have been saved.")

    return

def resetOptions():
    '''"Reset options" button:
    
    Asks the user if they really want to reset the current options.

    If yes: Resets options to their default state. The default state has been determined by the author of this software. Updates the GUI, writes new settings to the INI file and displays a message stating that settings have been reset successfully.'''

    reallyReset = messagebox.askquestion('Reset options', 'Are you sure you want to reset the options? Current settings will be overwritten.', icon='warning')
    if reallyReset == 'yes':
        # read .ini file
        config.read(iniLocation)

        # reset to default values
        config.set("Body and head", "bodyCheckVar", "True")
        config.set('Body and head', 'csscheckvar', "False")
        config.set('Body and head', 'javascriptcheckvar', "False")
        config.set("Body and head", "pagetitleentrytext", "")
        config.set("Heading IDs and nav", "headingsidvar", "True")
        config.set("Heading IDs and nav", "navigationvar", "False")
        config.set("Heading IDs and nav", "navigationtypevar", "paragraph")
        config.set('Format templates', 'detectheadingsentry1', "FVMW Heading")
        config.set('Format templates', 'detectheadingsentry2', "FVMW Heading2")
        config.set('Format templates', 'detectheadingsentry3', "FVMW Heading3")
        config.set('Format templates', 'detectimagesentry', "FVMW Image")
        config.set('Format templates', 'imagesdimensionsentry', "")
        config.set('Format templates', 'detectvideosentry', "FVMW Video")
        config.set('Format templates', 'videosdimensionsentry', "")
        config.set('Format templates', 'detectaudioentry', "FVMW Audio")
        config.set('Format templates', 'detectMediaentry', "FVMW Media")
        config.set('Format templates', 'detectblockquotesentry', "FVMW Blockquote")
        config.set('Format templates', 'detecttablecaptionsentry', "FVMW TableCaption")
        config.set('Format templates', 'detectbibliographyentry', "FVMW Bibliography")
        config.set('Format templates', 'detectignorepnumentry', "FVMW IgnorePNum")
        config.set('Format templates', 'detectcodeentry', "FVMW Code")
        config.set('Format templates', 'customstylemap', "")
        config.set('Tooltips', 'tooltipscheckvar', "True")
        config.set('Tooltips', 'abbreviatetooltipsentry', "500")
        config.set('Citability', 'paragraphnumbercheckvar', "True")
        config.set('Citability', 'pagenumbercheckvar', "False")
        config.set('Citability', 'pagenumberstartcheckvar', "1")

        # write to file
        with open(iniLocation, "w") as configFile:
            config.write(configFile)

        # update GUI with reset values
        bodyCheckVar.set(True)
        cssCheckVar.set("False")
        cssCheck.configure(state="disabled")
        javascriptCheckVar.set("False")
        javascriptCheck.configure(state="disabled")
        pageTitleEntryText.set("")
        pageTitleEntry.configure(state="disabled")

        headingsIDVar.set(True)
        navigationVar.set(False)
        navigationCheck.configure(state="normal")
        navigationTypeVar.set("paragraph")
        navigationPar.configure(state="disabled")
        navigationBut.configure(state="disabled")

        detectHeadingsEntry1Text.set("FVMW Heading")
        detectHeadingsEntry2Text.set("FVMW Heading2")
        detectHeadingsEntry3Text.set("FVMW Heading3")
        detectImagesEntryText.set("FVMW Image")
        imagesDimensionsEntryText.set("")
        detectVideosEntryText.set("FVMW Video")
        videosDimensionsEntryText.set("")
        detectAudioEntryText.set("FVMW Audio")
        detectMediaEntryText.set("FVMW Media")
        detectBlockquotesEntryText.set("FVMW Blockquote")
        detectTableCaptionsEntryText.set("FVMW TableCaption")
        detectBibliographyEntryText.set("FVMW Bibliography")
        detectIgnorePNumEntryText.set("FVMW IgnorePNum")
        detectIgnorePNumEntry.configure(state="normal")
        detectCodeEntryText.set("FVMW Code")
        customStyleMapEntry.delete("1.0", "end")

        tooltipsCheckVar.set(True)
        abbreviateTooltipsEntryText.set("500")
        abbreviateTooltipsEntry.configure(state="normal")

        paragraphNumberCheckVar.set(True)
        pageNumberCheckVar.set(False)
        pageNumberStartCheckEntryText.set("1")

        messagebox.showinfo("Reset successful", "Settings have been reset to original values.")

    return


## READ INI ON STARTUP
# open .ini
__location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__))) # get current location
iniLocation = os.path.join(__location__, 'SciDocx2Web.ini')
config = ConfigParser()
config.read(iniLocation)

# read .ini values
conf_bodycheckvar = config.getboolean('Body and head', 'bodyCheckVar')
conf_csscheckvar = config.getboolean('Body and head', 'csscheckvar')
conf_javascriptcheckvar = config.getboolean('Body and head', 'javascriptcheckvar')
conf_pagetitleentrytext = config.get('Body and head', 'pagetitleentrytext')
conf_headingsidvar = config.getboolean('Heading IDs and nav', 'headingsidvar')
conf_navigationvar = config.getboolean('Heading IDs and nav', 'navigationvar')
conf_navigationtypevar = config.get('Heading IDs and nav', 'navigationtypevar')
conf_detectheadingsentry1 = config.get('Format templates', 'detectheadingsentry1')
conf_detectheadingsentry2 = config.get('Format templates', 'detectheadingsentry2')
conf_detectheadingsentry3 = config.get('Format templates', 'detectheadingsentry3')
conf_detectimagesentry = config.get('Format templates', 'detectimagesentry')
conf_imagesdimensionsentry = config.get('Format templates', 'imagesdimensionsentry')
conf_detectvideosentry = config.get('Format templates', 'detectvideosentry')
conf_videosdimensionsentry = config.get('Format templates', 'videosdimensionsentry')
conf_detectaudioentry = config.get('Format templates', 'detectaudioentry')
conf_detectmediaentry = config.get('Format templates', 'detectMediaentry')
conf_detectblockquotesentry = config.get('Format templates', 'detectblockquotesentry')
conf_detecttablecaptionsentry = config.get('Format templates', 'detecttablecaptionsentry')
conf_detectbibliographyentry = config.get('Format templates', 'detecbibliographyentry')
conf_detectignorepnumentry = config.get('Format templates', 'detectignorepnumentry')
conf_detectcodeentry = config.get('Format templates', 'detectcodeentry')
conf_customstylemap = config.get('Format templates', 'customstylemap')
conf_tooltipscheckvar = config.getboolean('Tooltips', 'tooltipscheckvar')
conf_abbreviatetooltipsentry = config.get('Tooltips', 'abbreviatetooltipsentry')
conf_paragraphnumbercheckvar = config.getboolean('Citability', 'paragraphnumbercheckvar')
conf_pagenumbercheckvar = config.getboolean('Citability', 'pagenumbercheckvar')
conf_pagenumberstartcheckvar = config.get('Citability', 'pagenumberstartcheckvar')

## GUI SETUP
# window
window = tk.Tk()
window.title('SciDocx2Web')
window.resizable(False, False)

# add scrollbar
scrollCanvas = tk.Canvas(window, width=600, height=450)
scrollCanvas.grid(row=3, column=0, columnspan=2, sticky="NEWS")

scrollBar = ttk.Scrollbar(window, orient="vertical", command=scrollCanvas.yview)
scrollBar.grid(row=3, column=2, sticky="NS")

scrollCanvas.configure(yscrollcommand=scrollBar.set, scrollregion=scrollCanvas.bbox("all"))
scrollCanvas.bind("<Configure>", lambda e: scrollCanvas.configure(scrollregion=scrollCanvas.bbox("all")))

scrollSecondFrame = ttk.Frame(scrollCanvas, width=200, height=600)
scrollSecondFrame.grid(row=0, column=0, sticky="NW")

scrollCanvas.create_window((0,0), window=scrollSecondFrame, anchor="nw")

def _on_mousewheel(event):
    '''Makes the mouse wheel scroll the canvas.'''

    scrollCanvas.yview_scroll(int(-1*(event.delta/120)), "units")

    return

scrollCanvas.bind_all("<MouseWheel>", _on_mousewheel)

# variable used for counting up rows (makes it easier to rearrange the UI without having to manually update all values)
row = 1

# --Body and head settings--
# frame
frameBody = tk.LabelFrame(scrollSecondFrame, text='Body and head settings')
frameBody.grid(sticky="W", row=row, column=0, pady=(20, 10), padx=(20,0))

# "Only export the body?"
bodyCheckVar = tk.BooleanVar(value=conf_bodycheckvar)
exportBodyCheck = tk.Checkbutton(frameBody, text='Only export the body?',variable=bodyCheckVar, onvalue=True, offvalue=False, command=ablePageTitleAndCssAndJavascript)
exportBodyCheck.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

# "Add suggested css?"
row += 1

cssCheckVar = tk.BooleanVar(value=conf_csscheckvar)
cssCheck = tk.Checkbutton(frameBody, text='Add suggested css?',variable=cssCheckVar, onvalue=True, offvalue=False, justify="left")
if conf_bodycheckvar:
    cssCheck.configure(state="disable")
else:
    cssCheck.configure(state="normal")
cssCheck.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

# "Add javascript to highlight navigation while scrolling?"
row += 1

javascriptCheckVar = tk.BooleanVar(value=conf_javascriptcheckvar)
javascriptCheck = tk.Checkbutton(frameBody, text='Add javascript to highlight navigation while scrolling?',variable=javascriptCheckVar, onvalue=True, offvalue=False, justify="left")
if conf_bodycheckvar:
    javascriptCheck.configure(state="disable")
else:
    javascriptCheck.configure(state="normal")
javascriptCheck.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

# "Page title"
row += 1

pageTitleLabel = ttk.Label(frameBody, text='Page title:')
pageTitleLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

pageTitleEntryText = tk.StringVar(value=conf_pagetitleentrytext)
pageTitleEntry = tk.Entry(frameBody, textvariable=pageTitleEntryText)
if conf_bodycheckvar:
    pageTitleEntry.configure(state="disabled")
else:
    pageTitleEntry.configure(state="normal")
pageTitleEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# --Navigation--
row += 1

# frame
frameHeadingsNav = tk.LabelFrame(scrollSecondFrame, text='Navigation')
frameHeadingsNav.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

# "Add IDs to headings?"
headingsIDVar = tk.BooleanVar(value=conf_headingsidvar)
headingsIDCheck = tk.Checkbutton(frameHeadingsNav, text='Add IDs to headings?',variable=headingsIDVar, onvalue=True, offvalue=False, command=ableNavigation)
headingsIDCheck.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

# "Create navigation?"
navigationVar = tk.BooleanVar(value=conf_navigationvar)
navigationCheck = tk.Checkbutton(frameHeadingsNav, text='Create navigation?',variable=navigationVar, onvalue=True, offvalue=False, command=ableNavigationElement)
if conf_headingsidvar:
    navigationCheck.configure(state="normal")
else:
    navigationCheck.configure(state="disabled")
navigationCheck.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,0))

# "Paragraph" + "Button" radio buttons
navigationTypeVar = tk.StringVar(value=conf_navigationtypevar)

navigationPar = tk.Radiobutton(frameHeadingsNav, text='Paragraph',variable=navigationTypeVar, value="paragraph")
if conf_navigationvar:
    navigationPar.configure(state="normal")
else:
    navigationPar.configure(state="disabled")
navigationPar.grid(sticky="W", row=row, column=2, pady=(0, 0), padx=(20,20))

row += 1

navigationBut = tk.Radiobutton(frameHeadingsNav, text='Button',variable=navigationTypeVar, value="button")
if conf_navigationvar:
    navigationBut.configure(state="normal")
else:
    navigationBut.configure(state="disabled")
navigationBut.grid(sticky="W", row=row, column=2, pady=(0, 10), padx=(20,20))

# --Tooltips settings--
# frame
frameTooltips = tk.LabelFrame(scrollSecondFrame, text='Tooltip settings')
frameTooltips.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

# "Add tooltips to footnote?"
row += 1

tooltipsCheckVar = tk.BooleanVar(value=conf_tooltipscheckvar)
tooltipsCheck = tk.Checkbutton(frameTooltips, text='Add tooltips to footnotes?',variable=tooltipsCheckVar, onvalue=True, offvalue=False, command=ableAbbreviateTooltips)
tooltipsCheck.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

# "Abbreviate tooltips..."
row += 1

abbreviateTooltipsLabel = tk.Label(frameTooltips, text='Abbreviate tooltips after how many symbols? Input a number.\nLeave empty to skip abbreviation.', justify="left")
abbreviateTooltipsLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

abbreviateTooltipsEntryText = tk.StringVar(value=conf_abbreviatetooltipsentry)
abbreviateTooltipsEntry = tk.Entry(frameTooltips, textvariable=abbreviateTooltipsEntryText)
if conf_tooltipscheckvar:
    abbreviateTooltipsEntry.configure(state="normal")
else:
    abbreviateTooltipsEntry.configure(state="disable")
abbreviateTooltipsEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# --Citability settings--
# frame
framePar = tk.LabelFrame(scrollSecondFrame, text='Citability settings')
framePar.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

# "Number the paragraphs?"
row += 1

paragraphNumberCheckVar = tk.BooleanVar(value=conf_paragraphnumbercheckvar)
paragraphNumberCheck = tk.Checkbutton(framePar, text='Number the paragraphs?',variable=paragraphNumberCheckVar, onvalue=True, offvalue=False, command=ableIgnorePNum)
paragraphNumberCheck.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,20))

row += 1

pageNumberCheckVar = tk.BooleanVar(value=conf_pagenumbercheckvar)
pageNumberCheck = tk.Checkbutton(framePar, text='Insert page numbers?',variable=pageNumberCheckVar, onvalue=True, offvalue=False, command=ablePageNum)
pageNumberCheck.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,20))

row += 1

pageNumberStartCheckLabel = tk.Label(framePar, text='Which docx page should be counted as the first page?\nInput a number.', justify="left")
pageNumberStartCheckLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

pageNumberStartCheckEntryText = tk.StringVar(value=conf_pagenumberstartcheckvar)
pageNumberStartCheckEntry = tk.Entry(framePar, textvariable=pageNumberStartCheckEntryText)
if conf_pagenumbercheckvar:
    pageNumberStartCheckEntry.configure(state="normal")
else:
    pageNumberStartCheckEntry.configure(state="disable")
pageNumberStartCheckEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(25,20))


# --Format template detection--
# frame
frameDetection = tk.LabelFrame(scrollSecondFrame, text='Format template detection')
frameDetection.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

# "Detect headings..."
row += 1

detectHeadingsLabel = tk.Label(frameDetection, text='Detect 1. level headings (h1) by which format template name?\nLeave empty to skip detection.', justify="left")
detectHeadingsLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectHeadingsEntry1Text = tk.StringVar(value=conf_detectheadingsentry1)
detectHeadingsEntry1 = tk.Entry(frameDetection, textvariable=detectHeadingsEntry1Text)
detectHeadingsEntry1.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

row += 1

detectHeadingsLabel = tk.Label(frameDetection, text='Detect 2. level headings (h2) by which format template name?\nLeave empty to skip detection.', justify="left")
detectHeadingsLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectHeadingsEntry2Text = tk.StringVar(value=conf_detectheadingsentry2)
detectHeadingsEntry2 = tk.Entry(frameDetection, textvariable=detectHeadingsEntry2Text)
detectHeadingsEntry2.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

row += 1

detectHeadingsLabel = tk.Label(frameDetection, text='Detect 3. level headings (h3) by which format template name?\nLeave empty to skip detection.', justify="left")
detectHeadingsLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectHeadingsEntry3Text = tk.StringVar(value=conf_detectheadingsentry3)
detectHeadingsEntry3 = tk.Entry(frameDetection, textvariable=detectHeadingsEntry3Text)
detectHeadingsEntry3.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# "Detect image references..."
row += 1

detectImagesLabel = tk.Label(frameDetection, text='Detect image references by which format template name?\nLeave empty to skip detection.', justify="left")
detectImagesLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectImagesEntryText = tk.StringVar(value=conf_detectimagesentry)
detectImagesEntry = tk.Entry(frameDetection, textvariable=detectImagesEntryText)
detectImagesEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

row += 1

imagesDimensionsLabel = tk.Label(frameDetection, text='Which dimensions should the image embed have?\nSeparate X and Y value with a comma (X,Y).\nLeave empty to use original dimensions.', justify="left")
imagesDimensionsLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

imagesDimensionsEntryText = tk.StringVar(value=conf_imagesdimensionsentry)
imagesDimensionsEntry = tk.Entry(frameDetection, textvariable=imagesDimensionsEntryText)
imagesDimensionsEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# "Detect video references..."
row += 1

detectVideosLabel = tk.Label(frameDetection, text='Detect video references by which format template name?\nLeave empty to skip detection.', justify="left")
detectVideosLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectVideosEntryText = tk.StringVar(value=conf_detectvideosentry)
detectVideosEntry = tk.Entry(frameDetection, textvariable=detectVideosEntryText)
detectVideosEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

row += 1

videosDimensionsLabel = tk.Label(frameDetection, text='Which dimensions should the video embed have?\nSeparate X and Y value with a comma (X,Y).\nLeave empty to use original dimensions.', justify="left")
videosDimensionsLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

videosDimensionsEntryText = tk.StringVar(value=conf_videosdimensionsentry)
videosDimensionsEntry = tk.Entry(frameDetection, textvariable=videosDimensionsEntryText)
videosDimensionsEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# "Detect audio references..."
row += 1

detectAudioLabel = tk.Label(frameDetection, text='Detect audio references by which format template name?\nLeave empty to skip detection.', justify="left")
detectAudioLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectAudioEntryText = tk.StringVar(value=conf_detectaudioentry)
detectAudioEntry = tk.Entry(frameDetection, textvariable=detectAudioEntryText)
detectAudioEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# "Detect media captions..."
row += 1

detectMediaLabel = tk.Label(frameDetection, text='Detect media captions by which format template name?\nLeave empty to skip detection.', justify="left")
detectMediaLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectMediaEntryText = tk.StringVar(value=conf_detectmediaentry)
detectMediaEntry = tk.Entry(frameDetection, textvariable=detectMediaEntryText)
detectMediaEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# "Detect table captions..."
row += 1

detectTableCaptionsLabel = tk.Label(frameDetection, text='Detect table captions by which format template name?\nLeave empty to skip detection.', justify="left")
detectTableCaptionsLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectTableCaptionsEntryText = tk.StringVar(value=conf_detecttablecaptionsentry)
detectTableCaptionsEntry = tk.Entry(frameDetection, textvariable=detectTableCaptionsEntryText)
detectTableCaptionsEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# "Detect blockquotes..."
row += 1

detectBlockquotesLabel = tk.Label(frameDetection, text='Detect blockquotes by which format template name?\nLeave empty to skip detection.', justify="left")
detectBlockquotesLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectBlockquotesEntryText = tk.StringVar(value=conf_detectblockquotesentry)
detectBlockquotesEntry = tk.Entry(frameDetection, textvariable=detectBlockquotesEntryText)
detectBlockquotesEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# "Detect bibliography..."
row += 1

detectBibliographyLabel = tk.Label(frameDetection, text='Detect bibliography by which format template name?\nLeave empty to skip detection.', justify="left")
detectBibliographyLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectBibliographyEntryText = tk.StringVar(value=conf_detectbibliographyentry)
detectBibliographyEntry = tk.Entry(frameDetection, textvariable=detectBibliographyEntryText)
detectBibliographyEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# "Detect paragraphs that should not be numbered..."
row += 1

detectIgnorePNumLabel =tk.Label(frameDetection, text='Detect paragraphs that should not be numbered by which\nformat template name?\nLeave empty to skip detection.', justify="left")
detectIgnorePNumLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectIgnorePNumEntryText = tk.StringVar(value=conf_detectignorepnumentry)
detectIgnorePNumEntry = tk.Entry(frameDetection, textvariable=detectIgnorePNumEntryText)
if conf_paragraphnumbercheckvar:
    detectIgnorePNumEntry.configure(state="normal")
else:
    detectIgnorePNumEntry.configure(state="disable")
detectIgnorePNumEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# "Detect code..."
row += 1

detectCodeLabel =tk.Label(frameDetection, text='Detect code paragraphs by which format template name?\nLeave empty to skip detection.', justify="left")
detectCodeLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

detectCodeEntryText = tk.StringVar(value=conf_detectcodeentry)
detectCodeEntry = tk.Entry(frameDetection, textvariable=detectCodeEntryText)
detectCodeEntry.grid(sticky="W", row=row, column=1, pady=(10, 10), padx=(20,20))

# "Custom style map"
row += 1

customStyleMapLabel =tk.Label(frameDetection, text='Additional custom style map entries.', justify="left")
customStyleMapLabel.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

row += 1

customStyleMapEntryText = tk.StringVar(value=conf_customstylemap)
customStyleMapEntry = scrolledtext.ScrolledText(frameDetection, width=60, height=20)
customStyleMapEntry.insert('1.0', customStyleMapEntryText.get())
customStyleMapEntry.grid(sticky="W", row=row, column=0, columnspan=2, pady=(10, 10), padx=(20,20))


# SEPARATOR
separator = ttk.Separator(window, orient='horizontal')
separator.grid(sticky="EW", row=row, columnspan=2)

# "Save options"
row += 1

resetButton = ttk.Button(window, text='Save options', command=saveOptions)
resetButton.grid(sticky="W", row=row, column=0, pady=(20, 10), padx=(20,0))

# "Reset options"
resetButton = ttk.Button(window, text='Reset options', command=resetOptions)
resetButton.grid(sticky="W", row=row, column=1, pady=(20, 10), padx=(20,0))

# "Browse input file"
row += 1

inputPathButton = ttk.Button(window, text='Browse input file', command=inputPathFunc)
inputPathButton.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))

inputPathEntry = tk.Entry(window)
inputPathEntry.grid(sticky="NSEW", row=row, column=1, pady=(10, 10), padx=(20,0))
inputPathEntry.config(state="disabled", justify="center")
window.grid_columnconfigure(1, weight=1)

# "Convert"
row += 1

submitButton = ttk.Button(window, text='Convert', command=submitFunc)
submitButton.grid(sticky="W", row=row, column=0, pady=(10, 10), padx=(20,0))


### CSS ###
# body
bodycss = 'body {margin-left: 15%; margin-right: 15%;}'

# tooltips
tooltipcss = '''\n/* Tooltip container */
.tooltippop {
position: relative;
}

/* Tooltip text */
.tooltippop [role="tooltip"] {
font-size: 12pt;
visibility: hidden;
width: max-content;
max-width: 400px;
background-color: #fff;
color: #454545;
text-align: left;
padding: 5px 5px;
border-radius: 5px;
border: 2px solid black;

/* Position the tooltip text */
position: absolute;
z-index: 1;
bottom: 125%;
left: 50%;
margin-left: -60px;

/* Fade in tooltip */
opacity: 0;
transition: opacity 0.3s;
}

/* Show the tooltip text when you mouse over the tooltip container */
.tooltippop:hover [role="tooltip"] {
visibility: visible;
opacity: 1;
}'''

# grids
gridcss = '''\n/* Grid */
.gridContainer {
    display: grid;
    gap: 50px 50px;
    grid-template-columns: 12% 68%;
}

.navGrid {
    grid-column-start: 1; 
    grid-column-end: 2; 
    grid-row-start: 1; 
    grid-row-end: 2;
    position: sticky;
    top: 0;
    align-self: start;
    padding-right: 5%;
    height: 100vh;
    overflow: auto;
}

.mainGrid {
    grid-column-start: 2; 
    grid-column-end: 3; 
    grid-row-start: 1; 
    grid-row-end: 2;
}'''

# paragraphs
paragraphcss = '\np {font-size: 18px; color: #454545; text-align: left; line-height: 2;}'

# pagenumbers
pagenumbercss = '.pagenumber {font-size: 14px; color: #454545; text-align: left; font-weight:400; background-color: #E7E7E7;}'

# buttons
buttoncss = 'button {font-size: 14px; text-align: left; width: 100%;}\nbutton a {display: block;}'

# highlighted nav
highlightnavcss = 'nav p a.highlightnav {background-color: #D7D7D7;}\nnav button a.active {background-color: #D7D7D7;}'

# headings
headingcss = 'h1, h2, h3 {font-size: 28px; color: #454545; text-align: left; padding-top: 18px; padding-bottom: 6px;}'

# links
linkscss = 'a:link {color: #0000ff; text-decoration:none;}\na:visited {color: #800080; text-decoration:none;}'

# lists
listscss = 'li {font-size: 18px; color: #454545;}'

# tables
tablecss = 'table {margin-top: 28px;}\ntable, th, td {border: 1px solid;}\ntd {padding: 0px 5px;}\ncaption {caption-side: bottom; text-align: left; font-size: 15px;}'

# horizontal rules
hrcss = 'hr {margin-top: 28px;}'

# blockquotes
blockquotecss = 'blockquote {display: block; font-size: 18px; color: #454545; text-align: left; padding-left: 5%; padding-right: 15%; padding-top: 18px; padding-bottom: 18px;}'

# mediacaptions
mediacaptioncss = '.mediacaption {display: block; font-size: 15px; color: #454545; text-align: left;}'

#bibliography
bibliographycss = '.bibliography {display: block; font-size: 18px; color: #454545; text-align: left; text-indent: -5%; margin-left: 5%;}'

# assembly of css
css = '<style>' + '\n' + bodycss + '\n' + tooltipcss + '\n' + gridcss + '\n' + paragraphcss + '\n\n' + pagenumbercss + '\n\n' + buttoncss + '\n\n' + highlightnavcss + '\n\n' + headingcss + '\n\n' + linkscss + '\n\n' + listscss + '\n\n' + tablecss + '\n\n' + hrcss + '\n\n' + blockquotecss + '\n\n' + mediacaptioncss + '\n\n' + bibliographycss + '\n' '\t\t</style>' + '\n\n'

# convert style code to XML element
cssXML = etree.fromstring(css)


### JAVASCRIPT ###
javascript = """<script>
function start() {
    let h1Marker = document.querySelectorAll('h1[id*="heading"], h2[id*="heading"], h3[id*="heading"]');
    let navPA = document.querySelectorAll('nav p a');
    let navButA = document.querySelectorAll('nav button a');

    window.addEventListener('scroll', () => {
        let current = '';

        h1Marker.forEach((section) => {
            const sectionTop = section.offsetTop;
            const sectionHeight = section.clientHeight;

            if (pageYOffset >= sectionTop - sectionHeight / 3) {
                current = section.getAttribute('id');
            }
        });
        navPA.forEach((a) => {
            a.classList.remove("highlightnav");
            if (a.getAttribute("href") == '#' + current) {
                a.classList.add("highlightnav");
            }
        });
        navButA.forEach((a) => {
            a.classList.remove("highlightnav");
            if (a.getAttribute("href") == '#' + current) {
                a.classList.add("highlightnav");
            }
        });
    });
};

window.addEventListener('load', () => {
    start();
});
</script>
"""
javascriptXML = etree.fromstring(javascript)


### CONVERT ###
def convertAndExport():
    '''Converts a DOCX file to an HTML file and exports it by calling functions from SciDocx2WebConversion.py. Displays a "Success" message if conversion was successful.'''

    # style map
    custom_style_map = SciConvert.style_map_func("", detectHeadingsEntry1.get(), detectHeadingsEntry2.get(), detectHeadingsEntry3.get(), detectImagesEntry.get(), detectVideosEntry.get(), detectAudioEntry.get(), detectMediaEntry.get(), detectBlockquotesEntry.get(), detectTableCaptionsEntry.get(), detectBibliographyEntry.get(), detectIgnorePNumEntry.get(), paragraphNumberCheckVar.get(), detectCodeEntry.get(), customStyleMapEntry.get('1.0', 'end'))

    # import and enclose input file with tags
    input = mammoth.convert_to_html(inputPath, style_map=custom_style_map).value
    bodyxml = SciConvert.enclose_body(input, bodyCheckVar.get(), pageTitleEntry.get())

    # create footnotes
    footnotes = SciConvert.create_footnotes_list(bodyxml, abbreviateTooltipsEntry.get())

    # abbreviate footnotes
    footnotesAbbr = SciConvert.abbreviate_footnotes(footnotes, abbreviateTooltipsEntry.get())

    # add wbr to footnotes
    footnotesAbbr = SciConvert.add_wbr_footnotes(footnotesAbbr, abbreviateTooltipsEntry.get())

    # insert footnotes into main text
    bodyxml = SciConvert.insert_footnotes(tooltipsCheckVar.get(), bodyxml, footnotesAbbr)

    # adjust footnote sups
    bodyxml = SciConvert.adjust_footnotes(tooltipsCheckVar.get(), bodyxml)

    # separate bottom footnotes
    commentBottomFootnotes = etree.Comment(' Bottom footnotes ')
    breakElement = etree.XML('<br/>')
    hrElement = etree.XML('<hr/>')
    bodyxml = SciConvert.footnotes_bottom_adjust(bodyxml, commentBottomFootnotes, breakElement, hrElement)

    # add wbr to main text
    bodyxml = SciConvert.add_wbr_text(bodyxml)

    # add heading IDs
    bodyxml = SciConvert.add_Head_IDs(headingsIDVar.get(), bodyxml)

    # remove TOCs and Heads from heading IDs
    bodyxml = SciConvert.remove_toc_and_head(bodyxml)

    # create navigation
    findH1 = bodyxml.xpath('.//*[self::h1 or self::h2 or self::h3]')
    navigationElement = etree.Element('nav')
    commentNavigation = etree.Comment(' Navigation ')
    h1Navigation = etree.Element('h1')
    h1Navigation.text = 'Navigation'
    navGridDiv = etree.Element('div')
    navGridDiv.attrib['class'] = 'navGrid'
    navGridDiv = SciConvert.create_navigation(navigationVar.get(), navigationTypeVar.get(), findH1, navigationElement, commentNavigation, h1Navigation, navGridDiv)

    # add cite to blockquotes
    tooltiptextPath = './/a[contains(@id, "footnote-ref")]/sup'
    bodyxml = SciConvert.add_cite(tooltiptextPath, bodyxml, footnotes)

    # embed images
    bodyxml = SciConvert.embed_images(bodyxml, imagesDimensionsEntry.get())

    # embed videos
    bodyxml = SciConvert.embed_videos(bodyxml, videosDimensionsEntry.get())

    # embed audio
    bodyxml = SciConvert.embed_audio(bodyxml)

    # add file insertion messages above mediacaptions
    bodyxml = SciConvert.file_insertion_message(bodyxml)

    # move table captions
    bodyxml = SciConvert.move_table_caption(bodyxml)

    # create page breaks
    bodyxml = SciConvert.page_breaks(pageNumberCheckVar.get(), pageNumberStartCheckEntry.get(), bodyxml, bodyCheckVar.get())

    # number paragraphs
    bodyxml = SciConvert.paragraph_numbering(paragraphNumberCheckVar.get(), bodyxml)

    # create sections (haven't been able to figure this out yet)
    #bodyxml = SciConvert.create_sections(bodyxml)

    # assemble file
    exportableBodyxml = SciConvert.assemble_html(navigationVar.get(), bodyCheckVar.get(), cssCheckVar.get(), cssXML, navGridDiv, bodyxml, javascriptXML, javascriptCheckVar.get())

    # unescape and escape HTML characters
    exportableBodyxml = SciConvert.escape_unescape(exportableBodyxml)

    # write file
    SciConvert.write_html(exportableBodyxml, outputPath)

    messagebox.showinfo("Success", "The file has been converted successfully.")

    return

### CREATE GUI WINDOW ###
window.mainloop()