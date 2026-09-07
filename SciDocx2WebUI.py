"""
Convert scientific papers in DOCX format to HTML. See this project's GitHub page for more info: https://github.com/Fulminis-ictus/SciDocx2Web

This module displays the GUI and calls SciDocx2WebConversion to start the actual conversion process.

Documentation last updated: 2026.09.06\n
Author: Tim Reichert\n
Version: 1.2

Uses and is dependent on Mammoth: https://github.com/mwilliamson/python-mammoth\n
Makes use of dwasyl's added page break detection functionailty: https://github.com/dwasyl/python-mammoth/commit/38777ee623b60e6b8b313e1e63f12dafd82b63a4

This version of the program is built using Python 3.13.15, Mammoth 1.12.1 and lxml 6.1.3. Using other versions can result in errors, such as missing text.
"""

### IMPORTS ###
# Extraction and Conversion modules
import mammoth # Convert docx to html
from lxml import etree # XML and XPath
import SciDocx2WebConversion as SciConvert # Handles this tool's conversion

# Path
import os.path

# GUI
from PySide6.QtWidgets import QCheckBox, QLabel, QLineEdit, QRadioButton, QTextEdit, QMainWindow, QApplication, QPushButton, QWidget, QVBoxLayout, QSpacerItem, QSizePolicy, QScrollArea, QMessageBox, QFileDialog, QDockWidget, QListWidget
from PySide6.QtCore import Qt

# Saving settings to and loading them from .ini
from configparser import ConfigParser

# .ini location
__location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__))) # get current location
iniLocation = os.path.join(__location__, 'SciDocx2Web.ini')
config = ConfigParser()

class MainWindow(QMainWindow):
    ### GUI SETUP ###
    def __init__(self):

        super().__init__()

        # Locations
        self.inputPath = None
        self.outputPath = None

        # Colors for disabled and enabled UI elements
        self.disabled_color = "gray"
        self.enabled_color = "lightgray"

        self.setWindowTitle('SciDocx2Web')

        # GUI Elements
        # --Body and head settings--
        label_body = QLabel('<font size=5>Body and head settings</font>')
        self.check_body_only = QCheckBox('Only export the body?', self)
        self.check_body_only.toggled.connect(lambda: self.presetChangedFunc("Export Body"))

        self.check_css = QCheckBox('Add suggested css?', self)

        self.check_javascript = QCheckBox('Add javascript to highlight navigation while scrolling?', self)

        self.label_page_title = QLabel('Page title')
        self.page_title = QLineEdit()

        # --Navigation--
        label_navigation = QLabel('<font size=5>Navigation</font>')
        self.add_IDs = QCheckBox('Add IDs to headings?', self)
        self.add_IDs.toggled.connect(lambda: self.presetChangedFunc("Add IDs"))

        self.create_nav = QCheckBox('Create navigation?', self)
        self.create_nav.toggled.connect(lambda: self.presetChangedFunc("Create Navigation"))

        self.navigationPar = QRadioButton('Paragraph', self.create_nav)
        self.navigationBut = QRadioButton('Button', self.create_nav)

        # --Tooltips settings--
        label_tooltip = QLabel('<font size=5>Tooltip settings</font>')
        self.add_tooltips = QCheckBox('Add tooltips to footnotes?', self)
        self.add_tooltips.toggled.connect(lambda: self.presetChangedFunc("Add Tooltips"))

        self.label_tooltip_abbreviate = QLabel('Abbreviate tooltips after how many symbols? Input a number.\nLeave empty to skip abbreviation.')
        self.tooltip_abbreviate = QLineEdit()

        # --Citability settings--
        label_citability = QLabel('<font size=5>Citability settings</font>')
        self.number_paragraphs = QCheckBox('Number the paragraphs?', self)
        self.number_paragraphs.toggled.connect(lambda: self.presetChangedFunc("Number Paragraphs"))

        self.insert_page_no = QCheckBox('Insert page numbers?', self)
        self.insert_page_no.toggled.connect(lambda: self.presetChangedFunc("Insert Page"))

        self.label_first_page = QLabel('Which docx page should be counted as the first page?\nInput a number.')
        self.first_page = QLineEdit()

        # --Format template detection--
        label_detect_templates = QLabel('<font size=5>Format template detection</font>')
        label_detect_headings1 = QLabel('Detect 1. level headings (h1) by which format template name?\nLeave empty to skip detection.')
        self.detect_headings1 = QLineEdit()

        label_detect_headings2 = QLabel('Detect 2. level headings (h2) by which format template name?\nLeave empty to skip detection.')
        self.detect_headings2 = QLineEdit()

        label_detect_headings3 = QLabel('Detect 3. level headings (h3) by which format template name?\nLeave empty to skip detection.')
        self.detect_headings3 = QLineEdit()

        label_detect_images = QLabel('Detect image references by which format template name?\nLeave empty to skip detection.')
        self.detect_images = QLineEdit()

        label_images_dimensions = QLabel('Which dimensions should the image embed have?\nSeparate X and Y value with a comma (X,Y).\nLeave empty to use original dimensions.')
        self.images_dimensions = QLineEdit()

        label_detect_videos = QLabel('Detect video references by which format template name?\nLeave empty to skip detection.')
        self.detect_videos = QLineEdit()

        label_video_dimensions = QLabel('Which dimensions should the video embed have?\nSeparate X and Y value with a comma (X,Y).\nLeave empty to use original dimensions.')
        self.video_dimensions = QLineEdit()

        label_detect_audio = QLabel('Detect audio references by which format template name?\nLeave empty to skip detection.')
        self.detect_audio = QLineEdit()

        label_detect_media = QLabel('Detect media captions by which format template name?\nLeave empty to skip detection.')
        self.detect_media = QLineEdit()

        label_detect_tables = QLabel('Detect table captions by which format template name?\nLeave empty to skip detection.')
        self.detect_tables = QLineEdit()

        label_detect_blockquotes = QLabel('Detect blockquotes by which format template name?\nLeave empty to skip detection.')
        self.detect_blockquotes = QLineEdit()

        label_detect_bibliography = QLabel('Detect bibliography captions by which format template name?\nLeave empty to skip detection.')
        self.detect_bibliography = QLineEdit()

        self.label_detect_ignore_pnum = QLabel('Detect paragraphs that should not be numbered by which format template name?\nLeave empty to skip detection.')
        self.detect_ignore_pnum = QLineEdit()

        label_detect_code = QLabel('Detect code paragraphs by which format template name?\nLeave empty to skip detection.')
        self.detect_code = QLineEdit()

        label_additional_styles = QLabel('Additional custom style map entries.')
        self.additional_styles = QTextEdit()

        # --File--
        self.save_options = QPushButton('Save options')
        self.save_options.clicked.connect(self.saveOptions)

        self.reset_options = QPushButton('Reset options')
        self.reset_options.clicked.connect(self.resetOptions)

        self.browse = QPushButton('Browse input file')
        self.browse.clicked.connect(self.inputPathFunc)

        self.display_file_path = QLineEdit()
        self.display_file_path.setAlignment(Qt.AlignHCenter)
        self.display_file_path.setReadOnly(True)

        self.convert = QPushButton('Convert')
        self.convert.clicked.connect(self.submitFunc)

        # Layout
        self.container = QWidget()
        self.scroll = QScrollArea()
        layout = QVBoxLayout()

        self.container.setLayout(layout)
        
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(self.container)
        
        self.setGeometry(600, 100, 600, 900)

        spacer = QSpacerItem(20, 20, QSizePolicy.Fixed, QSizePolicy.Fixed)
        
        self.setCentralWidget(self.scroll)

        layout.addWidget(label_body)
        layout.addWidget(self.check_body_only)
        layout.addWidget(self.check_css)
        layout.addWidget(self.check_javascript)
        layout.addWidget(self.label_page_title)
        layout.addWidget(self.page_title)
        layout.addItem(spacer)

        layout.addWidget(label_navigation)
        layout.addWidget(self.add_IDs)
        layout.addWidget(self.create_nav)
        layout.addWidget(self.navigationPar)
        layout.addWidget(self.navigationBut)
        layout.addItem(spacer)

        layout.addWidget(label_tooltip)
        layout.addWidget(self.add_tooltips)
        layout.addWidget(self.label_tooltip_abbreviate)
        layout.addWidget(self.tooltip_abbreviate)
        layout.addItem(spacer)
        
        layout.addWidget(label_citability)
        layout.addWidget(self.number_paragraphs)
        layout.addWidget(self.insert_page_no)
        layout.addWidget(self.label_first_page)
        layout.addWidget(self.first_page)
        layout.addItem(spacer)

        layout.addWidget(label_detect_templates)
        layout.addWidget(label_detect_headings1)
        layout.addWidget(self.detect_headings1)
        layout.addWidget(label_detect_headings2)
        layout.addWidget(self.detect_headings2)
        layout.addWidget(label_detect_headings3)
        layout.addWidget(self.detect_headings3)
        layout.addWidget(label_detect_images)
        layout.addWidget(self.detect_images)
        layout.addWidget(label_images_dimensions)
        layout.addWidget(self.images_dimensions)
        layout.addWidget(label_detect_videos)
        layout.addWidget(self.detect_videos)
        layout.addWidget(label_video_dimensions)
        layout.addWidget(self.video_dimensions)
        layout.addWidget(label_detect_audio)
        layout.addWidget(self.detect_audio)
        layout.addWidget(label_detect_media)
        layout.addWidget(self.detect_media)
        layout.addWidget(label_detect_tables)
        layout.addWidget(self.detect_tables)
        layout.addWidget(label_detect_blockquotes)
        layout.addWidget(self.detect_blockquotes)
        layout.addWidget(label_detect_bibliography)
        layout.addWidget(self.detect_bibliography)
        layout.addWidget(self.label_detect_ignore_pnum)
        layout.addWidget(self.detect_ignore_pnum)
        layout.addWidget(label_detect_code)
        layout.addWidget(self.detect_code)
        layout.addWidget(label_additional_styles)
        layout.addWidget(self.additional_styles)
        layout.addItem(spacer)
        layout.addWidget(self.save_options)
        layout.addWidget(self.reset_options)

        # Docked file options at the bottom
        self.container2 = QWidget()
        layout2 = QVBoxLayout()
        self.container2.setLayout(layout2)
        
        layout2.addWidget(self.browse)
        layout2.addWidget(self.display_file_path)
        layout2.addWidget(self.convert)

        dock = QDockWidget(self)
        dock.setAllowedAreas(Qt.DockWidgetArea.BottomDockWidgetArea)
        dock.setWidget(self.container2)

        self.addDockWidget(Qt.DockWidgetArea.BottomDockWidgetArea, dock)

        # Load values from .ini
        self.readIni()

        return
        

    ### FUNCTIONS ###
    def readIni(self):
        '''Load settings from .ini file at startup.'''
        # open .ini
        config.read(iniLocation)

        # read .ini values
        conf_bodycheckvar = config.getboolean('Body and head', 'bodyCheckVar')
        conf_csscheckvar = config.getboolean('Body and head', 'csscheckvar')
        conf_javascriptcheckvar = config.getboolean('Body and head', 'javascriptcheckvar')
        conf_pagetitleentrytext = config.get('Body and head', 'pagetitleentrytext')
        conf_headingsidvar = config.getboolean('Heading IDs and nav', 'headingsidvar')
        conf_navigationvar = config.getboolean('Heading IDs and nav', 'navigationvar')
        conf_navigationtypepar = config.getboolean('Heading IDs and nav', 'navigationtypepar')
        conf_navigationtypebut = config.getboolean('Heading IDs and nav', 'navigationtypebut')
        conf_tooltipscheckvar = config.getboolean('Tooltips', 'tooltipscheckvar')
        conf_abbreviatetooltipsentry = config.get('Tooltips', 'abbreviatetooltipsentry')
        conf_paragraphnumbercheckvar = config.getboolean('Citability', 'paragraphnumbercheckvar')
        conf_pagenumbercheckvar = config.getboolean('Citability', 'pagenumbercheckvar')
        conf_pagenumberstartcheckvar = config.get('Citability', 'pagenumberstartcheckvar')
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

        # set fields to ini values
        self.check_body_only.setChecked(conf_bodycheckvar)
        self.check_css.setChecked(conf_csscheckvar)
        self.check_javascript.setChecked(conf_javascriptcheckvar)
        self.page_title.setText(conf_pagetitleentrytext)

        self.add_IDs.setChecked(conf_headingsidvar)
        self.create_nav.setChecked(conf_navigationvar)
        self.navigationPar.setChecked(conf_navigationtypepar)
        self.navigationBut.setChecked(conf_navigationtypebut)

        self.add_tooltips.setChecked(conf_tooltipscheckvar)
        self.tooltip_abbreviate.setText(conf_abbreviatetooltipsentry)

        self.number_paragraphs.setChecked(conf_paragraphnumbercheckvar)
        self.insert_page_no.setChecked(conf_pagenumbercheckvar)
        self.first_page.setText(conf_pagenumberstartcheckvar)

        self.detect_headings1.setText(conf_detectheadingsentry1)
        self.detect_headings2.setText(conf_detectheadingsentry2)
        self.detect_headings3.setText(conf_detectheadingsentry3)
        self.detect_images.setText(conf_detectimagesentry)
        self.images_dimensions.setText(conf_imagesdimensionsentry)
        self.detect_videos.setText(conf_detectvideosentry)
        self.video_dimensions.setText(conf_videosdimensionsentry)
        self.detect_audio.setText(conf_detectaudioentry)
        self.detect_media.setText(conf_detectmediaentry)
        self.detect_tables.setText(conf_detectblockquotesentry)
        self.detect_blockquotes.setText(conf_detecttablecaptionsentry)
        self.detect_bibliography.setText(conf_detectbibliographyentry)
        self.detect_ignore_pnum.setText(conf_detectignorepnumentry)
        self.detect_code.setText(conf_detectcodeentry)
        self.additional_styles.setText(conf_customstylemap)

        # enable/disable fields based on loaded values
        self.presetChangedFunc("Export Body")
        self.presetChangedFunc("Add IDs")
        self.presetChangedFunc("Create Navigation")
        self.presetChangedFunc("Add Tooltips")
        self.presetChangedFunc("Number Paragraphs")
        self.presetChangedFunc("Insert Page")

    def presetChangedFunc(self, name):
        '''Enables or disables fields depending on what boxes are checked.
        
        Disables the "Page title" entry, "Add suggested css?" checkbox and "Add javascript to highlight navigation while scrolling?" checkbox if "Only export the body?" is checked. Does the opposite if it's unchecked.
        
        Disables the "Create navigation?" checkbox and the "Paragraph" and "Button" radio buttons if "Add IDs to headings?" is unchecked. Does the opposite if it's checked.
        
        Disables the "Paragraph" and "Button" radio buttons if "Create navigation?" is unchecked. Does the opposite if it's checked.
        
        Disables the "Abbreviate tooltips after how many symbols?" input field if "Add tooltips to footnotes?" is unchecked. Does the opposite if it's checked.
        
        Disables the "Detect paragraphs that should not be numbered..." input field if "Number the paragraphs?" is unchecked. Does the opposite if it's checked.
        
        Disables the "Which docx page should be counted..." input field if "Insert page numbers?" is unchecked. Does the opposite if it's checked.'''

        if name == "Export Body":
            # "Only export the body?" is checked: disable
            if self.check_body_only.isChecked():
                self.check_css.setCheckable(False)
                self.check_css.setChecked(False)
                self.check_css.setStyleSheet("color: " + self.disabled_color)
                self.label_page_title.setStyleSheet("color: " + self.disabled_color)
                self.page_title.setReadOnly(True)
                self.page_title.setStyleSheet("color: " + self.disabled_color)
                self.check_javascript.setCheckable(False)
                self.check_javascript.setChecked(False)
                self.check_javascript.setStyleSheet("color: " + self.disabled_color)
            # enable
            else:
                self.check_css.setCheckable(True)
                self.check_css.setStyleSheet("color: " + self.enabled_color)
                self.page_title.setReadOnly(False)
                self.label_page_title.setStyleSheet("color: " + self.enabled_color)
                self.page_title.setStyleSheet("color: " + self.enabled_color)
                self.check_javascript.setCheckable(True)
                self.check_javascript.setStyleSheet("color: " + self.enabled_color)

            return

        if name == "Add IDs":
            # "Add IDs to headings?" is checked: enable
            if self.add_IDs.isChecked():
                self.create_nav.setCheckable(True)
                self.create_nav.setStyleSheet("color: " + self.enabled_color)
            # disable
            else:
                self.create_nav.setCheckable(False)
                self.create_nav.setStyleSheet("color: " + self.disabled_color)
                self.navigationPar.setCheckable(False)
                self.navigationPar.setStyleSheet("color: " + self.disabled_color)
                self.navigationBut.setCheckable(False)
                self.navigationBut.setStyleSheet("color: " + self.disabled_color)

            return

        if name == "Create Navigation":
            # "Create navigation?" is checked: enable
            if self.create_nav.isChecked():
                self.navigationPar.setCheckable(True)
                self.navigationPar.setStyleSheet("color: " + self.enabled_color)
                self.navigationBut.setCheckable(True)
                self.navigationBut.setStyleSheet("color: " + self.enabled_color)
                # when this function is called from readIni(), the buttons would be reset instead of staying in the configuration saved in the .ini file without this if-check
                if (self.navigationPar.isChecked() == False) and (self.navigationBut.isChecked() == False):
                    self.navigationPar.setChecked(True)
                    self.navigationBut.setChecked(False)
            # disable
            else:
                self.navigationPar.setCheckable(False)
                self.navigationPar.setStyleSheet("color: " + self.disabled_color)
                self.navigationBut.setCheckable(False)
                self.navigationBut.setStyleSheet("color: " + self.disabled_color)

            return

        if name == "Add Tooltips":
            # "Add tooltips to footnotes?" is checked: enable
            if self.add_tooltips.isChecked():
                self.tooltip_abbreviate.setReadOnly(False)
                self.tooltip_abbreviate.setStyleSheet("color: " + self.enabled_color)
                self.label_tooltip_abbreviate.setStyleSheet("color: " + self.enabled_color)
            # disable
            else:
                self.tooltip_abbreviate.setReadOnly(True)
                self.tooltip_abbreviate.setStyleSheet("color: " + self.disabled_color)
                self.label_tooltip_abbreviate.setStyleSheet("color: " + self.disabled_color)

            return

        if name == "Number Paragraphs":
            # "Number the paragraphs?" is checked: enable
            if self.number_paragraphs.isChecked():
                self.label_detect_ignore_pnum.setStyleSheet("color: " + self.enabled_color)
                self.detect_ignore_pnum.setStyleSheet("color: " + self.enabled_color)
                self.detect_ignore_pnum.setReadOnly(False)
            # disable
            else:
                self.label_detect_ignore_pnum.setStyleSheet("color: " + self.disabled_color)
                self.detect_ignore_pnum.setStyleSheet("color: " + self.disabled_color)
                self.detect_ignore_pnum.setReadOnly(True)

            return

        if name == "Insert Page":
            # "Insert page numbers?" is checked: enable
            if self.insert_page_no.isChecked():
                self.label_first_page.setStyleSheet("color: " + self.enabled_color)
                self.first_page.setStyleSheet("color: " + self.enabled_color)
                self.first_page.setReadOnly(False)

            # disable
            else:
                self.label_first_page.setStyleSheet("color: " + self.disabled_color)
                self.first_page.setStyleSheet("color: " + self.disabled_color)
                self.first_page.setReadOnly(True)

            return

    def saveOptions(self):
        '''"Save options" button:
    
        Writes current settings to the INI file and display a message stating that settings has been saved successfully.'''

        # read .ini file
        config.read(iniLocation)

        # set to new values
        config.set("Body and head", "bodyCheckVar", str(self.check_body_only.isChecked()))
        config.set('Body and head', 'csscheckvar', str(self.check_css.isChecked()))
        config.set('Body and head', 'javascriptcheckvar', str(self.check_javascript.isChecked()))
        config.set("Body and head", "pagetitleentrytext", self.page_title.text())
        config.set("Heading IDs and nav", "headingsidvar", str(self.add_IDs.isChecked()))
        config.set("Heading IDs and nav", "navigationvar", str(self.create_nav.isChecked()))
        config.set("Heading IDs and nav", "navigationtypepar", str(self.navigationPar.isChecked()))
        config.set("Heading IDs and nav", "navigationtypebut", str(self.navigationBut.isChecked()))
        config.set('Tooltips', 'tooltipscheckvar', str(self.add_tooltips.isChecked()))
        config.set('Tooltips', 'abbreviatetooltipsentry', self.tooltip_abbreviate.text())
        config.set('Citability', 'paragraphnumbercheckvar', str(self.number_paragraphs.isChecked()))
        config.set('Citability', 'pagenumbercheckvar', str(self.insert_page_no.isChecked()))
        config.set('Citability', 'pagenumberstartcheckvar', str(self.first_page.text()))
        config.set('Format templates', 'detectheadingsentry1', self.detect_headings1.text())
        config.set('Format templates', 'detectheadingsentry2', self.detect_headings2.text())
        config.set('Format templates', 'detectheadingsentry3', self.detect_headings3.text())
        config.set('Format templates', 'detectimagesentry', self.detect_images.text())
        config.set('Format templates', 'imagesdimensionsentry', self.images_dimensions.text())
        config.set('Format templates', 'detectvideosentry', self.detect_videos.text())
        config.set('Format templates', 'videosdimensionsentry', self.video_dimensions.text())
        config.set('Format templates', 'detectaudioentry', self.detect_audio.text())
        config.set('Format templates', 'detectMediaentry', self.detect_media.text())
        config.set('Format templates', 'detecttablecaptionsentry', self.detect_tables.text())
        config.set('Format templates', 'detectblockquotesentry', self.detect_blockquotes.text())
        config.set('Format templates', 'detectbibliographyentry', self.detect_bibliography.text())
        config.set('Format templates', 'detectignorepnumentry', self.detect_ignore_pnum.text())
        config.set('Format templates', 'detectcodeentry', self.detect_code.text())
        config.set('Format templates', 'customstylemap', self.additional_styles.toPlainText())

        # write .ini file
        with open(iniLocation, "w") as configFile:
            config.write(configFile)

        QMessageBox.information(self, "Saved", "Settings have been saved.")

        return

    def resetOptions(self):
        '''"Reset options" button:
    
        Asks the user if they really want to reset the current options.

        If yes: Resets options to their factory state. Updates the GUI, writes new settings to the INI file and displays a message stating that settings have been reset successfully.'''

        reallyReset = QMessageBox.question(self, 'Reset options', 'Are you sure you want to reset the options to their factory settings? Current settings will be overwritten.', QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)

        if reallyReset == QMessageBox.StandardButton.Yes:
            # read .ini file
            config.read(iniLocation)

            # reset to default values
            config.set("Body and head", "bodyCheckVar", "True")
            config.set('Body and head', 'csscheckvar', "False")
            config.set('Body and head', 'javascriptcheckvar', "False")
            config.set("Body and head", "pagetitleentrytext", "")
            config.set("Heading IDs and nav", "headingsidvar", "True")
            config.set("Heading IDs and nav", "navigationvar", "False")
            config.set("Heading IDs and nav", "navigationtypepar", "True")
            config.set("Heading IDs and nav", "navigationtypebut", "False")
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

            # Update GUI with reset values. Disabled/enabled states will be updated automatically
            self.check_body_only.setChecked(True)
            self.check_css.setChecked(False)
            self.check_javascript.setChecked(False)
            self.page_title.clear()

            self.add_IDs.setChecked(True)
            self.create_nav.setChecked(False)
            self.navigationPar.setChecked(True)
            self.navigationBut.setChecked(False)

            self.add_tooltips.setChecked(True)
            self.tooltip_abbreviate.setText("500")

            self.number_paragraphs.setChecked(True)
            self.insert_page_no.setChecked(False)
            self.first_page.setText("1")

            self.detect_headings1.setText("FVMW Heading")
            self.detect_headings2.setText("FVMW Heading2")
            self.detect_headings3.setText("FVMW Heading3")
            self.detect_images.setText("FVMW Image")
            self.images_dimensions.setText("")
            self.detect_videos.setText("FVMW Video")
            self.video_dimensions.setText("")
            self.detect_audio.setText("FVMW Audio")
            self.detect_media.setText("FVMW Media")
            self.detect_tables.setText("FVMW Blockquote")
            self.detect_blockquotes.setText("FVMW TableCaption")
            self.detect_bibliography.setText("FVMW Bibliography")
            self.detect_ignore_pnum.setText("FVMW IgnorePNum")
            self.detect_code.setText("FVMW Code")
            self.additional_styles.clear()

            QMessageBox.information(self, "Reset successful", "Settings have been reset to original values.")

            return

    def inputPathFunc(self):
        '''"Browse" button:

        Prompts the user to choose the file that shall be converted and replaces the input path field text with the path of the chosen file.'''

        # delete current text in input field
        self.display_file_path.clear()
        
        # get input path
        self.inputPath = QFileDialog.getOpenFileName()[0]

        # insert input path into field
        if (self.inputPath != None) and (self.inputPath != ""):
            self.display_file_path.setText(self.inputPath)
            self.display_file_path.end(True) # scroll view to the right-most part of the path text

        return

    def submitFunc(self):
        '''"Convert" button:

        Prompts the user to choose the output path. Starts the conversion process by calling "convertAndExport()".\n
        Throws an error if no input file has been chosen yet.'''

        # get output path
        self.outputPath = QFileDialog.getSaveFileName()[0]

        # start conversion process
        if self.inputPath != None:
            if (self.outputPath != None) and (self.outputPath != ""):
                self.convertAndExport()
        else:
            QMessageBox.warning(self, "No input file given", "Choose an input file.")

        return

    def convertAndExport(self):
        '''Converts a DOCX file to an HTML file and exports it by calling functions from SciDocx2WebConversion.py. Displays a "Success" message if conversion was successful.'''

        # style map
        custom_style_map = SciConvert.style_map_func("", self.detect_headings1.text(), self.detect_headings2.text(), self.detect_headings3.text(), self.detect_images.text(), self.detect_videos.text(), self.detect_audio.text(), self.detect_media.text(), self.detect_blockquotes.text(), self.detect_tables.text(), self.detect_bibliography.text(), self.detect_ignore_pnum.text(), self.number_paragraphs.isChecked(), self.detect_code.text(), self.additional_styles.toPlainText())

        # import and enclose input file with tags
        input = mammoth.convert_to_html(self.inputPath, style_map=custom_style_map).value
        bodyxml = SciConvert.enclose_body(input, self.check_body_only.isChecked(), self.page_title.text())

        # remove unwanted links that the text processor adds
        bodyxml = SciConvert.remove_empty_elements(bodyxml)

        # create footnotes
        footnotes = SciConvert.create_footnotes_list(bodyxml, self.tooltip_abbreviate.text())

        # abbreviate footnotes
        footnotesAbbr = SciConvert.abbreviate_footnotes(footnotes, self.tooltip_abbreviate.text())

        # add wbr to footnotes
        footnotesAbbr = SciConvert.add_wbr_footnotes(footnotesAbbr, self.tooltip_abbreviate.text())

        # insert footnotes into main text
        bodyxml = SciConvert.insert_footnotes(self.add_tooltips.isChecked(), bodyxml, footnotesAbbr)

        # adjust footnote sups
        bodyxml = SciConvert.adjust_footnotes(self.add_tooltips.isChecked(), bodyxml)

        # separate bottom footnotes
        commentBottomFootnotes = etree.Comment(' Bottom footnotes ')
        breakElement = etree.XML('<br/>')
        hrElement = etree.XML('<hr/>')
        bodyxml = SciConvert.footnotes_bottom_adjust(bodyxml, commentBottomFootnotes, breakElement, hrElement)

        # add wbr to main text
        bodyxml = SciConvert.add_wbr_text(bodyxml)

        # add heading IDs
        bodyxml = SciConvert.add_Head_IDs(self.add_IDs.isChecked(), bodyxml)

        # create navigation
        findH1 = bodyxml.xpath('.//*[self::h1 or self::h2 or self::h3]')
        navigationElement = etree.Element('nav')
        commentNavigation = etree.Comment(' Navigation ')
        h1Navigation = etree.Element('h1')
        h1Navigation.text = 'Navigation'
        navGridDiv = etree.Element('div')
        navGridDiv.attrib['class'] = 'navGrid'
        navGridDiv = SciConvert.create_navigation(self.create_nav.isChecked(), self.navigationPar.isChecked(), self.navigationBut.isChecked(), findH1, navigationElement, commentNavigation, h1Navigation, navGridDiv)

        # add cite to blockquotes
        tooltiptextPath = './/a[contains(@id, "footnote-ref")]/sup'
        bodyxml = SciConvert.add_cite(tooltiptextPath, bodyxml, footnotes)

        # embed images
        bodyxml = SciConvert.embed_images(bodyxml, self.images_dimensions.text())

        # embed videos
        bodyxml = SciConvert.embed_videos(bodyxml, self.video_dimensions.text())

        # embed audio
        bodyxml = SciConvert.embed_audio(bodyxml)

        # add file insertion messages above mediacaptions
        bodyxml = SciConvert.file_insertion_message(bodyxml)

        # move table captions
        bodyxml = SciConvert.move_table_caption(bodyxml)

        # create page breaks
        bodyxml = SciConvert.page_breaks(self.insert_page_no.isChecked(), self.first_page.text(), bodyxml, self.check_body_only.isChecked())

        # number paragraphs
        bodyxml = SciConvert.paragraph_numbering(self.number_paragraphs.isChecked(), bodyxml)

        # create sections (haven't been able to figure this out yet)
        #bodyxml = SciConvert.create_sections(bodyxml)

        # assemble file
        exportableBodyxml = SciConvert.assemble_html(self.create_nav.isChecked(), self.check_body_only.isChecked(), self.check_css.isChecked(), cssXML, navGridDiv, bodyxml, javascriptXML, self.check_javascript.isChecked())

        # unescape and escape HTML characters
        exportableBodyxml = SciConvert.escape_unescape(exportableBodyxml)

        # write file
        SciConvert.write_html(exportableBodyxml, self.outputPath)

        QMessageBox.information(self, "Success", "The file has been converted successfully.")

        return



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
tablecss = 'table {margin-top: 28px;}\ntable, th, td {border: 1px solid;}\ntd {padding: 0px 5px;}\ncaption {caption-side: bottom; text-align: left; font-size: 15px; color: #454545;}'

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

### CREATE GUI WINDOW ###
app = QApplication()

window = MainWindow()
window.show()

app.exec()