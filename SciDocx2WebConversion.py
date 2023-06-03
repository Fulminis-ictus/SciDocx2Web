"""
Convert scientific papers in DOCX format to HTML. See this project's GitHub page for more info: https://github.com/Fulminis-ictus/SciDocx2Web

This is a module that is called by SciDocx2WebUI. It handles all of the actual conversion functions.

Documentation last updated: 2023.05.16\n
Author: Tim Reichert\n
Version: 1.0 (first public release)

Uses and is dependent on Mammoth: https://github.com/mwilliamson/python-mammoth\n
Makes use of dwasyl's added page break detection functionailty: https://github.com/dwasyl/python-mammoth/commit/38777ee623b60e6b8b313e1e63f12dafd82b63a4
"""

### IMPORTS ###
# Extraction and Conversion modules
import re # RegEx
from lxml import etree # XML and XPath
from html import unescape # Replace HTML entities with their actual symbols
from html import escape # Re-add escape characters in code inside <code> tags

# GUI
from tkinter import messagebox


### MAIN CODE ###
## STYLE MAP
def style_map_func(custom_style_map, headings1, headings2, headings3, images, videos, audio, media, blockquotes, tableCaptions, bibliography, ignorePNum, paragraphNumberCheck, code, addStyleMap):
    '''Generates a style map based on the format template name input fields. Empty input fields are ignored. The style map detects format templates applied to text and encloses them with an html element. For more information see: https://github.com/mwilliamson/python-mammoth#custom-style-map
    
    Headings1 -> h1:fresh\n
    Headings2 -> h2:fresh\n
    Headings3 -> h3:fresh\n
    Image -> img.insertimage:fresh\n
    Video -> iframe.insertvideo:fresh\n
    Audio -> audio.insertaudio:fresh\n
    Media -> p.mediacaption:fresh\n
    Blockquotes -> blockquote:fresh\n
    Table Captions -> caption:fresh\n
    Bibliography -> p.bibliography:fresh\n
    Ignore Paragraph Numbering -> p.ignorePNum:fresh\n
    Code -> code'''

    if headings1 != "":
        custom_style_map += f"p[style-name='{headings1}'] => h1:fresh"
    if headings2 != "":
        custom_style_map += f"\np[style-name='{headings2}'] => h2:fresh"
    if headings3 != "":
        custom_style_map += f"\np[style-name='{headings3}'] => h3:fresh"
    if images != "":
        custom_style_map += f"\np[style-name='{images}'] => img.insertimage:fresh"
    if videos != "":
        custom_style_map += f"\np[style-name='{videos}'] => iframe.insertvideo:fresh"
    if audio != "":
        custom_style_map += f"\np[style-name='{audio}'] => audio.insertaudio:fresh"
    if media != "":
        custom_style_map += f"\np[style-name='{media}'] => p.mediacaption:fresh"
    if blockquotes != "":
        custom_style_map += f"\np[style-name='{blockquotes}'] => blockquote:fresh"
    if tableCaptions != "":
        custom_style_map += f"\np[style-name='{tableCaptions}'] => caption:fresh"
    if bibliography != "":
        custom_style_map += f"\np[style-name='{bibliography}'] => p.bibliography:fresh"
    if (ignorePNum != "") and (paragraphNumberCheck):
        custom_style_map += f"\np[style-name='{ignorePNum}'] => p.ignorePNum:fresh"
    if code != "":
        custom_style_map += f"\np[style-name='{code}'] => code:fresh"
    if addStyleMap != "":
        custom_style_map += "\n" + addStyleMap

    return custom_style_map

## IMPORT INPUT FILE
def enclose_body(input, bodyCheckVar, pageTitleEntryText):
    '''Encloses the imported file with <body> tags to make it navigable with xpath.

    If "Only export the body?" is unchecked:\n
    Adds <html>, <head> and <title> elements. Also sets the charset to UTF-8 and adds the page title if the "Page title:" field isn't empty.
    
    Additionally, the grid container and main grid <div> tags are inserted and new page markers are unescaped.'''

    # "Only export the body?" is checked
    if bodyCheckVar:
        bodyxml = '<body><div class="gridContainer"><div class="mainGrid">' + input + '</div></div></body>'

    # "Only export the body?" is unchecked and the page title field isn't empty
    elif pageTitleEntryText != "":
        bodyxml = f'<html><head><title>{pageTitleEntryText}</title><meta charset=\"UTF-8\"></meta></head><body><div class="gridContainer"><div class="mainGrid">' + input + '</div></div></body></html>'

    # "Only export the body?" is unchecked and the page title field is empty
    else:
        bodyxml = f'<html><head><meta charset=\"UTF-8\"></meta></head><body><div class="gridContainer"><div class="mainGrid">' + input + '</div></div></body></html>'

    # unescape new page markers
    bodyxml = re.sub(r'&lt;sub class=&quot;pagenumber&quot;&gt;NEW_PAGE_BEGINNING!&lt;/sub&gt;', r'<sub class="pagenumber">NEW_PAGE_BEGINNING!</sub>', bodyxml)

    bodyxml = etree.fromstring(bodyxml)

    return bodyxml

## FOOTNOTES
def create_footnotes_list(bodyxml, abbreviateFootnotesNumber):
    '''Compiles a list of footnotes. Iterates over all <li> elements in the footnote list created by mammoth and saves their contents to a python list.

    There was one case where footnotes started counting at 0 instead of 1. This case has been accounted for by checking whether the list is empty after searching for list items with the ID "footnote-1".

    An empty footnote list is created if there are no footntoes.
    
    If footnotes are abbreviated:\n
    Get footnote text without HTML tags to prevent tags that are never closed due to the abbreviation process.
    
    Else:\n
    Get the footnote text with all HTML tags.'''

    footnotes = []

    if bodyxml.xpath('.//li[@id="footnote-1"]/..') == []:
        footnotePath = bodyxml.xpath('.//li[@id="footnote-0"]/..')
    else:
        footnotePath = bodyxml.xpath('.//li[@id="footnote-1"]/..')

    for ol in footnotePath:
        for li in ol:
            for p in li:
                # Get footnote text without HTML tags
                if abbreviateFootnotesNumber != "":
                    footnoteText = ''.join(p.itertext())
                    footnoteText = re.sub(r' ↑', r'', footnoteText)
                    footnotes.append(footnoteText)
                # Get footnote text with HTML tags
                else:
                    footnoteText = etree.tostring(p).decode('utf-8')
                    footnoteText = re.sub(r'(<p>)(.*?)(<a href="#footnote-ref)(.*)', r'\2', footnoteText)
                    footnotes.append(footnoteText)

    return footnotes

def abbreviate_footnotes(footnotes, abbreviateFootnotesNumber):
    '''Abbreviates footnotes according to the number in the "Abbreviate tooltips after how many symbols?" input field and adds "[...]" to the end of abbreviated footnotes.

    If the input field is empty:\n
    Skips this function.

    If something other than an integer is input:\n
    Tooltip abbreviation is skipped with an error message indicating that the abbreviation was unsuccessful but the rest of the conversion process continues as normal.'''

    if abbreviateFootnotesNumber != "":
        if abbreviateFootnotesNumber.isdigit():
            abbreviateFootnotesNumber = int(abbreviateFootnotesNumber)
            for i in range(len(footnotes)):
                if len(footnotes[i]) > abbreviateFootnotesNumber:
                    footnotes[i] = footnotes[i][:abbreviateFootnotesNumber] + '[...]'
        else:
            messagebox.showerror('Tooltip abbreviation unsuccessful', 'The file conversion will continue but the tooltip abbreviation was unsuccessful. The tooltip abbreviation input field ("Abbreviate tooltips after how many symbols?") only accepts integers.')

    return footnotes

def add_wbr_footnotes(footnotesAbbr, abbreviateFootnotesNumber):
    '''Finds all slashes within footnotes and adds <wbr> tags after them to ensure that links automatically receive a line break when necessary. 
    
    If footnotes are abbreviated:\n
    Apply to all text. Since all HTML tags have been removed in a previous step there's no need to worry about HTML tags being affected.
    
    Else:\n
    Look for text in <a> tags specifically so that other HTML tags aren't affected.'''

    # insert <wbr> throughout the whole text
    if abbreviateFootnotesNumber != "":
        for i in range(len(footnotesAbbr)):
            footnotesAbbr[i] = re.sub(r'/', r'/<wbr>', footnotesAbbr[i])

    # insert <wbr> only inside the <a> elements
    else:
        for i in range(len(footnotesAbbr)):
            if re.compile(r'(<a.*?>)').search(footnotesAbbr[i]): # check if a link exists in the footnote before proceeding.
                replace = re.sub(r'(.*?)(<a.*?>)(.*?)(</a>)(.*)', r'\3', footnotesAbbr[i])
                replace = re.sub(r'/', r'/<wbr>', replace)
                footnotesAbbr[i] = re.sub(r'(<a.*?>)(.*?)(</a>)', r'\1' + replace + r'\3', footnotesAbbr[i])

    return footnotesAbbr

def insert_footnotes(tooltipsCheckVar, bodyxml, footnotesAbbr):
    '''If "Add tooltips to footnotes?" is checked:\n
    Finds <sup> tags that contain links with an ID starting with "footnote-ref", which denotes footnote numbers in the main text. It then creates a tooltip <span> and appends it to the end of the found <sup> tags. The whole element, including the <sup> element and tooltip <span> element, is then enclosed with a tooltippop <span> element. Structure derived from: https://www.w3schools.com/howto/howto_css_tooltip.asp
    
    Also adds "role" and "aria-attribute" attributes for accessibility.
    
    Conversion example: <sup><a href="#footnote-N" id="footnote-ref-N">[N]</a></sup> ->\n
    <span class="tooltippop" aria-describedby="tooltip-N"><sup><a href="#footnote-N" id="footnote-ref-N">[N]</a><span role="tooltip" id="tooltip-N">Footnote content.</span></sup></span>'''

    if tooltipsCheckVar:
        i = 0
        for node in bodyxml.xpath('//sup/a[contains(@id, "footnote-ref")]/..'): 
            # create tooltip and tooltip text <span>s
            tooltiptextSpan = etree.Element('span')
            tooltiptextSpan.attrib['role'] = 'tooltip'
            tooltiptextSpan.attrib['id'] = 'tooltip-' + str(i + 1)
            tooltipSpan = etree.Element('span')
            tooltipSpan.attrib['class'] = 'tooltippop'
            tooltipSpan.attrib['aria-describedby'] = 'tooltip-' + str(i + 1)

            # insert footnote text and append tooltiptext <span> to end of current footnote sections 
            tooltiptextSpan.text = footnotesAbbr[i]
            node.append(tooltiptextSpan)

            # enclose footnote sections (including the tooltiptext <span>s) with tooltip <span>s
            tooltipSpan.extend(node)
            node.append(tooltipSpan)
            i += 1

    #print(etree.tostring(bodyxml))
    return bodyxml

def adjust_footnotes(tooltipsCheckVar, bodyxml):
    '''If "Add tooltips to footnotes?" is checked:\n
    Removes the <sup> elements and creates new ones within the <a> tags. Also removes the square brackets around the footnote numbers, and moves the numbers from the <a> tags into the <sup> tags.\n
    This whole process ensures that only the footnote numbers, not the footnote text within the tooltips, is enclosed by the <sup> elements.

    Example: <sup><span class="tooltip"><a href="#footnote-N" id="footnote-ref-N">[N]</a><span class="tooltiptext">Footnote content.</span></span></sup> ->\n
    <span class="tooltippop" aria-describedby="tooltip-N"><a href="#footnote-N" id="footnote-ref-N"><sup>N</sup></a><span role="tooltip" id="tooltip-N">Footnote content.</span></span>
    
    Else:\n
    Only the square brackets around the footnote numbers are removed.'''

    if tooltipsCheckVar:
        # remove <sup> elements
        while bodyxml.xpath('//sup/span[contains(@class, "tooltippop")]/..'):
            for node in bodyxml.xpath('//sup/span[contains(@class, "tooltippop")]/..'):
                for a in node:
                    node.addnext(a)
                node.getparent().remove(node)
        
        # add <sup> elements containing the footnote numbers at the right position
        for a in bodyxml.xpath('//a[contains(@id, "footnote-ref")]'):
            # transform text inside <a> from [N] to N, give it an aria-label, insert it into the <sup> elements, remove it from the <a> element
            a.text = re.sub(r'(\[)(.*)(\])', r'\2', a.text)
            a.attrib['label'] = 'Link to footnote list at the bottom of the document.'
            sup = etree.Element('sup')
            sup.text = a.text
            a.text = ''

            # move <sup> elements to the right positions
            sup.extend(a)
            a.append(sup)
    
    else:
        for node in bodyxml.xpath('//a[contains(@id, "footnote-ref")]'):
            # transform text from [N] to N
            node.text = re.sub(r'(\[)(.*)(\])', r'\2', node.text)

    return bodyxml

def footnotes_bottom_adjust(bodyxml, commentBottomFootnotes, breakElement, hrElement):
    '''If a footnote list at the bottom of the main text exists:

    Adds elements before the footnote list at the bottom to separate it from the rest: <br/>, <hr/> and the comment "Bottom footnotes".
    
    Also adds "aria-label" attributes that describe the links at the end of the footntoes as links back to the footnotes in the main text.'''

    bottomFootnotes = bodyxml.find('.//li[@id="footnote-1"]/..')

    if bottomFootnotes != None:
        # add seperators
        bottomFootnotes.addprevious(commentBottomFootnotes)
        bottomFootnotes.addprevious(breakElement)
        bottomFootnotes.addprevious(hrElement)
        bottomFootnotes.addprevious(breakElement)
        
        # add aria-label
        for li in bottomFootnotes:
            for p in li:
                for a in p.xpath('//a[contains(@href, "footnote-ref")]'):
                    a.attrib['aria-label'] = 'Link back to footnote in main text.'

    return bodyxml

## TEXT
def add_wbr_text(bodyxml):
    '''Finds all slashes in links within the main text and adds <wbr> tags to ensure that links automatically receive line breaks after slashes.'''

    for a in bodyxml.xpath('//a'):
        if a.text != None:
            a.text = re.sub(r'/', r'/<wbr>', a.text)

    return bodyxml

def add_Head_IDs(headingsIDVar, bodyxml):
    '''If "Add IDs to headings?" is checked:\n
    Adds IDs to all found <h1> elements following the form "headingN" where "N" is an integer counting upwards.
    
    Example: <h1 id="heading1">First heading</h1>, <h1 id="heading2">Second heading</h1>, <h2 id="heading3">Subheading of second heading</h2> etc.'''

    if headingsIDVar:
        i = 1
        for node in bodyxml.xpath('//h1|//h2|//h3'):
            node.attrib['id'] = 'heading' + str(i)
            i += 1

    return bodyxml

def remove_word_lnks(bodyxml):
    '''Removes various links that word inserts into the document, including "Table Of Contents", "_heading" and "_Hlk" link. They are interpreted as non-closed "a"-tags that can lead to display errors or mess with the navigation.'''

    def repeat_removal(findID, bodyxml):
        if findID != None:
            for node in findID:
                for a in node:
                    tail = a.tail
                    node.remove(a)
                    node.text = tail
        
        return bodyxml

    findID = bodyxml.xpath('.//a[contains(@id, "_Toc")]/..')
    repeat_removal(findID, bodyxml)

    findID = bodyxml.xpath('.//a[contains(@id, "_heading")]/..')
    repeat_removal(findID, bodyxml)

    findID = bodyxml.xpath('.//a[contains(@id, "_Hlk")]/..')
    repeat_removal(findID, bodyxml)

    return bodyxml

def create_navigation(navigationVar, navigationTypeVar, findH1, navigationElement, commentNavigation, h1Navigation, navGridDiv):
    '''If "Create navigation?" is checked:\n
    Creates a navigation by compiling a list of all <h1>, <h2> and <h3> elements. Each navigation item is a paragraph or a button, depending on which radio button option is activated in the GUI. Adds links with href attributes that link each navigation item to their respective heading.
    
    Example: <button><a href="#heading1">Navigation to first heading</a></button>

    When compiling the headings text, the <h1>, <h2> and <h3> tags and new page markers are removed from the navigation.
    
    Encloses everything with a <div> with the class "navGrid".'''

    if navigationVar:
        if findH1 != None:
            i = 1
            for node in findH1:
                # create <p> or <button>
                if navigationTypeVar == 'paragraph':
                    elementPorB = etree.Element('p')
                else:
                    elementPorB = etree.Element('button')
                # create link
                a = etree.SubElement(elementPorB, 'a')
                a.attrib['href'] = '#heading' + str(i)
                # remove <h1>, <h2> and <h3> tags, as well as new page markers
                headingsText = etree.tostring(node).decode('utf-8')
                headingsText = re.sub(r'(<h1 .*?>)(.*?)(</h1>)', r'\2', headingsText)
                headingsText = re.sub(r'(<h2 .*?>)(.*?)(</h2>)', r'\2', headingsText)
                headingsText = re.sub(r'(<h3 .*?>)(.*?)(</h3>)', r'\2', headingsText)
                headingsText = re.sub(r'(<sub class="pagenumber">NEW_PAGE_BEGINNING!</sub>)', r'', headingsText)
                a.text = headingsText
                navigationElement.append(elementPorB)
                i += 1
            
            # create navigation grid
            navGridDiv.append(commentNavigation)
            navGridDiv.append(h1Navigation)
            navGridDiv.append(navigationElement)

    return navGridDiv

def add_cite(tooltiptextPath, bodyxml, footnotes):
    '''Adds a cite attribute to <blockquote> elements. Inserts the footnote text as the cite value by using the footnote number at the end of the blockquote (minus one since the list starts counting at 0 but the footnotes start counting at 1) as the index of the footnote list.'''

    for node in bodyxml.xpath('//blockquote'):
        if node.xpath(tooltiptextPath):
            footnotenumber = int(node.xpath(tooltiptextPath)[0].text)
            citetext = footnotes[footnotenumber - 1]
            node.attrib['cite'] = citetext

    return bodyxml

def embed_images(bodyxml, dimensions):
    '''Embeds image links that have the "insertimage" attribute. The text marked as an image should be a link (not a hyperlink!), which is then inserted into the "src" attribute of the image.
    
    Else if the height and width input is empty:\n
    Don't insert any width and height parameters.
    
    Else:\n
    Don't insert any width and height parameters and display an error.'''

    # get width and height
    splitDimensions = dimensions.split(',')
    # correct input 
    if len(splitDimensions) == 2 and splitDimensions[0].isdigit() and splitDimensions[1].isdigit():
        width = splitDimensions[0]
        height = splitDimensions[1]
    # input field is empty
    elif (len(splitDimensions) == 1) and (splitDimensions[0] == ''):
        width = None
        height = None
    # faulty input
    else:
        width = None
        height = None
        messagebox.showinfo('No image dimensions used', 'Faulty input in the audio dimensions input field. No "width" and "height" parameters have been inserted. Make sure the input is two integers seperated by a comma.')

    # insert src and width and height
    for node in bodyxml.xpath('//img[@class="insertimage"]'):
        node.attrib['src'] = node.text
        if width != None:
            node.attrib['width'] = width
        if height != None:
            node.attrib['height'] = height
        node.text = ''

    return bodyxml

def embed_videos(bodyxml, dimensions):
    '''Embeds video links that have the "insertvideo" attribute. The text marked as a video should be a link (not a hyperlink!), which is then inserted into the "src" attribute of the video.
    
    If the height and width input consists of two numbers split by a comma:\n 
    Use these two numbers as height and width of the iframe.

    Else if the height and width input is empty:\n
    Don't insert any width and height parameters.
    
    Else:\n
    Don't insert any width and height parameters and display an error.'''

    # get width and height
    splitDimensions = dimensions.split(',')

    # correct input 
    if len(splitDimensions) == 2 and splitDimensions[0].isdigit() and splitDimensions[1].isdigit():
        width = splitDimensions[0]
        height = splitDimensions[1]
    # input field is empty
    elif (len(splitDimensions) == 1) and (splitDimensions[0] == ''):
        width = None
        height = None
    # faulty input
    else:
        width = None
        height = None
        messagebox.showinfo('No video dimensions used', 'Faulty input in the video dimensions input field. No "width" and "height" parameters have been inserted. Make sure the input is two integers seperated by a comma.')

    # insert src and width and height
    for node in bodyxml.xpath('//iframe[@class="insertvideo"]'):
        node.attrib['src'] = node.text
        if width != None:
            node.attrib['width'] = width
        if height != None:
            node.attrib['height'] = height
        node.text = ''

    return bodyxml

def embed_audio(bodyxml):
    '''Embeds audio links that have the "insertaudio" attribute. The text marked as audio should be a link (not a hyperlink!), which is then inserted into the "src" attribute of the <source> element inside the <audio> element.'''

    # create <source> and insert src
    for node in bodyxml.xpath('//audio[@class="insertaudio"]'):
        source = etree.Element('source')
        source.attrib['src'] = node.text
        node.append(source)
        node.text = ''
        node.attrib['controls'] = 'True'

    return bodyxml

def file_insertion_message(bodyxml):
    '''Adds the comment "Insert Media" before <p> elements that have the media caption class to alert the user to the fact that they might need to manually insert media at that line.'''

    for node in bodyxml.xpath('//p[contains(@class, "mediacaption")]'):
        commentMedia = etree.Comment(' Insert Media ')
        node.addprevious(commentMedia)

    return bodyxml

def move_table_caption(bodyxml):
    '''Moves the <caption> elements to the beginning of the <table> elements for semantic reasons.'''

    for node in bodyxml.xpath('//caption'):
        if node.getprevious().tag == 'table':
            node.getprevious().insert(0, node)

    return bodyxml

def page_breaks(pageNumberCheckVar, pageNumberStartCheckVar, bodyxml, bodyCheckVar):
    '''If "Insert page numbers?" is checked\n
    Inserts a <sub class="pagenumber"> element at the very beginning of the text (only page breaks are marked automatically, meaning the first page needs to receive an element manually). Finds all <sub class="pagenumber"> elements and inserts a page number following the form {N} where N is an integer counting upwards. N is calculated as 2-pageNumberStart. If the page number start is set to 1 then the page number at the very top would receive the number 1. If the page number start is set to 2 then the page number at the very top would receive the number 0 etc. Page number indicators that would receive a number <= 0 receive no text instead, which makes them invisible.
    
    Else:\n
    Finds <sub class="pagenumber"> tags that denote a new page beginning and deletes their content.'''

    if pageNumberCheckVar and pageNumberStartCheckVar != "":
        # insert page number element at the very beginning (<sub class="pagenumber"></sub>). Will receive text in next step
        pageSub = etree.Element('sub')
        pageSub.attrib['class'] = 'pagenumber'

        if bodyCheckVar:
            bodyxml[0].insert(0, pageSub)
        else:
            bodyxml[1][0][0].insert(0, pageSub)

        # insert page number into beginning of new page markers. Insert no text if 
        if pageNumberStartCheckVar.isdigit():
            i = 2 - int(pageNumberStartCheckVar)

            for pageNumber in bodyxml.xpath('.//sub[@class="pagenumber"]'):
                if i <= 0:
                    pageNumber.text = ''
                else:
                    pageNumber.text = '{' + str(i) + '}'
                i += 1
        else:
            messagebox.showerror('Page number insertion unsuccessful', 'The file conversion will continue but the page number insertion was unsuccessful. The starting page number input field only accepts integers.')
            for pageNumber in bodyxml.xpath('.//sub[@class="pagenumber"]'):
                pageNumber.text = ''
    else:
        for pageNumber in bodyxml.xpath('.//sub[@class="pagenumber"]'):
                pageNumber.text = ''

    return bodyxml

def paragraph_numbering(paragraphNumberCheckVar, bodyxml):
    '''If "Number the paragraphs?" is checked:\n
    Adds numbers to the beginning of each paragraph and each blockquote following the form [N] where N is an integer counting upwards. Skips paragraphs that have the "ignorePNum", "mediacaption" or "bibliography" classes.'''

    if paragraphNumberCheckVar:
        i = 1
        for node in bodyxml.xpath('//div[@class="mainGrid"]/p[not(@class="mediacaption") and not(@class="ignorePNum") and not(@class="bibliography")]|//blockquote'):
            if node.text != None:
                node.text = '[' + str(i) + '] ' + node.text
                i += 1
            else:
                node.text = '[' + str(i) + '] '
                i += 1

    return bodyxml

def create_sections(bodyxml):
    '''This code was planned to create sections based on chapters but it's not functional.'''

    # find h1, h2 or h3 headings.
    # insert everything between found heading and next heading into section.

    #for heading in bodyxml.xpath('.//h1|.//h2|.//h3'):
    #    print(etree.tostring(heading))

    #headings = bodyxml.findall(".//h1")
    findH1 = bodyxml.findall('.//h1')

    for heading in findH1:
        elements = bodyxml.xpath('.//' + heading.tag + '[@id="' + heading.attrib["id"] + '"]/following-sibling::h1[1]/preceding-sibling::*')
        print(elements)
        elements.reverse()

        section = etree.Element('section')
        bodyxml.xpath('.//' + heading.tag + '[@id="' + heading.attrib["id"] + '"]')[0].addprevious(section)

        for element in elements:
            section.insert(0, element)
        
        #print(bodyxml.getelementpath(heading))
    #print(findH1)
    #headings = bodyxml.xpath('.//h1/following-sibling::h1[1]/preceding-sibling::*')
    #print(headings)

    return bodyxml


## FINALISATION
def assemble_html(navigationVar, bodyCheckVar, cssCheckVar, cssXML, navGridDiv, bodyxml, javascriptXML, javascriptCheckVar):
    '''Assembles the individual sections (navigation, css, body) and writes them to an html file.
    
    If "Create navigation?" is unchecked:\n
    Skips navigation section.
    
    If "Add suggested css?" is unchecked:\n
    Skips css section.
    
    If "Only export the body?" is unchecked:\n
    "<!DOCTYPE html>" is added to the beginning to create a well formed html file. This is not necessary if only the body is exported since it's an incomplete html file without the <head> element.'''

    # navigation
    if navigationVar: 
        if bodyCheckVar:
            bodyxml[0].insert(0, navGridDiv)
        else:
            bodyxml[1][0].insert(0, navGridDiv)

    # javascript
    if javascriptCheckVar:
        bodyxml[0].insert(0, javascriptXML)

    # css
    if cssCheckVar:
        bodyxml[0].insert(0, cssXML)

    # doctype declaration
    if bodyCheckVar:
        exportableBodyxml = etree.tostring(bodyxml, encoding='unicode', pretty_print=True)
    else:
        exportableBodyxml = "<!DOCTYPE html>\n" + etree.tostring(bodyxml, encoding='unicode', pretty_print=True)

    return exportableBodyxml

def escape_unescape(exportableBodyxml):
    '''Unescapes the file to make sure HTML tags are applied properly instead of them being displayed as escaped HTML. Re-escapes HTML symbols that are marked as example code.'''

    # unescape
    exportableBodyxml = unescape(exportableBodyxml)

    # re-escape code
    def insertPreCode(match):
        return r'<code>' + escape(match.group(2)) + r'</code>'
    exportableBodyxml = re.sub(r'(<code>)(.*?)(</code>)', insertPreCode, exportableBodyxml)

    return exportableBodyxml

def write_html(exportableBodyxml, outputPath):
    '''Write to file.'''
    
    with open(outputPath, 'w', encoding='utf-8') as f:
        f.write(exportableBodyxml)
        f.close()

    return