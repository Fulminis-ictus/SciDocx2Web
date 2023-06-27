# SciDocx2Web

Convert your scientific articles from DOCX to HTML.

* [What is this](#what-is-this)
* [How to use this tool](#how-to-use)
    * [GUI](#gui)
    * [CSS](#css)
* [Known problems or missing features - Found a bug?](#known-problems)
* [FAQ](#faq)
* [Acknowledgements](#acknowledgements)

## What is this? <a name="what-is-this"></a>

SciDocx2Web allows you to convert your academic DOCX articles to HTML code. This tool has been tested with Word and LibreOffice. Support for other text editors can not be guaranteed. The output HTML code possesses several features that could be convenient for scientific publications and that semantically enhance it. Included features are:

- Footnotes:
    - Tooltips that contain the footnote text are displayed above footnote numbers.
    - Footnote tooltips can be abbreviated to prevent them from going beyond web page boundaries.
- Navigation:
    - Each heading receives an ID.
    - A navigation that either consists of paragraphs or buttons is created. It links to the respective headings.
    - The chapter you're scrolling through is highlighted in the navigation.
    - The text and navigation are separated into two sections. The navigation is sticky, making it scroll alongside the main text.
- Citability:
    - The start of each new page is marked for easier cross referencability between the DOCX file and the HTML page.
    - Each paragraph is numbered.
- Predefined custom style map to mark the following elements in the HTML code:
    - Headings (first, second and third level).
    - Image embeds.
    - Video embeds.
    - Audio embeds.
    - Media captions.
    - Table captions.
    - Block quotes.
    - Bibliography.
    - Code.
- Slashes in links and in footnotes are detected and wbr tags are inserted to prevent them from going beyond page boundaries.
- Cite attributes are inserted into block quotes. These cite attributes contain the footnote text of said block quote.
- Table captions are moved into the table elements for semantic reasons.
- Footnotes, footnote tooltips and links between footnotes and the footnote list at the bottom of the page receive aria and role attributes for increased accessibility.
- Predefined CSS that can be used as a basis.

The GUI has settings that let you choose which features to include or not to include in your output.

This project was created as part of the publication plans of the [Forschungsgemeinschaft VideospielMusikWissenschaft](https://videospielmusikwissenschaft.de) (FVMW, Research Group for Video Game Music Studies). We decided that having footnote tooltips and some form of citability aid would benefit our publications. You can find examples with similar features [here](https://www.mtosmt.org/issues/mto.19.25.3/mto.19.25.3.medina.gray.html) or [here](https://zfdg.de/wp_2021_001). Incorporating said features manually for every article would be very time-consuming, and as such this tool was created. I added additional features, as well as a GUI with options to make it usable for people outside our research group.

Possible alternatives to this tool are:
- [Mammoth](https://github.com/mwilliamson/python-mammoth), which this project uses and is reliant on.
- [Pandoc](https://pandoc.org/).
- [docx2python](https://github.com/ShayHill/docx2python).
- [PubCSS](https://github.com/thomaspark/pubcss/).
- [Paper Now](https://github.com/PeerJ/paper-now).


## How to use this tool <a name="how-to-use"></a>

The easiest way to use this tool is by downloading the project folder that was created with auto-py-to-exe from the [releases tab](https://github.com/Fulminis-ictus/SciDocx2Web/releases). Just open the EXE file. Note that it's unfortunately common for EXE files created from python files to be marked as viruses. I'm attempting to get it whitelisted but might not have contacted the company who owns your virus program yet.

Alternatively, you can clone this repository or download the files and run SciDocx2WebUI.py. You might need to install required modules such as Mammoth, lxml etc. You'll also need to add [darwyl's page break amendment](https://github.com/dwasyl/python-mammoth/commit/38777ee623b60e6b8b313e1e63f12dafd82b63a4) to Mammoth's body_xml.py if you want to make use of the page numbering feature.

You can find the code's documentation over [here](https://fulminis-ictus.github.io/SciDocx2Web/).

### GUI

Once you run the EXE file, a GUI will open where you can choose different options that change the output:
- **"Only export the body?"** Will only export a body element. This means that no HTML declaration and no head will be inserted into the HTML file. Note that you won't be able to check the CSS, JavaScript and page title options because of that. This option could be useful if you'd like to copy paste the code into a website builder like Wordpress.
- **"Add suggested css?"** Inserts CSS into the head of the document. [See below](https://github.com/Fulminis-ictus/SciDocx2Web#css) for further info.
- **"Add javascript to highlight navigation while scrolling?"** Adds JavaScript to the head. Said JavaScript highlights the heading in the navigation while you're scrolling through the respective chapter.
- **"Page title:"** Inserts the name of the page into the head. Leave this field blank to not insert a page title.
- **"Add IDs to headings"** Adds IDs to all headings that can then be referenced by the navigation. Note that you will need to have marked text as headings (see "Format template detection" below) for SciDocx2Web to be able to detect those headings and add IDs to them.
- **"Create navigation?"** Creates a navigation that is inserted at the top of the article (displayed left of the text and sticky, if the CSS option has been checked). You can choose between displaying the navigation items as paragraphs or as buttons.
- **"Add tooltips to footnotes?"** Encloses footnotes with additional tooltip code. Note that the tooltips will only be displayed properly if the right CSS is included in the HTML file! [See below](https://github.com/Fulminis-ictus/SciDocx2Web#css).
- **"Abbreviate tooltips after how many symbols?"** If footnotes are too long, then the tooltips could potentially reach beyond the page, which could make part of them unreadable. To prevent this, one might want to abbreviate the footnote tooltips. The footnotes at the bottom of the page are unaffected by this. Note that abbreviating footnote tooltips also removes any markup from said footnote tooltips. This is necessary to prevent closing tags from being removed by the abbreviation, which would have a negative effect on the rest of the HTML code.
- **"Number the paragraphs?"** If this option is checked, then all paragraphs, except those that are media captions, bibliographies or which have the "ignorePNum" class (see format templates), are numbered. The numbers are enclosed with square brackets [N] and inserted at the beginning of each paragraph.
- **"Insert page numbers?"** This option inserts page numbers into the text where a new page begins. The numbers are enclosed in curly brackets {N}, are written as subscript, and receive a light gray background if the "Add suggested css?" option is checked. Note that the page numbers are inserted where said page starts, not where it ends, similar to articles where the page number is in the header as opposed to the footer.
- **"Which docx page should be counted as the first page?"** If your DOCX file starts counting pages at a later page, because your first page features nothing but the abstract, for example, then you might want to set this number to that page. This feature exists to make it possible to cross-reference the HTML version and the print/digital PDF version of the article.
- **"Format template detection"**:
    - In Word and LibreOffice you can choose which format templates to apply to marked text. You can also create new format templates and name them however you want. You can create one called "HeadingLevel1", for example, and mark all text that you want displayed as h1 elements with this format template. In SciDocx2Web's GUI you'd then type in that format template name into the top "Detect headings by which format template name?" input field. The text that you marked in Word or LibreOffice will then be marked as a first level heading in the HTML code.
    - You can use the image, video and audio format templates to embed media in your HTML file. Make sure that whatever you've marked with said template is not a hyperlink but just plain text. You might have to make sure that the video link is an actual embed link. https://youtu.be/gpdYKamOjUo might not display properly while https://www.youtube.com/embed/gpdYKamOjUo will. You can also link local files by using a path to said file instead of a link (for example: path-to-video/videoname.mp4).
    - You can input your preferred dimensions for image and video embeds as "X,Y" (without quotation marks). If the input values are faulty (for example if you've input three numbers seperated by commas or if you've input letters) or if the input field is empty, then no "width" and "height" attributes will be inserted and the images and videos will retain their original dimensions.
    - The "Media" and "Table Caption" entries are separate because table captions should, due to semantic reasons, be inside of table elements.
    - If the option to insert the predefined CSS is checked, then bibliographies will be displayed with a hanging indent to improve readability.
    - The field "Detect paragraphs that should not be numbered..." is only active if the below "Number the paragraphs?" option is checked. Paragraphs marked with the respective format template will be ignored when numbering the paragraphs. This might be useful if you have an abstract at the top of the document, for example, that you don't want to number.
    - Text marked with the "code" format template will receive escaped HTML entities to make sure that it isn't interpreted as actual HTML code.
    - You can create additional style map rules in the big entry field. Consult [this guide](https://github.com/mwilliamson/python-mammoth#writing-style-maps) for writing style maps.
    - Any field that is left empty will be skipped.
- **"Save options" and "Reset options"** You can save your preferred options, so you don't have to input them every time you open the program. The saved options are written to SciDocx2Web.ini. You can use the "Reset options" button if you'd like to reset them back to their original settings.

### CSS

The CSS inserted into the document is more about function than visual appeal. If you'd like to add your own CSS or edit or remove the inserted CSS, then simply open the file in a code editor of your choice. CSS code that is important for some features to work correctly, is:
- Footnote tooltips (based on [this code](https://www.w3schools.com/css/css_tooltip.asp)):
    - .tooltippop
    - .tooltippop:hover
    - [role="tooltip"]
- Grid (displaying the main text and the navigation side by side):
    - .gridContainer
    - .navGrid
    - .mainGrid
- Highlighting navigation items when you scroll through the respective chapter:
    - nav p a.highlightnav


## Known problems or missing features - Found a bug? <a name="known-problems"></a>

Make sure to consult the [FAQ](https://github.com/Fulminis-ictus/SciDocx2Web#faq), in case your question is answered there!

If you find any bugs or have any feature requests, then feel free to open an [issue](https://github.com/Fulminis-ictus/SciDocx2Web/issues). Feel free to also open an issue if you have any suggestions or if you have code that you'd like to contribute. Make sure to be thorough when describing your issue. You can upload the file you were trying to convert as well as your SciDocx2Web.ini file, so your problem can be replicated and possible solutions tested.

Note that this tool has only been tested in Word and LibreOffice. If you're using a different text editor, then this tool might not work.

Currently know issues or missing features are:
- Option to automatically create sections based on chapters or pages is missing.
- Abbreviated tooltips don't possess markup (meaning text won't be displayed in cursive, links lose their hyperlinks etc.).
- Page numbers for prefaces aren't supported (roman numerals). It might be worth looking into extracting page numbers from headers and footers instead of looking for page breaks but Mammoth doesn't support that feature. It'd have to be implemented into Mammoth's extraction code.
- Breaks (not paragraphs) are ignored, meaning no "br" tags are inserted anywhere and may have to be added manually.


## FAQ

- **Why aren't my format templates being detected?** Double check that the name entered in the "Format template detection" section of the GUI is written exactly the same as the format template name in your text editor.
- **Why are there empty numbered paragraphs in my output document?** Make sure there are no "seemingly empty" paragraphs in your DOCX document that contain an easy to overlook space. Empty paragraphs are usually ignored, but if you accidentally pressed the space bar while editing that paragraph, then it's counted as a paragraph.
- **Why are no tooltips being displayed above footnotes when I hover over them?** Make sure the "Only export the body?" option is unchecked and the "Add suggested css?" option is checked. The tooltip CSS is necessary for the tooltips to work (see [here](https://github.com/Fulminis-ictus/SciDocx2Web#css)).
- **How can I fix the problem that footnote tooltips are exiting the page's boundaries?** Insert a number into the "Abbreviate tooltips after how many symbols?" input field in the options. Note that this also removes any markup inside the tooltips (meaning text won't be cursive anymore, links won't be hyperlinks etc.).
- **"Why is my video embed not working?"** Make sure you're using the embed link and not the normal link. YouTube's embed link looks something like this: https://www.youtube.com/embed/VIDEOID.
- **Why is the program being flagged as a virus?** It's unfortunately [common](https://medium.com/@markhank/how-to-stop-your-python-programs-being-seen-as-malware-bfd7eb407a7) for EXE files created from python files to be false positives. I'm attempting to get it whitelisted, but might not have contacted the company who owns your virus program yet.
- **Why are you using tooltip spans instead of title attributes?** Title attributes unfortunately aren't very accessible. [(1)](https://www.24a11y.com/2017/the-trials-and-tribulations-of-the-title-attribute/), [(2)](https://sarahmhigley.com/writing/tooltips-in-wcag-21/)


## Acknowledgements

- This tool uses and is heavily reliant on [Mammoth for Python](https://github.com/mwilliamson/python-mammoth), created by Michael Williamson. It uses Mammoth's initial output, as well as its style map function as a basis for further conversion steps.
- The [page break extraction amendment](https://github.com/dwasyl/python-mammoth/commit/38777ee623b60e6b8b313e1e63f12dafd82b63a4) by dwasyl was used as a basis to implement page break markers.
- Thanks to everyone who provided feedback during the development of this project.
