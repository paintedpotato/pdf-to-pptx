# pdf-to-pptx
Converts numerically indexed text, to presentation slides. Good for Apocryphal books available in pdf format, for ease of screen projection during sermonettes

# Technical Documentation for PDF to PowerPoint Scripture Converter

## Table of Contents

1. [Introduction](#introduction)
2. [System Requirements](#system-requirements)
3. [Installation](#installation)
4. [Code Structure](#code-structure)
5. [Functions Overview](#functions-overview)
6. [Usage Instructions](#usage-instructions)
7. [Customization](#customization)
8. [Error Handling](#error-handling)
9. [Conclusion](#conclusion)

---

## 1. Introduction <a name="introduction"></a>

This project is designed to convert a PDF document, assumed to contain a book of scripture, into a PowerPoint presentation. Each verse in the scripture corresponds to one slide in the PowerPoint, with the verse's text centered on a navy blue background. The slide's title includes the name of the book, chapter number, and verse number in a smaller, bolded font.

The goal of this converter is to facilitate presentations or teachings based on scripture, offering a visually appealing, automated approach to breaking down a text into slides.

---

## 2. System Requirements <a name="system-requirements"></a>

To use this script, ensure the following prerequisites are met:

- **Python** (Version 3.6 or later)
- The following Python libraries:
  - `PyPDF2` (for PDF text extraction)
  - `python-pptx` (for creating and manipulating PowerPoint files)

Operating system compatibility:
- Windows
- macOS
- Linux

---

## 3. Installation <a name="installation"></a>

1. **Install Python:** Download and install Python from the official website: [Python Downloads](https://www.python.org/downloads/).

2. **Install Required Libraries:**
   Open your terminal or command prompt and run the following commands:
   
   ```bash
   pip install PyPDF2 python-pptx
   ```

---

## 4. Code Structure <a name="code-structure"></a>

The Python script is organized into several key functions to handle PDF text extraction, processing, and PowerPoint creation:

- **`extract_text_from_pdf(pdf_file_path)`**: Extracts the text from the PDF document.
- **`process_scripture_text(text)`**: Processes the extracted text and identifies chapters and verses.
- **`add_verse_slide(prs, book_name, chapter, verse, text)`**: Adds a slide for each verse with a predefined format.
- **`convert_pdf_to_pptx(pdf_file_path, pptx_file_path, book_name)`**: Coordinates the entire process from PDF extraction to PowerPoint creation.

---

## 5. Functions Overview <a name="functions-overview"></a>

### `extract_text_from_pdf(pdf_file_path)`

**Description:**
- Reads the content of the input PDF file using `PyPDF2` and returns the extracted text.

**Parameters:**
- `pdf_file_path` (str): Path to the input PDF file.

**Returns:**
- `text` (str): Raw text extracted from the PDF file.

---

### `process_scripture_text(text)`

**Description:**
- Processes the extracted text to identify chapters and verses. It assumes that the text follows a pattern such as `Chapter:Verse` for each division.

**Parameters:**
- `text` (str): Raw text extracted from the PDF.

**Returns:**
- `scripture_dict` (dict): A dictionary where each key is a tuple (`chapter`, `verse`), and the value is the corresponding verse text.

---

### `add_verse_slide(prs, book_name, chapter, verse, text)`

**Description:**
- Creates a new PowerPoint slide with a navy blue background. The slide contains a title displaying the book, chapter, and verse in smaller, bolded text, and the verse itself is centered in a larger font.

**Parameters:**
- `prs` (Presentation): The PowerPoint presentation object.
- `book_name` (str): The name of the scripture book.
- `chapter` (str): The chapter number.
- `verse` (str): The verse number.
- `text` (str): The verse text.

---

### `convert_pdf_to_pptx(pdf_file_path, pptx_file_path, book_name)`

**Description:**
- The main function that drives the entire conversion process. It extracts the text from the PDF, processes the chapters and verses, and creates a PowerPoint presentation with a slide for each verse.

**Parameters:**
- `pdf_file_path` (str): Path to the input PDF file.
- `pptx_file_path` (str): Path where the PowerPoint presentation will be saved.
- `book_name` (str): The name of the scripture book (used in slide titles).

---

## 6. Usage Instructions <a name="usage-instructions"></a>

1. **Prepare the PDF File:**
   - Ensure that your scripture text is in PDF format and follows a structured chapter:verse format (e.g., `1:1`, `1:2`, etc.).

2. **Run the Script:**
   - Modify the following variables in the script to match your file paths:
     - `pdf_file_path`: Path to the PDF file containing the scripture.
     - `pptx_file_path`: Path where the PowerPoint should be saved.
     - `book_name`: The name of the scripture book (e.g., "Genesis").
   
   Example usage:
   ```python
   pdf_file_path = 'scripture.pdf'
   pptx_file_path = 'scripture_presentation.pptx'
   book_name = 'Genesis'
   
   convert_pdf_to_pptx(pdf_file_path, pptx_file_path, book_name)
   ```

3. **Check Output:**
   - After running the script, check the directory specified by `pptx_file_path`. The generated PowerPoint file will contain a slide for each verse.

---

## 7. Customization <a name="customization"></a>

You can customize the PowerPoint presentation by adjusting several elements:

1. **Fonts and Text Alignment:**
   - Modify font sizes and styles for both the title and body of the slide in the `add_verse_slide` function. The following lines set the font sizes:
   
     ```python
     title_paragraph.font.size = Pt(24)  # For the title
     p.font.size = Pt(32)  # For the verse text
     ```

2. **Slide Layout:**
   - The `slide_layouts[5]` in the `add_verse_slide` function indicates a blank slide layout. You can switch to a different layout if needed.

3. **Background Color:**
   - To change the background color, modify this line:
     ```python
     slide.background.fill.fore_color.rgb = RGBColor(0, 0, 128)  # Navy blue
     ```
   - Use a different RGB color for a custom background.

4. **Text Color:**
   - Both the title and verse text use white (`RGBColor(255, 255, 255)`). You can change the text color by modifying the following lines:
     ```python
     title_paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
     p.font.color.rgb = RGBColor(255, 255, 255)  # White text
     ```

---

## 8. Error Handling <a name="error-handling"></a>

### Common Issues:

1. **File Not Found:**
   - Ensure that the path to the PDF file is correct. If the PDF file is not found, a `FileNotFoundError` will be raised.
   - Solution: Double-check the `pdf_file_path` and make sure the file exists in that location.

2. **Malformed PDF:**
   - If the PDF is not well-formed or has an unusual structure, the text extraction process may not work as expected.
   - Solution: Ensure the PDF has a plain-text structure, or preprocess the PDF to remove unwanted formatting.

3. **Invalid Text Patterns:**
   - The regular expression used to identify verses (`r"(\d+):(\d+)"`) assumes the scripture is structured in the form of `Chapter:Verse`. If the structure is different, the script may not detect verses correctly.
   - Solution: Modify the regular expression in the `process_scripture_text` function to match your document's structure.

---

## 9. Conclusion <a name="conclusion"></a>

This project provides a useful tool for transforming a structured PDF containing scripture into a PowerPoint presentation. It automates the creation of slides, each containing a chapter and verse, making it easier to present or share religious or scholarly texts. The script is designed for flexibility, allowing customization in terms of design, layout, and font style.

For further customizations, you can modify the relevant parts of the code or enhance the script by adding more features (e.g., support for different languages or complex PDF structures).
