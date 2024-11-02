# Document-Analysis-v2
designed to analyze files within a selected folder, extract relevant information, summarize the content, and provide visual insights through word clouds

Overview

This code is a Document Analysis Tool designed to analyze files within a selected folder, extract relevant information, summarize the content, and provide visual insights through word clouds. It processes .docx, .xlsx, and .pdf files using Natural Language Processing (NLP), Optical Character Recognition (OCR) for PDFs, and summarization. The output is a Word document containing an analysis report for each file, including keyword counts and summaries.
Functional Breakdown and Use Scenarios
1. GUI Setup and File Selection: load_folder() Function

    Purpose: Provides a user interface that allows the user to select a folder for analysis.
    Functionality: When the "Load Folder" button is pressed, the user selects a folder, and the tool processes all supported files in that folder.
    Use Scenario: A user wants to analyze a batch of reports located in a single directory and get a summary report of each document's keywords, entities, and a condensed summary.

2. NLP Setup for Text Analysis: spacy and pipeline

    Purpose: Initializes the SpaCy language model for Named Entity Recognition (NER) and keyword matching, and sets up a Hugging Face summarization model for text summarization.
    Functionality:
        SpaCy provides tools for text processing and extracting entities like organizations, locations, and people.
        Hugging Face Summarization Pipeline creates concise summaries of text.
    Use Scenario: The models load during initialization and provide the base for all subsequent analysis, making the tool capable of reading, summarizing, and extracting information from large text blocks.

3. Synonym Expansion for Keywords: Keyword Synonym Finder

    Purpose: Creates a set of expanded keywords using WordNet synonyms, allowing the tool to recognize more variations of each keyword.
    Functionality: Expands each keyword in keywords with its synonyms, resulting in a broader range of terms the tool can detect.
    Use Scenario: If a user’s keyword list includes "analysis" and "data," this function expands it to include synonyms, improving the accuracy of keyword detection in varied texts.

4. Analyze Text with NLP: analyze_text_with_nlp(text, keywords)

    Purpose: Analyzes text to count occurrences of specified keywords and identifies Named Entities (like companies, locations, and individuals).
    Functionality:
        Searches for each keyword within the text.
        Uses NER to recognize entities of interest (e.g., names, places, events) that are not in the keyword list.
    Use Scenario: Useful for users who want to count specific terms in each document while also automatically identifying important named entities, enhancing the depth of analysis.

5. OCR for Image-Based PDFs: ocr_image(image_path)

    Purpose: Applies OCR to images or image-based PDFs, converting the text content into searchable text.
    Functionality: Uses Tesseract OCR to recognize and extract text from images.
    Use Scenario: Essential for users working with scanned PDF documents. When text extraction from PDFs fails, OCR is applied to retrieve text, ensuring even non-digital documents are analyzed effectively.

6. Word Cloud Generator: generate_wordcloud(found_keywords)

    Purpose: Creates a word cloud to visually display keyword frequencies.
    Functionality: Generates a graphical representation of keywords, where the size of each word corresponds to its frequency.
    Use Scenario: When analyzing multiple reports, users can quickly grasp the prominence of various keywords in a document through the word cloud, which aids in understanding document themes at a glance.

7. Document Processing Functions:

    Each of these functions processes a specific file type (.docx, .xlsx, or .pdf) by reading its content and extracting relevant text for analysis.

process_word_doc(file_path)

    Purpose: Extracts text from .docx (Word) files.
    Functionality: Reads the text from each paragraph in the document, then sends it to analyze_text_with_nlp() for keyword and entity detection.
    Use Scenario: Ideal for users with a collection of reports or notes in Word format, needing a summary of the text content and keyword analysis.

process_excel(file_path)

    Purpose: Extracts text from .xlsx (Excel) files.
    Functionality: Reads each cell in each sheet of the Excel file and compiles the text for analysis.
    Use Scenario: Useful for analyzing textual data stored in spreadsheets, such as project descriptions, comments, or structured notes, which are then summarized and analyzed for trends.

process_pdf(file_path)

    Purpose: Extracts text from .pdf files and applies OCR if necessary.
    Functionality: Reads text from PDF pages. If text extraction fails (e.g., scanned PDFs), it applies OCR to the images.
    Use Scenario: Beneficial for handling mixed PDF types, such as digitally created and scanned PDFs. Ensures the user can analyze PDFs even when traditional text extraction methods are ineffective.

8. Report Generation: Adding Analysis to the Word Report in load_folder()

    Purpose: Compiles a Word document report summarizing the analysis for each processed file.
    Functionality:
        For each analyzed file, adds a section with keyword counts, entities, and a summary (if keywords are found).
        Generates a word cloud for each document and includes it in the analysis.
    Use Scenario: Allows users to create a consolidated summary of their document folder, with detailed keyword counts, identified entities, and visual summaries. The final Word report is especially valuable for teams needing an aggregated analysis of documents in a single report.

9. Summarization with Hugging Face: Summarizer in load_folder()

    Purpose: Produces a concise summary of each document’s keywords and main ideas.
    Functionality: Uses a Hugging Face model to create a short text summary when keywords are found in the document.
    Use Scenario: Useful when a user needs a quick overview of the main points of each document, making it easier to understand the content without reading through every detail.
