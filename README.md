#                          Summary of functionality
#1. Document Type Handling:

    # The tool processes various document types: Word documents (.docx), Excel spreadsheets (.xlsx), and PDFs (.pdf).
    # It can handle both text-based and image-based PDFs, using OCR to extract text from scanned documents when necessary.

#2. Keyword Analysis:

    # The tool analyzes documents for specific legal-related keywords (e.g., "contract," "agreement," "payment").
    # It uses SpaCy for Named Entity Recognition (NER) to detect named entities such as organizations, people, products, and events.
    # Synonym Expansion: With NLTK’s WordNet, it expands keywords by including synonyms, making keyword search more comprehensive.

#3. Date and Monetary Amount Extraction:

    # Dates: Uses dateparser to extract dates, accommodating various date formats across documents.
    # Monetary Amounts: Uses regular expressions to identify and extract currency values, capturing amounts with symbols like $, £, and €.

#4. Text Summarization:

    # Transformers Pipeline: Generates a concise summary of each document, providing a quick overview of the contents.
    # The Hugging Face transformers library is used to produce summaries based on extracted keywords.

#5. Visualization with Word Cloud:

    # For each document, the tool creates a word cloud to visualize the frequency of keywords, giving a quick snapshot of the document’s focus.

#6. Report Generation:

    # Comprehensive Report: For each processed document, the tool generates a report that includes:
        # The frequency of keywords.
        # Extracted dates and monetary amounts.
        # A summary of the document’s key content.
    # The report is saved as a Word document in the selected folder, making it easy to review and share insights.

#7. User-Friendly GUI:

    # A simple GUI built with Tkinter allows users to select a folder, initiate analysis, and receive a notification when the report is complete.
   
#8. Dynamic Summarization:
      #  Introduced a new dynamic_summarization function that adjusts max_length and min_length for the summarizer based on the input text length.
      #  This avoids warnings and ensures efficient summarization for different document lengths.

#9. Efficient Report Generation:
      #  Updated the report generation to use the dynamic_summarization function for summaries, making the summarization process responsive to shorter or longer texts.

 # These changes make the summarization more adaptive, reduce unnecessary warnings, and improve overall efficiency in handling various document types and lengths.
 # Use Cases:

    # 1. Legal Document Analysis: The tool is ideal for processing contracts, agreements, invoices, and other legal documents, extracting relevant information like dates, monetary values, and key legal terms.
    # 2. Compliance and Auditing: Useful for identifying important compliance dates and payment terms within financial and regulatory documents.
    # 3. Summarization and Quick Insights: Helps users quickly review large volumes of documents by generating summaries and visual keyword maps.

 # This comprehensive tool provides both detailed insights and a streamlined overview of each document, making it an efficient solution for legal, financial, and compliance-related text analysis tasks.

#                                End Of Summary 
