import tkinter as tk
from tkinter import filedialog, scrolledtext
import docx
import language_tool_python

def check_spelling_grammar(doc_path):
    # Load the Word document
    doc = docx.Document(doc_path)
    
    # Initialize the LanguageTool instance for Australian English
    tool = language_tool_python.LanguageTool('en-AU')
    
    errors = []
    
    # Iterate through each paragraph in the document
    for i, paragraph in enumerate(doc.paragraphs):
        # Check spelling and grammar for the paragraph text
        matches = tool.check(paragraph.text)
        
        # Store errors with suggestions
        for match in matches:
            error_info = {
                'paragraph': i + 1,
                'error_text': match.context,
                'suggestion': match.replacements,
                'message': match.message,
                'offset': match.offset,
                'error_length': match.errorLength
            }
            errors.append(error_info)
    
    return errors

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
    if file_path:
        entry_filepath.delete(0, tk.END)
        entry_filepath.insert(0, file_path)

def analyze_document():
    file_path = entry_filepath.get()
    if file_path:
        errors = check_spelling_grammar(file_path)
        
        # Clear previous results
        text_output.delete(1.0, tk.END)
        
        # Display errors and suggestions
        for error in errors:
            text_output.insert(tk.END, f"Paragraph {error['paragraph']}:\n")
            text_output.insert(tk.END, f"Error: {error['error_text'][error['offset']:error['offset']+error['error_length']]}\n")
            text_output.insert(tk.END, f"Message: {error['message']}\n")
            text_output.insert(tk.END, f"Suggestions: {', '.join(error['suggestion']) if error['suggestion'] else 'None'}\n")
            text_output.insert(tk.END, '\n')

# Set up the GUI
root = tk.Tk()
root.title("Spelling and Grammar Checker")
root.geometry("800x600")

# File path input
frame_filepath = tk.Frame(root)
frame_filepath.pack(pady=10)
label_filepath = tk.Label(frame_filepath, text="Document Path:")
label_filepath.pack(side=tk.LEFT)
entry_filepath = tk.Entry(frame_filepath, width=50)
entry_filepath.pack(side=tk.LEFT, padx=5)
button_browse = tk.Button(frame_filepath, text="Browse", command=browse_file)
button_browse.pack(side=tk.LEFT)

# Analyze button
button_analyze = tk.Button(root, text="Analyze Document", command=analyze_document)
button_analyze.pack(pady=10)

# Output text area
text_output = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=100, height=30)
text_output.pack(padx=10, pady=10)

# Start the GUI loop
root.mainloop()
