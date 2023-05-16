import os
import win32com.client as win32

def convert_to_pdf(input_file, output_file):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    doc.SaveAs(output_file, FileFormat=17)
    doc.Close()
    word.Quit()

def convert_to_word(input_file, output_file):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    doc.SaveAs(output_file, FileFormat=16)
    doc.Close()
    word.Quit()

def convert_to_txt(input_file, output_file):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    doc.SaveAs(output_file, FileFormat=2)
    doc.Close()
    word.Quit()

# Example usage
input_file = 'input.docx'  # Replace with the path to the input file
output_file_pdf = 'output.pdf'
output_file_word = 'output.docx'
output_file_txt = 'output.txt'

# Convert to PDF
convert_to_pdf(input_file, output_file_pdf)
print(f"File converted to PDF: {output_file_pdf}")

# Convert to Word
convert_to_word(input_file, output_file_word)
print(f"File converted to Word: {output_file_word}")

# Convert to TXT
convert_to_txt(input_file, output_file_txt)
print(f"File converted to TXT: {output_file_txt}")
