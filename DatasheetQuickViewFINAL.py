import os
import tkinter as tk
from tkinter import messagebox
import comtypes.client

# Set the folder containing Word documents
input_folder = r"P:\Christian Williams\1-Datasheet Creation\Script Output\Data Package"

# Get all Word documents in the folder
word_files = [f for f in sorted(os.listdir(input_folder)) if f.endswith(('.docx', '.doc'))]
word_index = 0  # Track the current document index
doc_history = []  # Track previously opened documents

if not word_files:
    print("No Word documents found in the specified folder.")
    exit()

# Initialize Word application
word = comtypes.client.CreateObject("Word.Application")
word.Visible = True  # Make Word visible

current_doc = None  # Track currently opened document

def open_doc(index):
    """Opens a Word document at the given index and sets zoom to 80%."""
    global word_index, current_doc

    if current_doc:
        current_doc.Close(False)  # Close current document without saving

    if 0 <= index < len(word_files):
        doc_path = os.path.join(input_folder, word_files[index])
        current_doc = word.Documents.Open(doc_path)

        # Set zoom level to 80%
        try:
            word.ActiveWindow.View.Zoom.Percentage = 80
        except Exception as e:
            print(f"Error setting zoom: {e}")

        word_index = index  # Update current document index

def approve_and_next():
    """Approves the current document and moves to the next one."""
    global word_index

    if word_index < len(word_files) - 1:
        doc_history.append(word_index)  # Save current index before moving forward
        open_doc(word_index + 1)
    else:
        word.Quit()  # Quit Word when done
        messagebox.showinfo("Review Complete", "All documents have been reviewed.")
        root.destroy()  # Close the GUI

def go_back():
    """Goes back to the previous document if available."""
    global word_index

    if doc_history:
        last_index = doc_history.pop()  # Retrieve the last visited document index
        open_doc(last_index)
    else:
        messagebox.showwarning("No Previous Document", "You're already at the first document.")

# Create the GUI window
root = tk.Tk()
root.title("Document Approval")
root.geometry("300x150")

# Approval button
approve_button = tk.Button(root, text="Approve & Next", font=("Arial", 12), command=approve_and_next)
approve_button.pack(expand=True, pady=10)

# Back button
back_button = tk.Button(root, text="Back", font=("Arial", 12), command=go_back)
back_button.pack(expand=True, pady=10)

# Start with the first document
open_doc(word_index)

# Run the GUI
root.mainloop()
