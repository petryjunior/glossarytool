# Import the necessary modules
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.messagebox as messagebox
import textwrap
import os
import re
import chardet
import csv
import time

# Global variables for entry, glossaries, term_list, and result_text
entry = None
glossaries = []  # A list to store the loaded glossaries along with their file names
term_set = set()  # A set to store unique terms
term_list = None
result_text = None
apply_exact_match_filter = None
exact_match_filter = False
apply_whole_word_match = None
whole_word_match = False
selected_glossary_label = None
selected_glossaries_states = []
glossary_checkbox_frame = None
window = None
DEBOUNCE_INTERVAL = 500  # Adjust this value as needed
debounce_after_id = None  # Initialize the debounce timer ID

def throttle(delay):
    def decorator(fn):
        last_called = [0]  # Using a mutable object to store last_called time

        def throttled(*args, **kwargs):
            now = time.time()
            if now - last_called[0] >= delay:
                last_called[0] = now
                return fn(*args, **kwargs)

        return throttled

    return decorator

def detect_csv_delimiter(filename):
    with open(filename, 'rb') as f:
        result = chardet.detect(f.read())
    if 'encoding' in result:
        encoding = result['encoding']
        if encoding:
            sample_size = min(4096, os.path.getsize(filename))
            with open(filename, 'r', encoding=encoding) as f:
                try:
                    dialect = csv.Sniffer().sniff(f.read(sample_size), delimiters=(',', ';', '\t'))
                    return dialect.delimiter
                except csv.Error:
                    pass
    return ','  # Default delimiter if detection fails

# Function to get the file name from the full path
def get_file_name(file_path):
    return os.path.basename(file_path)

# Function to populate the term list when the user searches for a term
def populate_term_list():
    global entry, glossaries, term_list, exact_match_filter, whole_word_match
    term = entry.get().strip().lower()
    term_list.delete(0, tk.END)

    if not term:
        return

    filtered_terms = set()

    for filename, glossary in glossaries:
        if exact_match_filter:
            terms = glossary[glossary[glossary.columns[0]].str.lower() == term][glossary.columns[0]].tolist()
        elif whole_word_match:
            terms = glossary[glossary[glossary.columns[0]].str.contains(fr'\b{re.escape(term)}\b', case=False, na=False, regex=True)][glossary.columns[0]].tolist()
        else:
            terms = glossary[glossary[glossary.columns[0]].str.contains(re.escape(term), case=False, na=False, regex=True)][glossary.columns[0]].tolist()
        filtered_terms.update([t.lower() for t in terms])  # Convert to lowercase and update the set

    for term in sorted(filtered_terms):
        term_list.insert(tk.END, term)

# Function to load the glossary data from the selected Excel file
def load_glossaries():
    global glossaries, term_set, term_list, selected_glossaries_states, filter_frame
    filenames = filedialog.askopenfilenames(
        filetypes=[("Excel and CSV files", "*.xlsx;*.csv")]
    )

    # Check if the user closed the file dialog without choosing a file
    if not filenames:
        return  # Exit the function

    # Declare the message variable here
    message = ""

    # Initialize i before the loop
    i = 0

    if filenames:
        glossaries = []
        term_set = set()
        term_list.delete(0, tk.END)
        
        # Initialize the selected_glossaries_states list with True values for each loaded glossary
        selected_glossaries_states = [tk.BooleanVar(value=True) for _ in range(len(filenames))]
        
        # Clear any existing checkboxes and labels
        for widget in filter_frame.winfo_children():
            widget.destroy()
        
        large_glossaries_detected = 0  # Count the number of large glossaries detected
            
        for idx, filename in enumerate(filenames):
            _, file_extension = os.path.splitext(filename)
            if file_extension.lower() == '.xlsx':
                glossary_df = pd.read_excel(filename, engine='openpyxl')
            elif file_extension.lower() == '.csv':
                # Automatically detect the CSV delimiter
                delimiter = detect_csv_delimiter(filename)
                try:
                    glossary_df = pd.read_csv(filename, delimiter=delimiter, encoding='utf-8')
                except pd.errors.ParserError as e:
                    print(f"Error parsing {filename}: {e}")
                    glossary_df = pd.DataFrame()  # Create an empty DataFrame to avoid errors
            else:
                continue  # Unsupported file type

            # Check the size of the glossary
            if len(glossary_df) > 10000:
                large_glossaries_detected += 1
        
        # Process glossaries
        for idx, filename in enumerate(filenames):
            _, file_extension = os.path.splitext(filename)
            if file_extension.lower() == '.xlsx':
                glossary_df = pd.read_excel(filename, engine='openpyxl')
            elif file_extension.lower() == '.csv':
                try:
                    # Specify UTF-8 encoding and tab as the delimiter for CSV files
                    glossary_df = pd.read_csv(filename, delimiter=delimiter, encoding='utf-8')
                except pd.errors.ParserError as e:
                    print(f"Error parsing {filename}: {e}")
                    glossary_df = pd.DataFrame()  # Create an empty DataFrame to avoid errors
            else:
                continue  # Unsupported file type

            # Check the size of the glossary
            if len(glossary_df) > 10000:
                # Skip populating the term list
                checkbox = ttk.Checkbutton(filter_frame, text=get_file_name(filename), variable=selected_glossaries_states[i], command=update_term_list)
                
            elif selected_glossaries_states[i].get():  # Check if the glossary is selected
                # Populate the term list for smaller glossaries of selected glossaries
                terms = glossary_df[glossary_df.columns[0]].tolist()
                for term in terms:
                    if isinstance(term, str):
                        term_lower = term.lower()
                        term_set.add(term_lower)
                for term in sorted(term_set):
                    term_list.insert(tk.END, term.title())
            else:
                continue  # Skip this glossary if it's not selected

            
            glossaries.append((filename, glossary_df))  # Store the file name along with the DataFrame
            if len(glossary_df) > 10000:
                continue  # Skip populating the term list for very large glossaries
            terms = glossary_df[glossary_df.columns[0]].tolist()
            for term in terms:
                if isinstance(term, str):
                    term_lower = term.lower()  # Convert to lowercase
                    term_set.add(term_lower)    # Store unique terms in lowercase
        
        for term in sorted(term_set):
            term_list.insert(tk.END, term.title())  # Insert terms with title case

        # Update the message if large glossaries were detected
        if large_glossaries_detected > 0:
            message = f"You have loaded one or more glossaries with over 10,000 entries. " \
                      f"For optimized performance, the list of terms will not be populated with content from " \
                      f"these glossaries, but you can still use the search box to look up terms in them."
    
    # Display a message if a large glossary was detected
    if message:
        messagebox.showinfo("Large Glossary Detected", message)
    
    # Create checkboxes for each loaded glossary
    for i, (filename, _) in enumerate(glossaries):
        checkbox = ttk.Checkbutton(filter_frame, text=get_file_name(filename), variable=selected_glossaries_states[i], command=lambda i=i: toggle_glossary(i))
        checkbox.grid(row=i, column=1, sticky="w")

    # Re-add whole word match and exact match filters
    whole_word_match_checkbox = ttk.Checkbutton(filter_frame, text="Whole Word Match", variable=apply_whole_word_match, command=update_term_list)
    whole_word_match_checkbox.grid(row=0, column=0, sticky="w")
    exact_match_filter_checkbox = ttk.Checkbutton(filter_frame, text="Exact Match", variable=apply_exact_match_filter, command=update_term_list)
    exact_match_filter_checkbox.grid(row=1, column=0, sticky="w")

    # Automatically update the term list and results box when glossaries are loaded
    update_term_list()

    entry.focus_force()

def toggle_glossary(index):
    # Store the current filter states
    exact_match_filter_state = apply_exact_match_filter.get()
    whole_word_match_state = apply_whole_word_match.get()

    if selected_glossaries_states[index].get():
        # If the glossary checkbox is checked, add the glossary entries
        add_glossary_entries(index)
    else:
        # If the glossary checkbox is unchecked, remove the glossary entries
        remove_glossary_entries(index)

    # Reapply the stored filter states
    apply_exact_match_filter.set(exact_match_filter_state)
    apply_whole_word_match.set(whole_word_match_state)

    if selected_glossaries_states[index].get():
        _, glossary_df = glossaries[index]
        
        # Check the size of the glossary
        if len(glossary_df) <= 10000:
            terms = glossary_df[glossary_df.columns[0]].tolist()
            for term in terms:
                if isinstance(term, str):
                    term_lower = term.lower()
                    term_set.add(term_lower)
                    term_list.insert(tk.END, term.title())
      

    update_term_list()

def remove_glossary_entries(glossary_index):
    global term_set, term_list
    _, glossary_df = glossaries[glossary_index]
    terms = glossary_df[glossary_df.columns[0]].tolist()
    
    for term in terms:
        if isinstance(term, str):  # Check if the term is a string
            term_lower = term.lower()
            if term_lower in term_set:
                term_set.remove(term_lower)
                try:
                    term_list.delete(term_list.get(0, tk.END).index(term_lower.title()))
                except ValueError:
                    pass

    term_list.delete(0, tk.END)
    
    for term in sorted(term_set):
        term_list.insert(tk.END, term.title())

def add_glossary_entries(glossary_index):
    global term_set, term_list, exact_match_filter, whole_word_match
    _, glossary_df = glossaries[glossary_index]

    for idx, (_, glossary) in enumerate(glossaries):
        if not selected_glossaries_states[idx].get():
            continue  # Skip if the glossary is not selected

        term = entry.get().strip().lower()

        filtered_terms = set()

        if exact_match_filter:
            terms = glossary[glossary[glossary.columns[0]].str.lower() == term][glossary.columns[0]].tolist()
        elif whole_word_match:
            terms = glossary[glossary[glossary.columns[0]].str.contains(fr'\b{re.escape(term)}\b', case=False, na=False, regex=True)][glossary.columns[0]].tolist()
        else:
            terms = glossary[glossary[glossary.columns[0]].str.contains(re.escape(term), case=False, na=False, regex=True)][glossary.columns[0]].tolist()
        
        filtered_terms.update([t.lower() for t in terms])  # Convert to lowercase and update the set

    # Clear the term list
    term_list.delete(0, tk.END)

    for term in sorted(filtered_terms):
        term_list.insert(tk.END, term)

# Function to update the list of terms based on the current search term and filters
@throttle(delay=0.1)  # Adjust the delay as needed, e.g., 0.5 seconds
def update_term_list(event=None):
    global entry, glossaries, term_list, selected_glossaries_states, apply_exact_match_filter, apply_whole_word_match
    term = entry.get().strip().lower()
    term_list.delete(0, tk.END)

    if not term:
        for term in sorted(term_set):
            term_list.insert(tk.END, term)
        return

    filtered_terms = set()

    for i, (filename, glossary) in enumerate(glossaries):
        if not selected_glossaries_states[i].get():
            continue  # Skip if the glossary is not selected

        # Apply the exact match and whole word match filters
        if apply_exact_match_filter.get():
            terms = glossary[glossary[glossary.columns[0]].str.lower() == term][glossary.columns[0]].tolist()
        elif apply_whole_word_match.get():
            terms = glossary[glossary[glossary.columns[0]].str.contains(fr'\b{re.escape(term)}\b', case=False, na=False, regex=True)][glossary.columns[0]].tolist()
        else:
            terms = glossary[glossary[glossary.columns[0]].str.contains(re.escape(term), case=False, na=False, regex=True)][glossary.columns[0]].tolist()
        filtered_terms.update([t.lower() for t in terms])  # Convert to lowercase and update the set

    for term in sorted(filtered_terms):
        term_list.insert(tk.END, term)

# Function to handle key release events and debounce the search
def on_key_release(event, window):
    global debounce_after_id
    if debounce_after_id:
        window.after_cancel(debounce_after_id)
    debounce_after_id = window.after(DEBOUNCE_INTERVAL, update_term_list)

# Function to handle the change in the exact match filter state
def on_exact_match_filter_change():
    global exact_match_filter
    exact_match_filter = apply_exact_match_filter.get()
    update_term_list()  # Update the term list when the filter state changes

# Function to handle the change in the whole word match filter state
def on_whole_word_match_change():
    global whole_word_match
    whole_word_match = apply_whole_word_match.get()
    update_term_list()  # Update the term list when the filter state changes

# Function to wrap cell content within the specified column width
def wrap_cell_content(content, column_size):
    wrapped_content = textwrap.fill(content, width=column_size)
    return wrapped_content

# Function to perform term lookup and update the right-hand side with the selected term's details
def show_term_details(event):
    global term_list, glossaries, result_text, selected_glossaries_states
    selected_index = term_list.curselection()
    if selected_index:
        selected_term = term_list.get(selected_index)
        result_text.config(state=tk.NORMAL)
        result_text.delete(1.0, tk.END)

        # Define a tag for the glossary filenames and configure its color to dark green
        result_text.tag_configure("filename", foreground="dark green")

        previous_glossary = None  # Track the previous glossary to detect when a new glossary starts
        for idx, (filename, glossary) in enumerate(glossaries):
            # Check if the glossary is selected before processing
            if selected_glossaries_states[idx].get():
                result = glossary[glossary[glossary.columns[0]].str.lower() == selected_term.lower()]
                
                if not result.empty:
                    glossary_filename = get_file_name(filename)
                    if previous_glossary is not None and previous_glossary != glossary_filename:
                        result_text.insert(tk.END, "\n")  # Add a line break between glossaries
                    previous_glossary = glossary_filename

                    # Apply the "filename" tag to the glossary filename for color formatting
                    result_text.insert(tk.END, f"From: {glossary_filename}\n", ("bold", "filename"))

                    previous_row = None  # Track the previous row to detect when a new instance starts
                    for _, row in result.iterrows():
                        if previous_row is not None and not row.equals(previous_row):
                            result_text.insert(tk.END, "\n")  # Add a line break between instances of the same term
                        previous_row = row.copy()  # Create a copy of the row to avoid altering the original DataFrame

                        for column in glossary.columns:
                            cell_content = str(row[column]).replace("nan", "N/A")
                            cell_content = cell_content.replace("_x000D_", "\n").replace("\n\n", "\n")  # Replace Excel line break characters
                            formatted_line = f"{column}: "
                            if column == glossary.columns[-1]:
                                result_text.insert(tk.END, formatted_line, "bold")
                                result_text.insert(tk.END, cell_content + "\n")
                            else:
                                result_text.tag_configure("bold", font=("TkDefaultFont", 10, "bold"))
                                result_text.insert(tk.END, formatted_line, "bold")
                                result_text.insert(tk.END, cell_content + "\n")

        result_text.config(state=tk.DISABLED)

# Function to handle the change in the exact match filter state
def on_exact_match_filter_change():
    global exact_match_filter
    exact_match_filter = apply_exact_match_filter.get()
    update_term_list()  # Update the term list when the filter state changes

# Function to handle the change in the whole word match filter state
def on_whole_word_match_change():
    global whole_word_match
    whole_word_match = apply_whole_word_match.get()
    update_term_list()  # Update the term list when the filter state changes

# Main function to create the GUI
def main():
    global entry, term_list, result_text, apply_exact_match_filter, apply_whole_word_match, selected_glossary_label, filter_frame, glossary_checkbox_frame
    window = tk.Tk()
    window.title("Glossary Lookup Tool")

    checkboxes = []  # List to store glossary selection checkboxes

    # Create the frame for the glossary button and search box
    top_frame = ttk.Frame(window)
    top_frame.grid(row=0, column=0, columnspan=2, pady=5, padx=5, sticky="w")

    # Create checkboxes for selecting glossaries
    for idx, (filename, _) in enumerate(glossaries):
        selected_glossaries_states.append(tk.BooleanVar(value=True))  # Initialize all checkboxes as selected
        glossary_checkbox = ttk.Checkbutton(window, text=get_file_name(filename), variable=selected_glossaries_states[idx])
        glossary_checkbox.grid(row=4 + idx, column=0, columnspan=2, pady=2, padx=5, sticky="w")
        checkboxes.append(glossary_checkbox)

    selected_glossaries_states = [tk.BooleanVar(value=False) for _ in glossaries]

    # Create the frame for the glossary button and search box
    top_frame = ttk.Frame(window)
    top_frame.grid(row=0, column=0, columnspan=2, pady=5, padx=5, sticky="w")

    # Create a button to choose the glossary file
    button = ttk.Button(top_frame, text="Choose Glossary File(s)", command=load_glossaries)
    button.pack(side=tk.LEFT, padx=5)

    # Create the label for the search box
    search_label = ttk.Label(top_frame, text="Enter term to look up:")
    search_label.pack(side=tk.LEFT)

    # Create the search box to the right of the label
    entry = tk.Entry(top_frame)
    entry.pack(side=tk.LEFT, padx=5)
    entry.bind("<KeyRelease>", lambda event, window=window: on_key_release(event, window))

    # Create a PanedWindow for the list of terms and results
    paned_window = ttk.PanedWindow(window, orient=tk.HORIZONTAL)
    paned_window.grid(row=1, column=0, columnspan=2, pady=5, padx=5, sticky="nsew")
    window.grid_rowconfigure(1, weight=1)  # Allow the row to expand vertically

    # Create the list of terms frame and add it to the PanedWindow
    term_list_frame = ttk.Frame(paned_window)
    paned_window.add(term_list_frame)

    term_list_scrollbar_y = ttk.Scrollbar(term_list_frame, orient=tk.VERTICAL)
    term_list_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)

    term_list = tk.Listbox(term_list_frame, yscrollcommand=term_list_scrollbar_y.set, font=("TkDefaultFont", 10), width=25)
    term_list.pack(fill=tk.BOTH, expand=True)
    term_list_scrollbar_y.config(command=term_list.yview)

    term_list.bind("<<ListboxSelect>>", show_term_details)  # Bind the event to the term list

    # Create the results frame and add it to the PanedWindow
    result_frame = ttk.Frame(paned_window)
    paned_window.add(result_frame)

    result_text_scrollbar_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL)
    result_text_scrollbar_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL)
    result_text_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    result_text_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

    result_text = tk.Text(result_frame, wrap="word", yscrollcommand=result_text_scrollbar_y.set,
                        xscrollcommand=result_text_scrollbar_x.set, font=("TkDefaultFont", 10))
    result_text.pack(fill=tk.BOTH, expand=True)

    result_text_scrollbar_y.config(command=result_text.yview)
    result_text_scrollbar_x.config(command=result_text.xview)

    # Create the frame for the filter checkboxes
    filter_frame = ttk.Frame(window)
    filter_frame.grid(row=2, column=0, columnspan=2, pady=5, padx=5, sticky="w")

    apply_exact_match_filter = tk.BooleanVar(value=False)  # Variable to store the checkbox state
    exact_match_filter_checkbox = ttk.Checkbutton(filter_frame, text="Exact Match Filter",
                                              variable=apply_exact_match_filter, command=on_exact_match_filter_change)
    exact_match_filter_checkbox.grid(row=0, column=0, sticky="w")  # Updated to use .grid() instead of .pack()

    apply_whole_word_match = tk.BooleanVar(value=False)  # Variable to store the checkbox state
    whole_word_match_checkbox = ttk.Checkbutton(filter_frame, text="Whole Word Match Filter",
                                            variable=apply_whole_word_match, command=on_whole_word_match_change)
    whole_word_match_checkbox.grid(row=1, column=0, sticky="w")  # Updated to use .grid() instead of .pack()

    # Add Export and Clear/Reset buttons to the top_frame
    export_button = ttk.Button(top_frame, text="Export Results", command=export_results)
    export_button.pack(side=tk.LEFT, padx=5)
    clear_button = ttk.Button(top_frame, text="Clear/Reset", command=clear_search)
    clear_button.pack(side=tk.LEFT, padx=5)

    # Center the window on the screen
    window_width = 800
    window_height = 550
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x_coordinate = int((screen_width / 2) - (window_width / 2))
    y_coordinate = int((screen_height / 2) - (window_height / 2))
    window.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    # Automatically open the file dialog to choose glossaries
    load_glossaries()

    # Add status bar at the bottom
    status_bar = ttk.Label(window, text="Loaded glossaries: 0", anchor="w")
    status_bar.grid(row=99, column=0, columnspan=2, sticky="we")

    # Update status bar on relevant events
    term_list.bind("<<ListboxSelect>>", lambda e: [show_term_details(e), update_status_bar()])
    entry.bind("<KeyRelease>", lambda event, window=window: [on_key_release(event, window), update_status_bar()])
    # Also update after loading glossaries
    def after_load():
        update_term_list()
        update_status_bar()
    # Patch load_glossaries to call after_load at the end
    orig_load_glossaries = load_glossaries
    def patched_load_glossaries():
        orig_load_glossaries()
        after_load()
    button.config(command=patched_load_glossaries)
    # Initial status bar update
    update_status_bar()

    # Start the GUI event loop
    window.mainloop()

# --- New helper functions ---
def export_results():
    content = result_text.get(1.0, tk.END).strip()
    if not content:
        messagebox.showinfo("Export Results", "There is no content to export.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
    if file_path:
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(content)
        messagebox.showinfo("Export Results", f"Results exported to {file_path}")

def clear_search():
    entry.delete(0, tk.END)
    apply_exact_match_filter.set(False)
    apply_whole_word_match.set(False)
    update_term_list()
    result_text.config(state=tk.NORMAL)
    result_text.delete(1.0, tk.END)
    result_text.config(state=tk.DISABLED)
    update_status_bar()

def update_status_bar(*args):
    loaded = len(glossaries)
    selected = None
    try:
        idx = term_list.curselection()
        if idx:
            selected = term_list.get(idx)
    except Exception:
        selected = None
    status = f"Loaded glossaries: {loaded}"
    if selected:
        status += f" | Selected term: {selected}"
    status_bar.config(text=status)

# --- End new helper functions ---

if __name__ == "__main__":
    main()
