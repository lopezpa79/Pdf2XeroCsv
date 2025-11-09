import customtkinter as ctk
import threading
import os
import json
import csv
import fitz  # PyMuPDF
import requests
from tkinter import filedialog
from datetime import datetime

# --- Configuration ---

# The URL for your LM Studio server
LM_STUDIO_URL = "http://localhost:1234/v1/chat/completions"

# The header for the Xero-compatible CSV file.
# Based on Xero's "Sales Invoice" template. Fields with * are mandatory.
XERO_CSV_HEADER = [
    "*ContactName",
    "*InvoiceNumber",
    "*InvoiceDate",
    "*DueDate",
    "InventoryItemCode",
    "*Description",
    "*Quantity",
    "*UnitAmount",
    "Discount",
    "*AccountCode",
    "*TaxType",
    "TrackingName1",
    "TrackingOption1",
    "TrackingName2",
    "TrackingOption2",
    "Currency"
]

# This is the "magic". This prompt instructs the local LLM to extract
# data and return *only* JSON.
SYSTEM_PROMPT = """
You are an expert, high-speed data extraction engine. Your job is to read unstructured text from an invoice and extract its details.

You MUST ONLY respond with a single, valid JSON object. Do not add any text before or after the JSON, such as "Here is the JSON..." or "```json".

The JSON object must have the following structure:
{
  "contact_name": "The customer's name",
  "invoice_number": "The invoice number",
  "invoice_date": "The invoice date (format as YYYY-MM-DD)",
  "due_date": "The due date (format as YYYY-MM-DD)",
  "lines": [
    {
      "description": "Description of the line item",
      "quantity": 1.0,
      "unit_price": 100.00
    }
  ]
}

If you cannot find a value for a field, return null for it.
For dates, try your best to format them as YYYY-MM-DD.
For "lines", return an array of all line items you can find.
"""

class InvoiceExtractorApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Invoice to Xero Extractor")
        self.geometry("600x400")
        ctk.set_appearance_mode("System")
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.select_files_button = ctk.CTkButton(
            self.main_frame,
            text="1. Select Invoice PDFs",
            command=self.select_files
        )
        self.select_files_button.grid(row=0, column=0, padx=20, pady=10, sticky="ew")

        self.process_button = ctk.CTkButton(
            self.main_frame,
            text="2. Process and Save CSV",
            command=self.start_processing_thread,
            state="disabled"
        )
        self.process_button.grid(row=1, column=0, padx=20, pady=10, sticky="ew")

        self.status_label = ctk.CTkLabel(self, text="Welcome! Select PDF files to begin.")
        self.status_label.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        
        self.pdf_file_paths = []

    def select_files(self):
        """Opens a file dialog to select one or more PDF files."""
        self.pdf_file_paths = filedialog.askopenfilenames(
            title="Select Invoice PDFs",
            filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*"))
        )
        
        if self.pdf_file_paths:
            self.status_label.configure(text=f"{len(self.pdf_file_paths)} file(s) selected.")
            self.process_button.configure(state="normal")
        else:
            self.status_label.configure(text="No files selected.")
            self.process_button.configure(state="disabled")

    def start_processing_thread(self):
        """Starts the file processing in a separate thread to avoid freezing the UI."""
        self.process_button.configure(state="disabled", text="Processing...")
        self.select_files_button.configure(state="disabled")
        
        # Run the heavy work in a separate thread
        thread = threading.Thread(target=self.process_files)
        thread.start()

    def process_files(self):
        """The core processing logic (runs in a separate thread)."""
        all_csv_rows = []
        total_files = len(self.pdf_file_paths)

        for i, file_path in enumerate(self.pdf_file_paths):
            self.status_label.configure(text=f"Processing file {i+1}/{total_files}: {os.path.basename(file_path)}...")
            try:
                # 1. Extract text from PDF
                pdf_text = self.extract_text_from_pdf(file_path)
                
                # 2. Query the Local LLM
                invoice_json = self.query_llm(pdf_text)
                
                # 3. Flatten JSON to Xero CSV rows
                if invoice_json:
                    csv_rows_for_this_invoice = self.flatten_json_to_xero_rows(invoice_json)
                    all_csv_rows.extend(csv_rows_for_this_invoice)
                
            except requests.exceptions.ConnectionError:
                self.status_label.configure(text="ERROR: Could not connect to LM Studio. Is it running?")
                self.reset_ui()
                return
            except json.JSONDecodeError:
                self.status_label.configure(text=f"ERROR: LLM returned invalid JSON for {os.path.basename(file_path)}")
                # Continue to the next file
            except Exception as e:
                self.status_label.configure(text=f"ERROR: {e}")
                # Continue to the next file

        if not all_csv_rows:
            self.status_label.configure(text="Processing complete, but no invoice data was extracted.")
            self.reset_ui()
            return

        # 4. Save the combined CSV
        try:
            self.save_csv(all_csv_rows)
        except Exception as e:
            self.status_label.configure(text=f"ERROR: Could not save CSV file. {e}")
        
        self.reset_ui()

    def reset_ui(self):
        """Resets the UI buttons to their initial state."""
        self.process_button.configure(state="normal", text="2. Process and Save CSV")
        self.select_files_button.configure(state="normal")
        self.pdf_file_paths = []

    def extract_text_from_pdf(self, file_path: str) -> str:
        """Opens a PDF and extracts all text content."""
        doc = fitz.open(file_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return text

    def query_llm(self, text: str) -> dict:
        """Sends the extracted text to the local LLM and gets JSON back."""
        headers = {"Content-Type": "application/json"}
        payload = {
            "model": "local-model",  # This doesn't matter for LM Studio
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": text}
            ],
            "temperature": 0.0,
            "stream": False
        }
        
        response = requests.post(LM_STUDIO_URL, headers=headers, json=payload)
        response.raise_for_status()  # Will raise an error for bad responses
        
        raw_response = response.json()['choices'][0]['message']['content']
        
        # Clean up common LLM "chattiness" (e.g., ```json ... ```)
        if raw_response.startswith("```json"):
            raw_response = raw_response[7:-3].strip()
        
        return json.loads(raw_response)

    def flatten_json_to_xero_rows(self, invoice_data: dict) -> list:
        """Converts the structured JSON from the LLM into flat rows for the Xero CSV."""
        rows = []
        
        # Get invoice-level data
        contact = invoice_data.get("contact_name")
        inv_num = invoice_data.get("invoice_number")
        inv_date = invoice_data.get("invoice_date")
        due_date = invoice_data.get("due_date")
        
        if not invoice_data.get("lines"):
            return []

        for line in invoice_data["lines"]:
            # Create a dictionary for one CSV row, starting with defaults
            row_dict = {key: "" for key in XERO_CSV_HEADER}
            
            # --- Fill in the data ---
            
            # Invoice-level data
            row_dict["*ContactName"] = contact
            row_dict["*InvoiceNumber"] = inv_num
            row_dict["*InvoiceDate"] = inv_date
            row_dict["*DueDate"] = due_date
            
            # Line-item data
            row_dict["*Description"] = line.get("description")
            row_dict["*Quantity"] = line.get("quantity")
            row_dict["*UnitAmount"] = line.get("unit_price")
            
            # --- Add Xero-specific defaults ---
            # These are required by Xero. You should change these
            # to match your own Chart of Accounts.
            row_dict["*AccountCode"] = "200"  # "200" is typically "Sales"
            row_dict["*TaxType"] = "GST on Income" # Or "Tax Free", etc.
            
            # Add the row in the correct header order
            rows.append([row_dict[key] for key in XERO_CSV_HEADER])
            
        return rows

    def save_csv(self, all_rows: list):
        """Asks the user where to save the final CSV file."""
        timestamp = datetime.now().strftime("%Y-%m-%d-%hh-%mm-%ss")
        save_path = filedialog.asksaveasfilename(
            title="Save Xero Import File",
            defaultextension=".csv",
            filetypes=(("CSV Files", "*.csv"),),
            initialfile=f"xero_import_{timestamp}.csv"
        )
        
        if not save_path:
            self.status_label.configure(text="Save cancelled.")
            return

        with open(save_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(XERO_CSV_HEADER)
            writer.writerows(all_rows)
            
        self.status_label.configure(text=f"Success! CSV saved to {save_path}")

if __name__ == "__main__":
    app = InvoiceExtractorApp()
    app.mainloop()