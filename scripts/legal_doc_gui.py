import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from docxtpl import DocxTemplate
import os
import sys
from datetime import datetime


def number_to_words(num):
    """Convert number to words (1-20+)"""
    words = {
        '1': 'one', '2': 'two', '3': 'three', '4': 'four', '5': 'five',
        '6': 'six', '7': 'seven', '8': 'eight', '9': 'nine', '10': 'ten',
        '11': 'eleven', '12': 'twelve', '13': 'thirteen', '14': 'fourteen', '15': 'fifteen',
        '16': 'sixteen', '17': 'seventeen', '18': 'eighteen', '19': 'nineteen', '20': 'twenty'
    }
    return words.get(str(num), str(num))


def generate_petition(template_path, output_path, petition_data):
    """Generate petition document from template and data"""
    # Load the template
    doc = DocxTemplate(template_path)

    # Process conditional fields based on boolean flags
    is_petitioner_company = petition_data.get("is_petitioner_company", False)
    is_multiple_dwelling = petition_data.get("is_multiple_dwelling", False)
    is_under_rent_stabilization = petition_data.get("is_under_rent_stabilization", False)

    # Set conditional values for company
    if is_petitioner_company:
        petitioner_company_not = " "  # Empty space if IS a company
        representative_name = petition_data.get("representative_name", "")
        representative_title = petition_data.get("representative_title", "")
        # Complete signature line for company including "By:" and comma
        company_signature_line = "By:______________________________________,"
    else:
        petitioner_company_not = "not"  # "not" if NOT a company
        representative_name = ""  # Blank instead of N/A
        representative_title = ""  # Blank instead of N/A
        company_signature_line = ""  # Completely empty - no "By:", no underline, no comma

    # Set conditional values for multiple dwelling checkboxes
    if is_multiple_dwelling:
        no_x = " "  # Blank space for "not multiple dwelling" checkbox
        x_is = "X"  # X for "is multiple dwelling" checkbox
        dwelling_reg_no = petition_data.get("dwelling_registration_no", "")

        # Agent name logic: if company AND multiple dwelling, use representative name
        if is_petitioner_company:
            agent_name = representative_name
        else:
            agent_name = petition_data.get("agent_name", "")

        # Build agent address from respondent address (street only, no apt/room)
        respondent_addr_line1 = petition_data["respondent_address_line1"]
        respondent_addr_line2 = petition_data["respondent_address_line2"]
        # Extract street address (first part before comma)
        street_only = respondent_addr_line1.split(",")[0].strip()
        agent_address = f"{street_only}, {respondent_addr_line2}"
    else:
        no_x = "X"  # X for "not multiple dwelling" checkbox
        x_is = " "  # Blank space for "is multiple dwelling" checkbox
        dwelling_reg_no = "N/A"
        agent_name = "N/A"
        agent_address = "N/A"

    rent_stabilization_not = " " if is_under_rent_stabilization else "not"

    # Convert number of family to words
    number_of_family = petition_data.get("number_of_family", "1")
    number_of_family_words = number_to_words(number_of_family)

    # Define the context (data to fill)
    context = {
        "PETITIONER_NAME": petition_data["petitioner_name"],
        "PETITIONER_ADDRESS_LINE1": petition_data["petitioner_address_line1"],
        "PETITIONER_ADDRESS_LINE2": petition_data["petitioner_address_line2"],
        "RESPONDENT_NAME": petition_data["respondent_name"],
        "RESPONDENT_ADDRESS_LINE1": petition_data["respondent_address_line1"],
        "RESPONDENT_ADDRESS_LINE2": petition_data["respondent_address_line2"],
        "DATED_DATE": petition_data["dated_date"],
        "TERMINATED_DATE": petition_data["terminated_date"],
        # Conditional fields
        "NO_X": no_x,
        "X_IS": x_is,
        "RENT_STABILIZATION_NOT": rent_stabilization_not,
        "PETITIONER_IS_COMPANY_NOT": petitioner_company_not,
        "REPRESENTATIVE_NAME": representative_name,
        "rep_Title": representative_title,
        "COMPANY_SIGNATURE_LINE": company_signature_line,
        "Nutl_No": dwelling_reg_no,
        "AGENT_NAME": agent_name,
        "AGENT_ADDRESS": agent_address,
        # Notice information
        "NOTICE_DAYS": petition_data["notice_days"],
        "NOTICE_TYPE": petition_data["notice_type"],
        # Family information
        "NUMBER_OF_FAMILY": number_of_family,
        "NUMBER_OF_FAMILY_WORDS": number_of_family_words
    }

    # Render the template
    doc.render(context)

    # Save the result
    doc.save(output_path)
    return output_path


class LegalDocApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Legal Document Generator v2.0")
        self.root.geometry("700x850")
        self.root.resizable(False, False)

        # Create scrollable frame
        self.create_scrollable_frame()
        self.create_form()

    def create_scrollable_frame(self):
        """Create canvas with scrollbar"""
        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Enable mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        # Bind mouse wheel to canvas and all widgets
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Store canvas reference for cleanup
        self.canvas = canvas

    def create_form(self):
        """Create all form fields"""
        row = 0

        # Title
        ttk.Label(self.scrollable_frame, text="Legal Document Generator",
                 font=("Arial", 16, "bold")).grid(row=row, column=0, columnspan=2, pady=10)
        row += 1

        ttk.Label(self.scrollable_frame, text="HO NPP Template",
                 font=("Arial", 10)).grid(row=row, column=0, columnspan=2, pady=5)
        row += 1

        ttk.Separator(self.scrollable_frame, orient="horizontal").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=10
        )
        row += 1

        # PETITIONER SECTION
        ttk.Label(self.scrollable_frame, text="PETITIONER INFORMATION",
                 font=("Arial", 11, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="Name: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.petitioner_name = ttk.Entry(self.scrollable_frame, width=50)
        self.petitioner_name.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="Address Line 1: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.petitioner_addr1 = ttk.Entry(self.scrollable_frame, width=50)
        self.petitioner_addr1.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="Address Line 2: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.petitioner_addr2 = ttk.Entry(self.scrollable_frame, width=50)
        self.petitioner_addr2.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        # Company checkbox
        self.is_petitioner_company = tk.BooleanVar()
        ttk.Checkbutton(self.scrollable_frame, text="Petitioner is a Company",
                       variable=self.is_petitioner_company,
                       command=self.toggle_company_fields).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=20, pady=5
        )
        row += 1

        # Representative fields
        ttk.Label(self.scrollable_frame, text="Representative Name:").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.representative_name = ttk.Entry(self.scrollable_frame, width=50, state="disabled")
        self.representative_name.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="Representative Title:").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.representative_title = ttk.Entry(self.scrollable_frame, width=50, state="disabled")
        self.representative_title.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        ttk.Separator(self.scrollable_frame, orient="horizontal").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=10
        )
        row += 1

        # RESPONDENT SECTION
        ttk.Label(self.scrollable_frame, text="RESPONDENT INFORMATION",
                 font=("Arial", 11, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="Name: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.respondent_name = ttk.Entry(self.scrollable_frame, width=50)
        self.respondent_name.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="Address Line 1: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.respondent_addr1 = ttk.Entry(self.scrollable_frame, width=50)
        self.respondent_addr1.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="Address Line 2: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.respondent_addr2 = ttk.Entry(self.scrollable_frame, width=50)
        self.respondent_addr2.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        ttk.Separator(self.scrollable_frame, orient="horizontal").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=10
        )
        row += 1

        # DATES SECTION
        ttk.Label(self.scrollable_frame, text="DATES",
                 font=("Arial", 11, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="Document Date: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.dated_date = DateEntry(self.scrollable_frame, width=47,
                                    background='darkblue', foreground='white',
                                    borderwidth=2, date_pattern='MM/dd/yyyy')
        self.dated_date.grid(row=row, column=1, padx=20, pady=5, sticky="w")
        row += 1

        ttk.Label(self.scrollable_frame, text="Terminated Date: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.terminated_date = DateEntry(self.scrollable_frame, width=47,
                                        background='darkblue', foreground='white',
                                        borderwidth=2, date_pattern='MM/dd/yyyy')
        self.terminated_date.grid(row=row, column=1, padx=20, pady=5, sticky="w")
        row += 1

        ttk.Separator(self.scrollable_frame, orient="horizontal").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=10
        )
        row += 1

        # PROPERTY SECTION
        ttk.Label(self.scrollable_frame, text="PROPERTY INFORMATION",
                 font=("Arial", 11, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=5)
        row += 1

        self.is_multiple_dwelling = tk.BooleanVar()
        ttk.Checkbutton(self.scrollable_frame, text="Is Multiple Dwelling",
                       variable=self.is_multiple_dwelling,
                       command=self.toggle_dwelling_fields).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=20, pady=5
        )
        row += 1

        self.is_rent_stabilization = tk.BooleanVar()
        ttk.Checkbutton(self.scrollable_frame, text="Under Rent Stabilization",
                       variable=self.is_rent_stabilization).grid(
            row=row, column=0, columnspan=2, sticky="w", padx=20, pady=5
        )
        row += 1

        ttk.Label(self.scrollable_frame, text="Dwelling Reg. No.:").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.dwelling_reg_no = ttk.Entry(self.scrollable_frame, width=50, state="disabled")
        self.dwelling_reg_no.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="Agent Name:").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.agent_name = ttk.Entry(self.scrollable_frame, width=50, state="disabled")
        self.agent_name.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="(Agent address auto-generated)",
                 font=("Arial", 8, "italic"), foreground="gray").grid(
            row=row, column=0, columnspan=2, sticky="w", padx=40, pady=2
        )
        row += 1

        ttk.Separator(self.scrollable_frame, orient="horizontal").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=10
        )
        row += 1

        # NOTICE INFORMATION SECTION
        ttk.Label(self.scrollable_frame, text="NOTICE INFORMATION",
                 font=("Arial", 11, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="Notice Days: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.notice_days = ttk.Combobox(self.scrollable_frame, width=47, state="readonly")
        self.notice_days['values'] = ('30', '60', '90')
        self.notice_days.current(0)  # Default to 30
        self.notice_days.grid(row=row, column=1, padx=20, pady=5, sticky="w")
        row += 1

        ttk.Label(self.scrollable_frame, text="Notice Type: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.notice_type = ttk.Combobox(self.scrollable_frame, width=47, state="readonly")
        self.notice_type['values'] = ('oral', 'written')
        self.notice_type.current(0)  # Default to oral
        self.notice_type.grid(row=row, column=1, padx=20, pady=5, sticky="w")
        row += 1

        ttk.Label(self.scrollable_frame, text="Number of Family: *").grid(row=row, column=0, sticky="w", padx=20, pady=5)
        self.number_of_family = ttk.Entry(self.scrollable_frame, width=50)
        self.number_of_family.insert(0, "1")  # Default to 1
        self.number_of_family.grid(row=row, column=1, padx=20, pady=5)
        row += 1

        ttk.Label(self.scrollable_frame, text="(Will generate both numeric and word forms)",
                 font=("Arial", 8, "italic"), foreground="gray").grid(
            row=row, column=0, columnspan=2, sticky="w", padx=40, pady=2
        )
        row += 1

        ttk.Separator(self.scrollable_frame, orient="horizontal").grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=10
        )
        row += 1

        # BUTTONS
        button_frame = ttk.Frame(self.scrollable_frame)
        button_frame.grid(row=row, column=0, columnspan=2, pady=20)

        ttk.Button(button_frame, text="Generate Document",
                  command=self.generate_document).grid(row=0, column=0, padx=10)

        ttk.Button(button_frame, text="Clear Form",
                  command=self.clear_form).grid(row=0, column=1, padx=10)
        row += 1

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(self.scrollable_frame, textvariable=self.status_var,
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=row, column=0, columnspan=2, sticky="ew", padx=20, pady=10)

    def toggle_company_fields(self):
        """Enable/disable company representative fields"""
        if self.is_petitioner_company.get():
            self.representative_name.config(state="normal")
            self.representative_title.config(state="normal")
        else:
            self.representative_name.config(state="disabled")
            self.representative_title.config(state="disabled")
            self.representative_name.delete(0, tk.END)
            self.representative_title.delete(0, tk.END)
        self.update_agent_name_field()

    def toggle_dwelling_fields(self):
        """Enable/disable dwelling fields"""
        if self.is_multiple_dwelling.get():
            self.dwelling_reg_no.config(state="normal")
            self.update_agent_name_field()
        else:
            self.dwelling_reg_no.config(state="disabled")
            self.agent_name.config(state="disabled")
            self.dwelling_reg_no.delete(0, tk.END)
            self.agent_name.delete(0, tk.END)

    def update_agent_name_field(self):
        """Update agent name field based on company and dwelling status"""
        is_company = self.is_petitioner_company.get()
        is_dwelling = self.is_multiple_dwelling.get()

        if is_dwelling:
            if is_company:
                # Company + Multiple Dwelling: auto-filled
                self.agent_name.config(state="disabled")
                self.agent_name.delete(0, tk.END)
            else:
                # Not Company + Multiple Dwelling: manual entry
                self.agent_name.config(state="normal")
        else:
            # Not Multiple Dwelling: disabled
            self.agent_name.config(state="disabled")
            self.agent_name.delete(0, tk.END)

    def validate_form(self):
        """Validate all required fields"""
        errors = []

        if not self.petitioner_name.get().strip():
            errors.append("Petitioner name is required")
        if not self.petitioner_addr1.get().strip():
            errors.append("Petitioner address line 1 is required")
        if not self.petitioner_addr2.get().strip():
            errors.append("Petitioner address line 2 is required")
        if not self.respondent_name.get().strip():
            errors.append("Respondent name is required")
        if not self.respondent_addr1.get().strip():
            errors.append("Respondent address line 1 is required")
        if not self.respondent_addr2.get().strip():
            errors.append("Respondent address line 2 is required")

        # Check company fields
        if self.is_petitioner_company.get():
            if not self.representative_name.get().strip():
                errors.append("Representative name is required when petitioner is a company")
            if not self.representative_title.get().strip():
                errors.append("Representative title is required when petitioner is a company")

        # Check dwelling fields
        if self.is_multiple_dwelling.get():
            if not self.dwelling_reg_no.get().strip():
                errors.append("Dwelling registration number is required for multiple dwelling")
            if not self.is_petitioner_company.get():
                if not self.agent_name.get().strip():
                    errors.append("Agent name is required for multiple dwelling")

        return errors

    def clear_form(self):
        """Clear all form fields"""
        self.petitioner_name.delete(0, tk.END)
        self.petitioner_addr1.delete(0, tk.END)
        self.petitioner_addr2.delete(0, tk.END)
        self.respondent_name.delete(0, tk.END)
        self.respondent_addr1.delete(0, tk.END)
        self.respondent_addr2.delete(0, tk.END)
        self.representative_name.delete(0, tk.END)
        self.representative_title.delete(0, tk.END)
        self.dwelling_reg_no.delete(0, tk.END)
        self.agent_name.delete(0, tk.END)
        self.is_petitioner_company.set(False)
        self.is_multiple_dwelling.set(False)
        self.is_rent_stabilization.set(False)
        self.notice_days.current(0)  # Reset to 30
        self.notice_type.current(0)  # Reset to oral
        self.number_of_family.delete(0, tk.END)  # Clear field
        self.number_of_family.insert(0, "1")  # Reset to 1
        self.toggle_company_fields()
        self.toggle_dwelling_fields()
        self.status_var.set("Form cleared")

    def get_template_path(self):
        """Get template file path"""
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
            template_path = os.path.join(base_path, "HO NPP Template.docx")
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            project_root = os.path.dirname(script_dir)
            template_path = os.path.join(project_root, "templates", "HO NPP Template.docx")
        return template_path

    def generate_document(self):
        """Generate the document"""
        # Validate
        errors = self.validate_form()
        if errors:
            messagebox.showerror("Validation Error", "\n".join(errors))
            return

        try:
            self.status_var.set("Generating document...")
            self.root.update()

            # Get template
            template_path = self.get_template_path()
            if not os.path.exists(template_path):
                messagebox.showerror("Error", f"Template not found:\n{template_path}")
                self.status_var.set("Error: Template not found")
                return

            # Prepare data
            petition_data = {
                "petitioner_name": self.petitioner_name.get().strip(),
                "petitioner_address_line1": self.petitioner_addr1.get().strip(),
                "petitioner_address_line2": self.petitioner_addr2.get().strip(),
                "respondent_name": self.respondent_name.get().strip(),
                "respondent_address_line1": self.respondent_addr1.get().strip(),
                "respondent_address_line2": self.respondent_addr2.get().strip(),
                "dated_date": self.dated_date.get_date().strftime("%B %d, %Y"),
                "terminated_date": self.terminated_date.get_date().strftime("%B %d, %Y"),
                "is_petitioner_company": self.is_petitioner_company.get(),
                "is_multiple_dwelling": self.is_multiple_dwelling.get(),
                "is_under_rent_stabilization": self.is_rent_stabilization.get(),
                "representative_name": self.representative_name.get().strip(),
                "representative_title": self.representative_title.get().strip(),
                "dwelling_registration_no": self.dwelling_reg_no.get().strip(),
                "agent_name": self.agent_name.get().strip(),
                "notice_days": self.notice_days.get(),
                "notice_type": self.notice_type.get(),
                "number_of_family": self.number_of_family.get()
            }

            # Generate output path
            if getattr(sys, 'frozen', False):
                base_path = os.path.dirname(sys.executable)
            else:
                script_dir = os.path.dirname(os.path.abspath(__file__))
                base_path = os.path.dirname(script_dir)

            output_dir = os.path.join(base_path, "output")
            os.makedirs(output_dir, exist_ok=True)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            respondent_name_safe = petition_data["respondent_name"].replace(" ", "_").replace(",", "")
            output_filename = f"output_HO_{respondent_name_safe}_{timestamp}.docx"
            output_path = os.path.join(output_dir, output_filename)

            # Generate document
            generate_petition(template_path, output_path, petition_data)

            # Success
            self.status_var.set("Document generated successfully!")
            result = messagebox.askyesno("Success",
                f"Document generated successfully!\n\nSaved to: {output_path}\n\nWould you like to open it now?")

            if result:
                os.startfile(output_path)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate document:\n{str(e)}")
            self.status_var.set("Error generating document")


def main():
    root = tk.Tk()
    app = LegalDocApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
