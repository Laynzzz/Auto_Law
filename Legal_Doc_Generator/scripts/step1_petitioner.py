from docxtpl import DocxTemplate
import os

def generate_petition(template_path, output_path, petition_data):
    # 1. Load the template
    doc = DocxTemplate(template_path)

    # 2. Process conditional fields based on boolean flags
    is_petitioner_company = petition_data.get("is_petitioner_company", False)
    is_multiple_dwelling = petition_data.get("is_multiple_dwelling", False)
    is_under_rent_stabilization = petition_data.get("is_under_rent_stabilization", False)

    # Set conditional values for company
    if is_petitioner_company:
        petitioner_company_not = " "  # Empty space if IS a company
        representative_name = petition_data.get("representative_name", "")
        representative_title = petition_data.get("representative_title", "")
    else:
        petitioner_company_not = "not"  # "not" if NOT a company
        representative_name = "N/A"
        representative_title = "N/A"

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

    # 3. Define the context (data to fill)
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
        "NO_X": no_x,  # X if NOT multiple dwelling, space if IS multiple dwelling
        "X_IS": x_is,  # X if IS multiple dwelling, space if NOT multiple dwelling
        "RENT_STABILIZATION_NOT": rent_stabilization_not,
        "PETITIONER_IS_COMPANY_NOT": petitioner_company_not,  # "not" if NOT company, space if IS company
        "REPRESENTATIVE_NAME": representative_name,  # Company representative name
        "rep_Title": representative_title,  # Company representative title
        "Nutl_No": dwelling_reg_no,  # Dwelling registration number
        "AGENT_NAME": agent_name,
        "AGENT_ADDRESS": agent_address
    }

    # 4. Render the template
    doc.render(context)

    # 5. Save the result
    doc.save(output_path)
    print(f"Success! Saved to {output_path}")

if __name__ == "__main__":
    # Configuration
    TEMPLATE = os.path.join("..", "templates", "HO NPP Template.docx")
    OUTPUT = os.path.join("..", "output_HO_step1.docx")

    # Sample data for testing
    petition_data = {
        "petitioner_name": "OCEANVIEW PROPERTIES, INC.",
        "petitioner_address_line1": "123 Main Street, Suite 100",
        "petitioner_address_line2": "New York, NY 10001",
        "respondent_name": "JOHN DOE",
        "respondent_address_line1": "456 Oak Avenue, Apt 5B",
        "respondent_address_line2": "Brooklyn, NY 11201",
        "dated_date": "February 7, 2026",
        "terminated_date": "January 15, 2026",
        # Conditional fields
        "is_petitioner_company": True,  # Set to False if petitioner is an individual
        "is_multiple_dwelling": True,  # Set to False to test non-multiple dwelling
        "is_under_rent_stabilization": True,  # Set to False to test not under rent stabilization
        # Company representative info (required if is_petitioner_company is True)
        "representative_name": "JOHN SMITH",
        "representative_title": "Property Manager",
        # Dwelling info (only used if is_multiple_dwelling is True)
        "dwelling_registration_no": "MD-12345",
        # Note: agent_name auto-set to representative_name when company AND multiple dwelling
        # Note: agent_address is auto-generated from respondent address (street only + city/state/zip)
    }

    # Execution
    if os.path.exists(TEMPLATE):
        generate_petition(TEMPLATE, OUTPUT, petition_data)
    else:
        print(f"Error: Could not find {TEMPLATE}")
