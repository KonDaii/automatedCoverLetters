import os
from docxtpl import DocxTemplate  # Importing the DocxTemplate class from docxtpl
from datetime import datetime  # for getting today's date

# Prompt the user to enter company details
company_name = input("Enter name of the company: ")  # Get the company name from user input
print("Company Name:", company_name)  # Debugging line to check input
position_name = input("Enter name of the position: ")  # Get the position name from user input
print("Position Name:", position_name)  # Debugging line to check input

# Prepare a dictionary with the company details
date = datetime.today()
# Function to add the suffix (st, nd, rd, th) to the day
def get_day_suffix(day):
    if 10 <= day % 100 <= 20:  # For 11th, 12th, 13th
        return f"{day}th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
        return f"{day}{suffix}"

# Format the date
formatted_date = date.strftime("%B ") + get_day_suffix(date.day) + date.strftime(", %Y")

context = {
    'company_name': company_name,
    'position_name': position_name,
    'date': formatted_date  # Formatting the date for better readability
}

# Load the master cover letter template
doc = DocxTemplate("master_cover_letter.docx")

# Render the cover letter template with the provided context
doc.render(context)

# Ensure the 'CoverLetters' folder exists, create it if not
folder_path = "CoverLetters"
os.makedirs(folder_path, exist_ok=True)

# Save the rendered cover letter with a specific file name based on company and position details
cover_letter_filename = f'{folder_path}/Cover_letter_{company_name}_{position_name}.docx'
doc.save(cover_letter_filename)

print(f"Cover letter saved as {cover_letter_filename}")
