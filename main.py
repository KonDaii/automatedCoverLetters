# Import necessary modules
from docxtpl import DocxTemplate  # Importing the DocxTemplate class from docxtpl
from docx2pdf import convert  # Importing the convert function from docx2pdf

# Prompt the user to enter company details
company_name = input("Enter name of the company: ")  # Get the company name from user input
division = input("Enter the division of the company: ")  # Get the division from user input
position_name = input("Enter name of the position: ")  # Get the position name from user input

# Prepare a dictionary with the company details
context = {
    'company_name': company_name,
    'position_name': position_name,
    'division': division
}

# Load the master cover letter template
doc = DocxTemplate("master_cover_letter.docx")

# Render the cover letter template with the provided context
doc.render(context)

# Save the rendered cover letter with a specific file name based on company and position details
doc.save('Cover_letter_'+company_name+'_'+position_name+'.docx')

# Convert the saved cover letter from DOCX format to PDF format using docx2pdf
convert('Cover_letter_'+company_name+'_'+position_name+'.docx')
