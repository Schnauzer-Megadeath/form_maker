import docx

not_finished = True
while(not_finished):

    # Create instance of word document
    doc = docx.Document()

    # Write the first section of the document
    doc_para = doc.add_paragraph('''Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.''')


    # Create the table 
    table = doc.add_table(rows=5, cols=4)
    table.style = "Table Grid"
    row1 = table.rows[0].cells
    row1[0].text = "Bitlocker";
    row1[1].text = "Serial Number";
    row1[2].text = "Employee Name / Job Title";
    row1[3].text = "Description";

    # Get the inputs for the table
    for i in range(4):
        bitlocker = input("Enter the bitlocker: ")
        serial_number = input("Enter the serial number: ")
        employee_name = input("Enter the employee name: ")
        job_title = input("Enter the job_title: ")
        description = input("Enter the description: ")

        name_desc = employee_name + "/" + job_title
        row = table.rows[i+1].cells
        row[0].text = bitlocker;
        row[1].text = serial_number;
        row[2].text = name_desc;
        row[3].text = description;



    # Write the third section of the document
    doc_para = doc.add_paragraph("ultrices neque. Vulputate eu scelerisque felis imperdiet proin. Tincidunt augue interdum velit euismod in pellentesque. Porttitor lacus luctus accumsan tortor posuer")
    doc_para = doc.add_paragraph("Now Sign here __________________")



    # Save
    doc.save("document.docx")

    # Check if the user would like to continue
    not_finished = input("Would you like to continue (y/n): ")
    not_finished = not_finished.lower()
    if (not_finished == "yes" or not_finished == "y"):
        pass
    elif (not_finished == "no" or not_finished == "n"):
        not_finished = False

