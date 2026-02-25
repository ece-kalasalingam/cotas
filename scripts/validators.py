#This is where the Workbook-Specific rules live. 
# #This is easily extendable for different workbook types.
def validate_course_setup_logic(data_store):
    errors = []
    # Logic: Total Weightage Check
    if "Assessment_Config" in data_store:
        rows = data_store["Assessment_Config"]
        # Column 1 is Weight %
        total = sum(float(r[1]) for r in rows[1:] if r[1] is not None)
        if total != 100:
            errors.append(f"Total weightage is {total}%, must be 100%")
    return errors