# KANO_Model

Code for analyzing KANO Model.

**Inputs:** Responses.xlsx (Contains the survey responses)

**Outputs:** kano_results.xlsx (results of KANO model), customer_satisfaction_diagram.png (Visualization of Customer Satisfaction)

**Code run command:**

**Input params:** path (folder path), start_cols, end_cols (start and end of functional features), mapping (if features to be renamed)

python som727_project.py (If features naming as it is in the responses, and new features)

python som727_project.py --mapping (If features names are to be mapped)

python som727_project.py --start_cols 1 --end_cols 27 (If starting and ending of functional columns are 1 and 27 respectively.)
