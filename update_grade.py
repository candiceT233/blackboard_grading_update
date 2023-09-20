import sys
import pandas as pd
import re

def convert_comment_to_html(comment):
    # Split the comment into lines
    lines = comment.split('\n')

    # Initialize an empty HTML string
    html_str = '<p>'

    # Flag to indicate if we should skip the line due to Total score
    skip_line = False
    first_line_contains_score = False

    # Iterate through the lines of the comment
    for line in lines:
        line = line.strip()
            
        # Check if the line starts with 'Q' followed by a digit
        if re.match(r'^Q\d+:', line):
            # Add a line break and the question label
            if not skip_line:
                html_str += f'<br>{line}<br>'
            skip_line = False
            # Check if this line contains a score
            if re.search(r'\(-\d+(\.\d+)?\)', line):
                first_line_contains_score = True
            else:
                first_line_contains_score = False
        elif line == 'General Comment: -':
            # Add 'General Comment: -' with a line break
            if not skip_line:
                html_str += f'{line}<br>'
            skip_line = False
        elif ('Total' in line or 'total' in line):
            # Set the skip_line flag to True to skip subsequent lines
            if first_line_contains_score:
                skip_line = True
            else:
                skip_line = False
                # Remove "Total:" and similar text and convert to lowercase
                line = line.lower().replace("total", "").replace("total:", "").strip()
                if line:
                    html_str += f'{line}&nbsp;<br>'
        elif line and not skip_line:
            # Add the line with a non-breaking space
            html_str += f'{line}&nbsp;<br>'

    # Close the <p> tag
    html_str += '</p>'

    return html_str

def comment_strint_clean_up(comment):
    # Split the comment into lines
    lines = comment.split('\n')
    
    # Initialize an empty string
    comment_str = ""
        # Flag to indicate if we should skip the line due to Total score
    skip_line = False

    # Iterate through the lines of the comment
    for line in lines:
        line = line.strip()
        if 'Total' in line or 'total' in line:
            # Set the skip_line flag to True to skip subsequent lines
            skip_line = True
        elif line and not skip_line:
            # Add the line with a non-breaking space
            comment_str += f'{line}\n'
    return comment_str

def read_master_to_dataframe(file_path, delimiter='\t'):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print("Error reading file: {}".format(e))
        sys.exit(1)
    
    return df

def update_grading_sheets(master_sheet_path, partial_sheet_path, output_sheet_path):
    # Load the master grading sheet and the partial grading sheet into DataFrames
    # master_df = pd.read_csv(master_sheet_path)

    master_df = read_master_to_dataframe(master_sheet_path, delimiter='\t')

    partial_df = pd.read_csv(partial_sheet_path)

    # Assuming the column names for "Last Name" and "First Name" in both sheets are the same
    last_name_column = "Last Name"
    first_name_column = "First Name"
    
    # print("Master DataFrame:")
    # print(master_df.tail(10))
    
    # print("Partial DataFrame:")
    # print(partial_df.tail(10))
    
    total_pts_column = [col for col in master_df.columns if "Total Pts" in col][0]
    if total_pts_column:
        print(f"Updating grade for {total_pts_column}")
    else:
        print("** No total points column found **")
        exit(1)
    
    grading_note_column = [col for col in master_df.columns if "Grading Notes" in col][0]
    if grading_note_column:
        print(f"Updating comment for {grading_note_column}")
    else:
        print("** No comment column found, only updating grade **")
    
    grade_note_format_column = [col for col in master_df.columns if "Notes Format" in col][0]
    if grade_note_format_column:
        print(f"Updating comment with HTML format")
    else:
        print("Updating comment in string format")
    
    # Iterate through the master DataFrame and find matching rows in the partial DataFrame
    counter = 0
    print("\nMatching Rows:")
    for index, row in master_df.iterrows():
        first_name = row[first_name_column]
        last_name = row[last_name_column]
        
        matching_row = partial_df[(partial_df[first_name_column] == first_name) & (partial_df[last_name_column] == last_name)]
        if not matching_row.empty:
            counter+=1            
            
            if total_pts_column:
                # total_pts_column = total_pts_column[0]
                # update final score
                master_df.at[index, total_pts_column] = matching_row.iloc[0]["Final Score"]
            
            if grading_note_column:
                
                # update note format only if master_df.at[index, grading_note_column]
                if master_df.at[index, grading_note_column]:
                    if grade_note_format_column:
                        # print("Converting comment to HTML")
                        # update comment
                        master_df.at[index, grading_note_column] = convert_comment_to_html(matching_row.iloc[0]["All Comments"])
                    else:
                        # print("Updating comment")
                        master_df.at[index, grading_note_column] = matching_row.iloc[0]["All Comments"]
                    
                    master_df.at[index, "Notes Format"] = "HTML"
            
            if total_pts_column and grading_note_column:
                print(f"Updated {first_name} {last_name} score: {master_df.at[index, total_pts_column]}\n comment: {master_df.at[index, grading_note_column]}\n")
                    
    
    print("\nTotal Matching Rows: {}".format(counter))

    # Save the updated master grading sheet to a new Excel file
    master_df.to_csv(output_sheet_path, index=False)



if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python update_grading_sheets.py master_sheet.csv partial_sheet.csv updated_master_sheet.csv")
        sys.exit(1)

    master_sheet_path = sys.argv[1]
    partial_sheet_path = sys.argv[2]
    output_sheet_path = sys.argv[3]

    update_grading_sheets(master_sheet_path, partial_sheet_path, output_sheet_path)
