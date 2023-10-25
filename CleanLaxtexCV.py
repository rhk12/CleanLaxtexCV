#create a python script to clean the latex cv

# import the necessary packages
import argparse
import math
import sys
import re
import shutil
import subprocess
import time
import glob
import csv
import random
import string
import datetime
import os 
from docx import Document 


def process_student_thesis_titles(text,dossier_file_path): 
    # Extract master's thesis titles
    masters_data = extract_student_titles(dossier_file_path)
 
    # Extract PhD dissertation titles
    phd_data = extract_phd_titles(dossier_file_path)
  
    # Extract postdoc titles
    postdoc_data = extract_postdoc_titles(dossier_file_path)
    
    # Extract undergrad thesis titles
    undergrad_data = extract_undergrad_student_titles(dossier_file_path)
     
    # Merge the dictionaries
    student_data = {**masters_data, **phd_data, **postdoc_data, **undergrad_data}

    # Replace problematic characters in titles
    student_data = replace_problematic_characters_in_titles(student_data)
        
    # Update the LaTeX content with thesis titles - does not work for postdoc titles see next function
    text = add_title_to_name2(text, student_data)
       
    # reformat phd section for name, title, date
    text = reformat_phd_section(text)
    
    # reformat masters section for name, title, date    
    text = update_masters_section(text)
    
    # reformat postdoc section for name, title, date
    text = update_postdoc_section(text)
    
    # reformat undergrad section for name, title, date
    text = update_undergraduate_section(text)
        
    return text

def replace_problematic_characters_in_titles(student_data):
    problematic_char = u'\u2013'  # En dash (U+2013)

    for name, title in student_data.items():
        if problematic_char in title:
            #print(f"Found character '{problematic_char}' in {name}: {title}")
            student_data[name] = title.replace(problematic_char, '--')

    return student_data

def update_undergraduate_section(text_data):
    # Define section label
    label = "Undergraduate Honors\nThesis"
    
    # Extract the Undergraduate Honors Thesis section
    section_pattern = rf"\\subsubsection{{{label}}}.*?\\begin{{enumerate}}.*?\\end{{enumerate}}"
    section_match = re.search(section_pattern, text_data, re.DOTALL)
    
    if not section_match:
        print("Undergraduate Honors Thesis section not found!")
        return text_data
    
    section_content = section_match.group(0)
    
    # Extract individual items
    item_pattern = r"Undergraduate Honors Thesis\. \((.*?)\).\\\\\s+Advised: (.*?), \"(.*?)\","
    items = re.findall(item_pattern, section_content, re.DOTALL)
    
    # Format items
    formatted_items = []
    for time_period, name, title in items:
        formatted_item = f"\\item {name}, \"{title}\", {time_period}"
        formatted_items.append(formatted_item)
    
    # Combine formatted items
    formatted_section = "\\subsubsection{Undergraduate Honors\nThesis}\n\\begin{enumerate}\n\\def\\labelenumi{\\arabic{enumi}.}\n" + "\n".join(formatted_items) + "\n\\end{enumerate}"
    
    # Replace the old section with the new formatted section in the text data
    updated_text_data = text_data.replace(section_content, formatted_section)
    
    return updated_text_data

def update_postdoc_section(text_data):
    # Define section label
    label = "Postdoctoral Mentorship"
    
    # Extract the Postdoctoral Mentorship section
    section_pattern = rf"\\subsubsection{{{label}}}.*?\\begin{{enumerate}}.*?\\end{{enumerate}}"
    section_match = re.search(section_pattern, text_data, re.DOTALL)
    
    if not section_match:
        print("Postdoctoral Mentorship section not found!")
        return text_data
    
    section_content = section_match.group(0)
    
    # Extract individual items
    item_pattern = r"Postdoctoral Mentorship\. \((.*?)\).\\\\\s+Advised: (.*?), \"(.*?)\","
    items = re.findall(item_pattern, section_content, re.DOTALL)
    
    # Format items
    formatted_items = []
    for time_period, name, title in items:
        formatted_item = f"\\item {name}, \"{title}\", {time_period}"
        formatted_items.append(formatted_item)
    
    # Combine formatted items
    formatted_section = "\\subsubsection{Postdoctoral Mentorship}\n\\begin{enumerate}\n\\def\\labelenumi{\\arabic{enumi}.}\n" + "\n".join(formatted_items) + "\n\\end{enumerate}"
    
    # Replace the old section with the new formatted section in the text data
    updated_text_data = text_data.replace(section_content, formatted_section)
    
    return updated_text_data

def update_masters_section(text_data):
    # Define section label
    label = "Master's Thesis"
    
    # Extract the Master's Thesis section
    section_pattern = rf"\\subsubsection{{{label}}}.*?\\begin{{enumerate}}.*?\\end{{enumerate}}"
    section_match = re.search(section_pattern, text_data, re.DOTALL)
    
    if not section_match:
        print("Master's Thesis section not found!")
        return text_data
    
    section_content = section_match.group(0)
    
    # Extract individual items
    item_pattern = r"Master's Thesis\. \((.*?)\).\\\\\s+Advised: (.*?), \"(.*?)\","
    items = re.findall(item_pattern, section_content, re.DOTALL)
    
    # Format items
    formatted_items = []
    for time_period, name, title in items:
        formatted_item = f"\\item {name}, \"{title}\", {time_period}"
        formatted_items.append(formatted_item)
    
    # Combine formatted items with consistent spacing
    #formatted_section = f"\\subsubsection{{{label}}}\n\\begin{{enumerate}}\n\\def\\labelenumi{{\\arabic{{enumi}}.}}\n" + "\n".join(formatted_items) + "\n\\end{{enumerate}}\n"
    # Combine formatted items
    formatted_section = "\\subsubsection{Master's Thesis}\n\\begin{enumerate}\n\\def\\labelenumi{\\arabic{enumi}.}\n" + "\n".join(formatted_items) + "\n\\end{enumerate}"

    # Replace the old section with the new formatted section in the text data
    updated_text_data = text_data.replace(section_content, formatted_section)
    
    return updated_text_data

def reformat_masters_section(text_data):
    # Define section label
    label = "Master's Thesis"
    
    # Extract the Master's Thesis section
    section_pattern = rf"(\\subsubsection{{{label}.*?}}.*?\\begin{enumerate}.*?\\end{enumerate})"
    section_match = re.search(section_pattern, text_data, re.DOTALL)
    
    if not section_match:
        return text_data
    
    section_content = section_match.group(1)
    
    # Pattern to extract student details
    student_pattern = rf"{label}. \((.*?)\).+?Advised: (.*?), \"(.*?)\","
    student_matches = re.findall(student_pattern, section_content, re.DOTALL)
    
    # Create a new formatted section
    new_section = f"\\subsubsection{{{label}}}\n\n\\begin{enumerate}\n"
    for time_period, name, title in student_matches:
        # Split name into first and last names
        names = name.split()
        first_name = names[0]
        last_name = names[-1]
        
        # Add to the new section
        new_section += f"\\item\n  {first_name} {last_name}, \"{title}\", {time_period}\n"
    
    new_section += "\\end{enumerate}\n"
    
    # Replace the old section with the new one
    text_data = text_data.replace(section_content, new_section)
    
    return text_data

def reformat_sections(text_data):
    # Define section labels
    section_labels = {
        "postdoc": "Postdoctoral Mentorship",
        "undergrad": "Undergraduate Honors Thesis",
        "masters": "Master's Thesis"
    }
    
    # Function to reformat a specific section
    def reformat_section(section_content, label):
        student_pattern = rf"{label}. \((.*?)\).+?Advised: (.*?), \"(.*?)\","
        student_matches = re.findall(student_pattern, section_content, re.DOTALL)
        
        # Create a new formatted section
        new_section = f"\\subsubsection{{{label}}}\n\n\\begin{enumerate}\n"
        for time_period, name, title in student_matches:
            # Split name into first and last names
            names = name.split()
            first_name = names[0]
            last_name = names[-1]
            
            # Add to the new section
            new_section += f"\\item\n  {first_name} {last_name}, \"{title}\", {time_period}\n"
        
        new_section += "\\end{enumerate}\n"
        return new_section
    
    # Iterate over each section label and reformat
    for section, label in section_labels.items():
        section_pattern = rf"(\\subsubsection{{{label}.*?}}.*?\\begin{enumerate}.*?\\end{enumerate})"
        section_match = re.search(section_pattern, text_data, re.DOTALL)
        
        if section_match:
            section_content = section_match.group(1)
            new_section = reformat_section(section_content, label)
            text_data = text_data.replace(section_content, new_section)
    
    return text_data

def extract_undergrad_student_titles(file_path):
    #file_path = r"C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\20231021-095357-CDT.docx"

    # Check if the file exists
    if not os.path.exists(file_path):
        return f"File not found at '{file_path}'"
    
    # If the file exists, proceed to read its content
    doc = Document(file_path)
    full_text = [para.text for para in doc.paragraphs]

    # Define the sub-section titles
    sub_section_title = "Undergraduate Honors Thesis Advisor"
    end_section_title = "THE SCHOLARSHIP OF Research and Creative Accomplishments"

    # Find the start of the sub-section
    start_index = next((i for i, text in enumerate(full_text) if sub_section_title in text), None)
    if start_index is None:
        return {}

    # Find the end of the sub-section
    end_index = next((i for i, text in enumerate(full_text[start_index:]) if end_section_title in text), None)
    if end_index is None:
        end_index = len(full_text)
    else:
        end_index += start_index

    # Extract undergrad student names and thesis titles
    undergrad_data = {}
    for line in full_text[start_index + 1:end_index]:
        if ", Undergraduate." in line:
            name, title_with_dates = line.split(", Undergraduate.", 1)
            title = title_with_dates.split(". (")[0].strip()
            
            # Remove the "Date Graduated" part from the title
            title = re.sub(r"\. Date Graduated:.*$", "", title)
            
            undergrad_data[name.strip()] = title

    return undergrad_data

def add_postdoc_work(text, postdoc_data):
    for name, title in postdoc_data.items():
        # Split the name into first and last names
        last_name, first_initial = name.split(', ')
        first_initial = first_initial.strip()[0]  # Get the first character of the first name
        
        # Construct two patterns: one for full name and one for first initial
        full_name_pattern = rf"(\\subsubsection\{{Postdoctoral Mentorship\}}.*?Advised: ){first_initial}.*?{last_name}"
        initial_pattern = rf"(\\subsubsection\{{Postdoctoral Mentorship\}}.*?Advised: ){first_initial} {last_name}"
        
        # Try matching with the full name
        if re.search(full_name_pattern, text, flags=re.DOTALL):
            text = re.sub(full_name_pattern, f"\\1{first_initial}. {last_name}, \"{title}\"", text, flags=re.DOTALL)
            print(f"Updated {first_initial}. {last_name}'s entry.")  # Debugging print statement
        # If that doesn't work, try matching with the first initial
        elif re.search(initial_pattern, text, flags=re.DOTALL):
            text = re.sub(initial_pattern, f"\\1{first_initial}. {last_name}, \"{title}\"", text, flags=re.DOTALL)
            print(f"Updated {first_initial}. {last_name}'s entry using initial.")  # Debugging print statement
        else:
            print(f"Pattern not found for: {first_initial}. {last_name} or {first_initial} {last_name}")  # Debugging print statement

    return text

def reformat_phd_section(text_data):
    # Extract the Ph.D. section
    phd_section_pattern = r"(\\subsubsection{Ph\.D\. Dissertation}.*?\\begin{enumerate}.*?\\end{enumerate})"
    phd_section_match = re.search(phd_section_pattern, text_data, re.DOTALL)
    
    if not phd_section_match:
        print("Ph.D. section not found in the provided text.")
        return text_data
    
    phd_section = phd_section_match.group(1)
    
    # Extract details for each student
    student_pattern = r"Ph\.D\. Dissertation\. \((.*?)\).+?Advised: (.*?), \"(.*?)\","
    student_matches = re.findall(student_pattern, phd_section, re.DOTALL)
    
    # Create a new formatted Ph.D. section
    new_phd_section = "\\subsubsection{Ph.D. Dissertation}\n\n\\begin{enumerate}\n"
    for time_period, name, title in student_matches:
        # Split name into first and last names
        names = name.split()
        first_name = names[0]
        rest_of_name = " ".join(names[1:])
        
        # Add to the new Ph.D. section
        new_phd_section += f"\\item\n  {first_name} {rest_of_name}, \"{title}\", {time_period}\n"

    
    new_phd_section += "\\end{enumerate}\n"
    
    # Replace the original Ph.D. section with the new one
    updated_text_data = text_data.replace(phd_section, new_phd_section)
    
    return updated_text_data

def add_title_to_name(text, data):
    for name, title in data.items():
        # Adjust for the name format in the postdoc section
        if ',' in name:
            last_name, first_initial = name.split(', ')
            formatted_name_full = f"{first_initial}. {last_name}"
            formatted_name_last = last_name
        else:
            formatted_name_full = name
            formatted_name_last = name.split()[0]  # Just the first name
        
        # Check which naming convention exists in the LaTeX document
        if formatted_name_full in text:
            pattern = re.escape(formatted_name_full)
        elif formatted_name_last in text:
            pattern = re.escape(formatted_name_last)
        else:
            print(f"Pattern not found for: {name}")
            continue
        
        # Extract the time period
        time_period_match = re.search(r"\((.*?)\)", text)
        if time_period_match:
            time_period = time_period_match.group(1)
        else:
            time_period = ""

        # If we're dealing with the postdoc section
        if "Postdoctoral Mentorship" in title:
            postdoc_section_pattern = r"(\\subsubsection\{Postdoctoral Mentorship\}.*?)" + pattern
            match = re.search(postdoc_section_pattern, text, re.DOTALL)
            if match:
                text = text.replace(match.group(0), match.group(0).replace(f"Advised: {formatted_name_full}", f"{formatted_name_last}, \"{title}\", {time_period}"))
        # If we're dealing with the undergraduate section
        elif "Undergraduate Honors Thesis" in title:
            undergrad_section_pattern = r"(\\subsubsection\{Undergraduate Honors\s*Thesis\}.*?Advised: )" + pattern
            match = re.search(undergrad_section_pattern, text, re.DOTALL)
            if match:
                text = text.replace(match.group(0), match.group(0).replace(f"Advised: {formatted_name_full}", f"{formatted_name_last}, \"{title}\", {time_period}"))
        # For Master's and Ph.D. students
        else:
            if re.search(pattern, text):
                text = re.sub(pattern, f"{formatted_name_last}, \"{title}\", {time_period}", text)
    return text

def add_title_to_name2(text, data):
    # Extract the section between \subsection{Directed Student Learning} and \subsection{Teaching Experience}
    start_index = text.find(r'\subsection{Directed Student Learning}')
    end_index = text.find(r'\subsection{Teaching Experience}', start_index)
    
    if start_index == -1 or end_index == -1:
        print("Directed Student Learning or Teaching Experience section not found!")
        return text

    section = text[start_index:end_index]

    for name, title in data.items():
        # Adjust for the name format in the postdoc section
        if ',' in name:
            last_name, first_initial = name.split(', ')
            formatted_name_full = f"{first_initial}. {last_name}"
            formatted_name_last = last_name
        else:
            formatted_name_full = name
            formatted_name_last = name.split()[0]  # Just the first name
        
        # Check which naming convention exists in the section
        if formatted_name_full in section:
            pattern = re.escape(formatted_name_full)
        elif formatted_name_last in section:
            pattern = re.escape(formatted_name_last)
        else:
            print(f"Pattern not found for: {name}")
            continue
        
        # Extract the time period
        time_period_match = re.search(r"\((.*?)\)", section)
        if time_period_match:
            time_period = time_period_match.group(1)
        else:
            time_period = ""

        # If we're dealing with the postdoc section
        if "Postdoctoral Mentorship" in title:
            postdoc_section_pattern = r"(\\subsubsection\{Postdoctoral Mentorship\}.*?)" + pattern
            match = re.search(postdoc_section_pattern, section, re.DOTALL)
            if match:
                section = section.replace(match.group(0), match.group(0).replace(f"Advised: {formatted_name_full}", f"{formatted_name_last}, \"{title}\", {time_period}"))
        # If we're dealing with the undergraduate section
        elif "Undergraduate Honors Thesis" in title:
            undergrad_section_pattern = r"(\\subsubsection\{Undergraduate Honors\s*Thesis\}.*?Advised: )" + pattern
            match = re.search(undergrad_section_pattern, section, re.DOTALL)
            if match:
                section = section.replace(match.group(0), match.group(0).replace(f"Advised: {formatted_name_full}", f"{formatted_name_last}, \"{title}\", {time_period}"))
        # For Master's and Ph.D. students
        else:
            if re.search(pattern, section):
                section = re.sub(pattern, f"{formatted_name_last}, \"{title}\", {time_period}", section)

    # Replace the original section with the modified one
    text = text[:start_index] + section + text[end_index:]
    
    return text

def add_undergrad_titles(text, undergrad_data):
    for name, title in undergrad_data.items():
        # Convert the name format from 'Last, F.' to 'First Last'
        last_name, first_initial = name.split(', ')
        # We'll use a placeholder for the first name since we don't have the full first name
        first_name_placeholder = f"{first_initial}.*"
        
        # Define the pattern to search for the student's entry in the LaTeX content
        pattern = rf"(Advised: {first_name_placeholder} {last_name})"
        
        # Check if the pattern exists in the text
        if re.search(pattern, text):
            # Replace the pattern with the pattern + title
            text =  re.sub(pattern, f"\\1, \"{title}\"", text)
            print(f"Updated {first_initial}. {last_name}'s entry.")
        else:
            print(f"Pattern for {first_initial}. {last_name} not found in the text!")
    
    return text

def extract_postdoc_titles(file_path):
    #file_path = r"C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\20231021-095357-CDT.docx"
    
    # Check if the file exists
    if not os.path.exists(file_path):
        return f"File not found at '{file_path}'"
    
    # If the file exists, proceed to read its content
    doc = Document(file_path)
    full_text = [para.text for para in doc.paragraphs]

    # Define the sub-section titles
    sub_section_title = "Postdoctoral Mentorship Advisor"
    end_section_title = "Research Activity Advisor"

    # Find the start of the sub-section
    start_index = next((i for i, text in enumerate(full_text) if sub_section_title in text), None)
    if start_index is None:
        return {}

    # Find the end of the sub-section
    end_index = next((i for i, text in enumerate(full_text[start_index:]) if end_section_title in text), None)
    if end_index is None:
        end_index = len(full_text)
    else:
        end_index += start_index

    # Extract postdoc names and work titles
    postdoc_data = {}
    for line in full_text[start_index + 1:end_index]:
        if ". " in line:
            name, title_with_dates = line.split(". ", 1)
            title = title_with_dates.split(". (")[0].strip()
            postdoc_data[name.strip()] = title

    return postdoc_data

def extract_phd_titles(file_path):
    #file_path = r"C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\20231021-095357-CDT.docx"
    
    # Check if the file exists
    if not os.path.exists(file_path):
        return f"File not found at '{file_path}'"
    
    # If the file exists, proceed to read its content
    doc = Document(file_path)
    full_text = [para.text for para in doc.paragraphs]

    # Define the sub-section titles
    sub_section_title = "Ph.D. Dissertation Advisor"
    end_section_title = "Ph.D. Dissertation Committee Member"  # This is the section that follows the Advisor section

    # Find the start of the sub-section
    start_index = next((i for i, text in enumerate(full_text) if sub_section_title in text), None)
    if start_index is None:
        print(f"Section '{sub_section_title}' not found")
        return {}

    # Find the end of the sub-section
    end_index = next((i for i, text in enumerate(full_text[start_index + 1:]) if end_section_title in text), None)
    if end_index is None:
        end_index = len(full_text)
    else:
        end_index += start_index + 1

    # Extract student names and dissertation titles
    student_data = {}
    for line in full_text[start_index + 1:end_index]:
        if ", " in line and "Ph.D." in line:
            name = line.split(",")[0].strip()
            title_start = line.find("Ph.D.") + 6
            title_end = line.find(".", title_start)
            title = line[title_start:title_end].strip()
            student_data[name] = title

    return student_data

def extract_student_titles(file_path):
    #file_path = r"C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\20231021-095357-CDT.docx"
    
    # Check if the file exists
    if not os.path.exists(file_path):
        return f"File not found at '{file_path}'"
    
    # If the file exists, proceed to read its content
    doc = Document(file_path)
    full_text = [para.text for para in doc.paragraphs]

    # Define the sub-section titles
    sub_section_titles = ["Master's Thesis Advisor", "Master\u2019s Thesis Advisor"]
    end_section_titles = ["Master's Thesis Committee Member", "Master\u2019s Thesis Committee Member"]

    # Find the start of the sub-section
    start_index = None
    for title in sub_section_titles:
        start_index = next((i for i, text in enumerate(full_text) if title in text), None)
        if start_index is not None:
            break

    if start_index is None:
        print(f"Section 'Master's Thesis Advisor' not found")
        return

    # Find the end of the sub-section
    end_index = None
    for title in end_section_titles:
        end_index = next((i for i, text in enumerate(full_text[start_index:]) if title in text), None)
        if end_index is not None:
            break

    if end_index is None:
        end_index = len(full_text)
    else:
        end_index += start_index

    # Extract student names and thesis titles
    student_data = {}
    for line in full_text[start_index + 1:end_index]:
        if ", " in line and "MS." in line:
            name = line.split(",")[0].strip()
            title_start = line.find("MS.") + 4
            title_end = line.find(".", title_start)
            title = line[title_start:title_end].strip()
            student_data[name] = title

    return student_data

def extract_subsubsection(text, title):
    # Use regex to extract the subsubsection based on the title
    pattern = re.compile(r'\\subsubsection{' + re.escape(title) + r'}.*?\\end{enumerate}', re.DOTALL)
    match = pattern.search(text)
    if match:
        return match.group(0)
    else:
        print(f"Title '{title}' not found!")
        return None

def reorder_student_sections(text):
    # Define the desired order
    order = ["Ph.D. Dissertation", "Master's Thesis", "Postdoctoral Mentorship", "Undergraduate Honors\nThesis"]
    
    # Extract each subsubsection based on the order
    sections = [extract_subsubsection(text, title) for title in order]
    
    # Check if any section is missing and remove it
    sections = [section for section in sections if section is not None]
    
    # Concatenate the reordered sections
    reordered_sections = "\n".join(sections)
    
    # Replace the original Directed Student Learning section with the reordered one
    start_idx = text.find(r'\subsection{Directed Student Learning}')
    end_idx = text.find(r'\subsection', start_idx + 1)
    if end_idx == -1:
        end_idx = len(text)
    
    return text[:start_idx] + r'\subsection{Directed Student Learning}' + '\n' + reordered_sections + text[end_idx:]

def boldface_name_in_publications(text):
    # Extract the publications section
    start_publications = text.find(r'\subsection{Publications}\label{publications}')
    start_presentations = text.find(r'\subsection{Presentations}\label{presentations}')
    
    if start_publications == -1 or start_presentations == -1:
        print("Publications or Presentations subsection not found!")
        return text

    publications_section = text[start_publications:start_presentations]
    
    # Boldface the name in the publications section
    patterns = [r'(Kraft, R\. H\.)', r'(Kraft,)', r'(Kraft, R\.)']
    for pattern in patterns:
        publications_section = re.sub(pattern, r'\\textbf{\1}', publications_section)
    
    # Replace the old publications section with the modified one
    modified_text = text[:start_publications] + publications_section + text[start_presentations:]
    
    return modified_text

def reorder_publications(content):
    # Extract the content from \subsection{Publications} up to \subsection{Presentations}
    pattern = r'(\\subsection\{Publications\}.*?)(\\subsection\{Presentations\})'
    match = re.search(pattern, content, re.DOTALL)
    
    if not match:
        return content  # No change if the pattern isn't found
    
    publications_section = match.group(1)
    presentations_label = match.group(2)
    
    # Split into individual subsubsections
    subsubsections = re.findall(r'(\\subsubsection\{.*?\}.*?)(?=\\subsubsection|$)', publications_section, re.DOTALL)
    
    # Define the desired order
    order = ["Journal Article", "Conference Proceeding", "Book Chapters", "Other"]
    
    # Create a dictionary with subsubsection titles as keys and their content as values
    subsubsection_dict = {}
    for subsubsection in subsubsections:
        title = re.search(r'\\subsubsection\{(.*?)\}', subsubsection).group(1)
        subsubsection_dict[title] = subsubsection

    # Construct the reordered content based on the desired order
    sorted_content = ''.join([subsubsection_dict[title] for title in order if title in subsubsection_dict])
    
    # Replace old subsubsections with sorted ones in the publications section
    updated_publications = publications_section.replace(''.join(subsubsections), sorted_content)
    
    # Construct the updated content
    return content.replace(publications_section, updated_publications)

def capitalize_subsections(text_content):
    # Find all \subsection{...} patterns
    subsections = re.findall(r'\\subsection\{(.*?)\}', text_content)

    # For each found subsection, replace it with its capitalized version
    for subsection in subsections:
        capitalized_subsection = subsection.upper()
        text_content = text_content.replace(r"\subsection{" + subsection + "}", r"\subsection{" + capitalized_subsection + "}")

    return text_content

def add_date_to_header(text_content):
    # Get the current month and year
    now = datetime.datetime.now()
    month_year = now.strftime("%B %Y")

    # Define the header settings
    header_settings = r"""
\usepackage{fancyhdr}
\usepackage{xcolor}
\fancypagestyle{firstpage}{
    \fancyhf{} % Clear all headers and footers first
    \rhead{\textcolor{lightgray}{""" + month_year + r"""}}
    \renewcommand{\headrulewidth}{0pt} % No header rule
}
\thispagestyle{firstpage}
"""

    # Insert the header settings right before \begin{document}
    marker = r"\begin{document}"
    preamble_end = text_content.find(marker)
    
    if preamble_end != -1:
        modified_content = text_content[:preamble_end] + header_settings + text_content[preamble_end:]
        return modified_content
    else:
        print("Warning: Marker for preamble end (\\begin{document}) not found in the text content.")
        return text_content  # Return the original content if marker is not found

def format_header(text):
    # Define the block of text to be removed
    to_remove = r"""
\textbf{Dr. Reuben H. Kraft}\\
The Pennsylvania State University\\
EN - Mechanical Engineering\\
(814) 867-4570\\
Email: rhk12@psu.edu
""".strip()  # .strip() is used to remove any leading/trailing whitespace

    # Replace the block with an empty string
    text = text.replace(to_remove, '')

    # Add the desired formatted name after \begin{document}
    new_header = r"\LARGE \textbf{\textsc{REUBEN KRAFT - CURRICULUM VITA }} \\\rule{\linewidth}{0.4pt}"
    # Add a line space
    new_header += '\n\n'
    new_header += r"\normalsize % Return to the default font size"
    
    # Find the position of \begin{document} and insert the new header after it
    insert_pos = text.find(r'\begin{document}') + len(r'\begin{document}')
    text = text[:insert_pos] + '\n\n' + new_header + '\n\n' + text[insert_pos:]

    return text
def add_custom_package(text, package_name="mystyle"):
    # Find the position of \author{}
    insert_pos = text.find(r'\author{')
    
    # If \author{} is found, insert the \usepackage command before it
    if insert_pos != -1:
        text = text[:insert_pos] + f"\\usepackage{{{package_name}}}\n" + text[insert_pos:]
    
    return text
def set_section_colors(text):
    # Define the LaTeX commands to change the colors
    color_commands = r"""
\usepackage{titlesec}
\usepackage{color}
\titleformat{\subsection}
  {\normalfont\large\bfseries\color{blue}} % format
  {\thesubsection} % label
  {1em} % sep
  {} % before-code
\titleformat{\subsubsection}
  {\normalfont\normalsize\bfseries\color{red}} % format
  {\thesubsubsection} % label
  {1em} % sep
  {} % before-code
"""

    # Insert the color commands before \begin{document}
    insert_pos = text.find(r'\begin{document}')
    if insert_pos != -1:
        text = text[:insert_pos] + color_commands + text[insert_pos:]

    return text

def process_courses(text):
    course_descriptions = {
        "330": "ME 330 (Computational Tools for Engineers)",
        "360": "ME 360 (Machine Design)",
        "461": "ME 461 (Introduction to Finite Element Analysis)",
        "563": "ME 563 (Nonlinear Finite Element Analysis)",
        "497": "Development course for Computational Tools for Engineers",
        "440": "Capstone Design",
    }
    excluded_courses = {"600", "596", "496", "494", "610"}
    
    # Find the start and end of the course listings section
    start_of_section = text.find(r'\subsection{Teaching Experience}\label{teaching-experience}')
    end_of_section = text.find(r'\subsection{Service}\label{service}')

    # If the start or end markers are not found, return the original text
    if start_of_section == -1 or end_of_section == -1:
        print("Warning: Course section markers not found.")
        return text

    # Process the course listings
    sections = re.split(r'\\subsubsection\{(\d{4})\}', text[start_of_section:end_of_section])[1:]
    grouped_sections = [(sections[i], sections[i+1]) for i in range(0, len(sections), 2)]

    new_course_section = r'\subsection{Teaching Experience}\label{teaching-experience}' + '\n\n'

    for year, courses in grouped_sections:
        course_nums = re.findall(r'(?:ME|M\s?E) (\d{3})', courses)
        course_nums = set(course_nums)  # remove duplicates
        course_nums = [num for num in course_nums if num not in excluded_courses]  # exclude certain courses
        course_nums = sorted(course_nums, reverse=True)  # Sort in descending order if needed

        courses_to_write = []
        for num in course_nums:
            if num in course_descriptions:
                courses_to_write.append(course_descriptions[num])
            else:
                courses_to_write.append(f"ME{num}")

        # Append to new course section text
        new_course_section += f"\\textbf{{{year}}}\n"
        new_course_section += r'\\'
        new_course_section += ', '.join(courses_to_write)
        new_course_section += '\n\n'

    # Replace the original course listings section with the new one
    new_text = text[:start_of_section] + new_course_section + text[end_of_section:]

    return new_text

def replace_straight_quotes_with_latex_quotes(text_data):
    # This pattern matches phrases enclosed in straight double quotes, accounting for potential newlines
    pattern = r'"([\s\S]*?)"'
    
    # This function replaces straight double quotes with LaTeX quotes
    def replace_quotes(match):
        content = match.group(1)
        # Check if the content already starts with LaTeX-style quotes
        if content.startswith("``"):
            return f'"{content}"'  # Return the content as-is
        return f"``{content}''"
    
    # Use re.sub to replace all occurrences in the text
    corrected_text = re.sub(pattern, replace_quotes, text_data)
    
    # Debugging for the Service subsection
    service_section_pattern = r"\\subsection{Service}.*?\\subsubsection{Department}"
    service_section_match = re.search(service_section_pattern, corrected_text, re.DOTALL)
    
    return corrected_text

def clean_service_section(text_data):
    # Define subsubsection labels
    labels = ["College", "Department", "University", "Profession", "Society"]
    
    for label in labels:
        # Extract the subsubsection content
        section_pattern = rf"\\subsubsection{{{label}}}.*?(?=\\subsubsection|$)"
        section_match = re.search(section_pattern, text_data, re.DOTALL)
        
        if not section_match:
            print(f"{label} section not found!")
            continue
        
        section_content = section_match.group(0)
        
        # Remove the redundant mention of the subsubsection title from each item
        cleaned_section = re.sub(rf"{label}, ", "", section_content)
        
        # Replace the old section with the cleaned section in the text data
        text_data = text_data.replace(section_content, cleaned_section)
    
    return text_data

def highlight_mentored_authors_astericks2(text_data):
    # Extract last names of mentored students and postdocs
    student_types = {
        'Ph.D. Dissertation': 'Ph.D. students',
        'Master\'s Thesis': 'Master students',
        'Postdoctoral Mentorship': 'Postdoc researchers',
        'Undergraduate Honors\nThesis': 'Undergrad students'
    }
    
    mentored_lastnames = []

    for subsection, label in student_types.items():
        start_index = text_data.find(rf'\subsubsection{{{subsection}}}')
        if start_index == -1:
            match = re.search(re.escape(subsection), text_data)
            if match:
                start_index = match.start()
            else:
                continue

        end_index = text_data.find(r'\subsubsection{', start_index + 1)
        if end_index == -1:
            end_index = len(text_data)

        section = text_data[start_index:end_index]
        pattern = r"\\item\s+(.*?),\s+``"
        matches = re.findall(pattern, section, re.DOTALL)
        
        for full_name in matches:
            lastname = full_name.split()[-1]
            mentored_lastnames.append(lastname)

    # Search through each publication and look for these last names
    start_publications = text_data.find(r'\subsection{Publications}')
    end_publications = text_data.find(r'\subsection{', start_publications + 1)
    if end_publications == -1:
        end_publications = len(text_data)

    publications_section = text_data[start_publications:end_publications]

    # Add a LaTeX asterisk after the name if it's found
    for lastname in mentored_lastnames:
        # Modified pattern to account for an optional middle initial
        pattern = rf"({lastname}(?:, [A-Z]\. ?[A-Z]?)?)"
        replacement = r"\1\\textsuperscript{*}"
        publications_section = re.sub(pattern, replacement, publications_section)

    # Replace the old publications section with the modified one
    modified_text = text_data[:start_publications] + publications_section + text_data[end_publications:]

    # Add mentored student note
    journal_article_section_start = modified_text.find(r'\subsubsection{Journal Article}\label{journal-article}')
    if journal_article_section_start == -1:
        print("Journal Article section not found in the text.")
        return modified_text

    note_text = "\n\\textit{Mentored student and postdoc co-authors have an astericks after their name.}\n"
    modified_text = modified_text[:journal_article_section_start] + note_text + modified_text[journal_article_section_start:]

    return modified_text

def underline_mentored_authors_with_note(text_data):
    # Extract last names of mentored students and postdocs
    student_types = {
        'Ph.D. Dissertation': 'Ph.D. students',
        'Master\'s Thesis': 'Master students',
        'Postdoctoral Mentorship': 'Postdoc researchers',
        'Undergraduate Honors\nThesis': 'Undergrad students'
    }
    
    mentored_lastnames = []

    for subsection, label in student_types.items():
        start_index = text_data.find(rf'\subsubsection{{{subsection}}}')
        if start_index == -1:
            match = re.search(re.escape(subsection), text_data)
            if match:
                start_index = match.start()
            else:
                continue

        end_index = text_data.find(r'\subsubsection{', start_index + 1)
        if end_index == -1:
            end_index = len(text_data)

        section = text_data[start_index:end_index]
        pattern = r"\\item\s+(.*?),\s+``"
        matches = re.findall(pattern, section, re.DOTALL)
        
        for full_name in matches:
            lastname = full_name.split()[-1]
            mentored_lastnames.append(lastname)

    # Search through each publication and look for these last names
    start_publications = text_data.find(r'\subsection{Publications}')
    end_publications = text_data.find(r'\subsection{', start_publications + 1)
    if end_publications == -1:
        end_publications = len(text_data)

    publications_section = text_data[start_publications:end_publications]

    # Add a LaTeX note about underlining mentored authors
    note_text = "\n\\textit{Mentored student and postdoc co-authors are underlined.}\n"
    journal_article_section_start = publications_section.find(r'\subsubsection{Journal Article}\label{journal-article}')
    publications_section = publications_section[:journal_article_section_start] + note_text + publications_section[journal_article_section_start:]

    # Underline the name if it's found
    for lastname in mentored_lastnames:
        # Modified pattern to account for an optional middle initial
        pattern = rf"({lastname}(?:, [A-Z]\. ?[A-Z]?)?)"
        replacement = r"\\underline{\1}"
        publications_section = re.sub(pattern, replacement, publications_section)

    # Replace the old publications section with the modified one
    modified_text = text_data[:start_publications] + publications_section + text_data[end_publications:]

    return modified_text

def create_custom_titles_for_sections(text_data):
    
    # change the title of the Jounral Article section
    text_data = text_data.replace(r'\subsubsection{Journal Article}\label{journal-article}', r'\subsubsection{Journal Articles}\label{journal-article}')
    
    # change the title of the Conference Proceeding section
    text_data = text_data.replace(r'\subsubsection{Conference Proceeding}\label{conference-proceeding}', r'\subsubsection{Conference Proceedings}\label{conference-proceeding}')
    
    # change the title of the Book Chapter section
    text_data = text_data.replace(r'\subsubsection{Book Chapter}\label{book-chapter}', r'\subsubsection{Book Chapters}\label{book-chapter}')

    # change the title of the Other section to Preprints and Technical Reports
    text_data = text_data.replace(r'\subsubsection{Other}\label{other}', r'\subsubsection{Preprints and Technical Reports}\label{other}')

    # in the Presentations section change the title of the Invited Talks section to Invited Talks and Seminars
    text_data = text_data.replace(r'\subsubsection{Invited}\label{invited}', r'\subsubsection{Invited Talks and Seminars}\label{invited}')

    # in the Presentations sections change the title of Uncategoried Presentations to Confererences and Workshops Presentations
    text_data = text_data.replace(r'\subsubsection{Uncategorized}\label{uncategorized}', r'\subsubsection{Conferences and Workshops}\label{uncategorized-presentations}')

    return text_data

def replace_section_colors(text, subsection_color='black', subsubsection_color='black'):
    # Define a pattern to search for the \titleformat command for \subsection and replace the color
    subsection_pattern = re.compile(r"(\\titleformat\{\\subsection\}.*?\\color\{)(blue)(\}.*)", re.DOTALL | re.VERBOSE)
    subsubsection_pattern = re.compile(r"(\\titleformat\{\\subsubsection\}.*?\\color\{)(red)(\}.*)", re.DOTALL | re.VERBOSE)

    # Define the replacement function for \subsection
    def subsection_replacement(match):
        #print("Subsection color found:", match.group(2))
        return f"{match.group(1)}{subsection_color}{match.group(3)}"

    # Define the replacement function for \subsubsection
    def subsubsection_replacement(match):
        #print("Subsubsection color found:", match.group(2))
        return f"{match.group(1)}{subsubsection_color}{match.group(3)}"

    # Perform the replacement for \subsection
    updated_text = subsection_pattern.sub(subsection_replacement, text)
    # Perform the replacement for \subsubsection
    updated_text = subsubsection_pattern.sub(subsubsection_replacement, updated_text)

    return updated_text


def main():
    #open a file to read in C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV named main.tex 
    # open the file for reading
    #f = open(r'C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\main.tex', 'r')
    f = open(r'C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\main.tex', 'r', encoding='utf-8')

    #open another file for writing
    f1 = open(r'C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\main_edited.tex', 'w')

    #dossier_file_path = r"C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\20231021-095357-CDT.docx"
    dossier_file_path = r"C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\20231024-203540-CDT.docx"
    
    #read in the entire file and store it in a variable called text 
    text = f.read()
        
    # add my formating package
    text = add_custom_package(text)
    
    # Set the section colors
    text = set_section_colors(text)
      
     # Format the header
    text = format_header(text)
    
    # add the date to the latex file
    text = add_date_to_header(text)

    # write_courses_to_file(text, f1)
    text = process_courses(text)
    
    # reorder the publications
    text = reorder_publications(text)
    
    # boldface my name in the publications
    text = boldface_name_in_publications(text)
        
    # reorder the student sections
    text = reorder_student_sections(text)
    
    # process the student thesis titles
    text = process_student_thesis_titles(text,dossier_file_path)

    # clean up the service section to remove extra text
    text = clean_service_section(text)

    # replace straight quotes with latex quotes
    text = replace_straight_quotes_with_latex_quotes(text)
    
    # add emphasis for mentees  - pick between astericks or underlines
    #text = highlight_mentored_authors_astericks2(text)
    text = underline_mentored_authors_with_note(text)
    
    # add custom titles for sections
    text = create_custom_titles_for_sections(text)
    
    # colorize the subsection and subsubsection titles
    text = replace_section_colors(text, subsection_color='black', subsubsection_color='black')
    
    # make subsection text uppercase - needs to be last
    text = capitalize_subsections(text)
  
    f1.write(text)
    
    #close the files
    f.close()
    f1.close()
    

if __name__ == "__main__":
    main()

    