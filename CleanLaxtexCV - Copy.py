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


def write_courses_to_file(text, f1):
    course_descriptions = {
        "330": "ME 330 (Computational Tools for Engineers)",
        "360": "ME 360 (Machine Design)",
        "461": "ME 461 (Introduction to Finite Element Analysis)",
        "563": "ME 563 (Nonlinear Finite Element Analysis)",
        "497": "Development course for Computational Tools for Engineers",
        "440": "Capstone Design",
    }
    excluded_courses = {"600", "596", "496", "494", "610"}
    courses_to_write = []

    # Split the text into sections based on the \subsubsection markers
    sections = re.split(r'\\subsubsection\{(\d{4})\}', text)[1:]
    # Group the sections into pairs of (year, courses)
    grouped_sections = [(sections[i], sections[i+1]) for i in range(0, len(sections), 2)]
    for year, courses in grouped_sections:
        # Extract unique course numbers
        course_nums = re.findall(r'(?:ME|M E) (\d{3})', courses)
        course_nums = set(course_nums)  # remove duplicates
        course_nums = [num for num in course_nums if num not in excluded_courses]  # exclude certain courses
        course_nums = sorted(course_nums, reverse=True)  # Sort in descending order if needed

        courses_to_write = []
        for num in course_nums:
            if num in course_descriptions:
                courses_to_write.append(course_descriptions[num])
            else:
                courses_to_write.append(f"ME{num}")

        # Write to file
        f1.write(f"\\textbf{{{year}}}\n")
        # add a line space in latex format
        f1.write(r'\\')
        f1.write(', '.join(courses_to_write))
        f1.write('\n\n')
        

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
    end_of_section = text.find(r'\subsubsection{2013}\label{section-10}') + len(r'\subsubsection{2013}\label{section-10}')  # Adjust if needed

    # Extract the course listings section
    course_section = text[start_of_section:end_of_section]

    # Process the course listings
    # Adjusted regex to capture both "ME" and "M E" formats
    sections = re.split(r'\\subsubsection\{(\d{4})\}', course_section)[1:]
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


def main():
    #open a file to read in C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV named main.tex 
    # open the file for reading
    f = open(r'C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\main.tex', 'r')
    #open another file for writing
    f1 = open(r'C:\Users\rhk12\OneDrive - The Pennsylvania State University\resume\CV\main_edited.tex', 'w')

    #read in the entire file and store it in a variable called text 
    text = f.read()
    
    # add my formating package
    text = add_custom_package(text)
    
    # Set the section colors
    text = set_section_colors(text)
    
     # Format the header
    text = format_header(text)

    #write_courses_to_file(text, f1)
    text = process_courses(text)
    
    f1.write(text)
    
    #search for the string and store the index in a variable called start
    #start = text.find(r'\subsection{Directed Student Learning}\label{directed-student-learning}')
    #write out the text from this location to the end of the file
    #f1.write(text[start:])
    #close the files
    f.close()
    f1.close()
    

if __name__ == "__main__":
    main()

    