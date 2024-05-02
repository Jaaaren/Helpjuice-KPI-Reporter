import pandas as pd
from datetime import datetime, timedelta
import re
from collections import Counter
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


def excel_date_to_datetime(excel_date):
    try:
        # Check if the date is a missing or NaN value
        if pd.isna(excel_date):
            return None
        
        # Attempt to convert the excel_date to an integer
        serial_date = int(excel_date)
        excel_epoch = datetime(1899, 12, 30)  # Excel's default date system starts from this date
        delta = timedelta(days=serial_date)
        return excel_epoch + delta
    except ValueError:
        # Handles cases where conversion to int fails because the input is not numeric
        return None

def clean_query(query):
    # Remove non-alphanumeric characters and convert to lowercase
    query = re.sub(r'\W+', ' ', query).lower()
    # Split the query into individual words
    words = query.split()
    # Return the list of words
    return words

# Path setup
base_path = './'  # Directory containing this script and all Excel files

# Create PDF document
pdf_path = base_path + 'wiki_kpi_report.pdf'
doc = SimpleDocTemplate(pdf_path, pagesize=letter)
story = []
styles = getSampleStyleSheet()
title_style = styles['Heading1']
content_style = styles['BodyText']

# Read 'categories_analytics.xls'
categories_df = pd.read_excel(base_path + 'categories_analytics.xls', header=None)

date_format = '%m/%d/%Y'
start_date_serial = categories_df.iloc[0, 1]
end_date_serial = categories_df.iloc[1, 1]
start_date = pd.to_datetime(categories_df.iloc[0, 1], format=date_format, errors='coerce')
end_date = pd.to_datetime(categories_df.iloc[1, 1], format=date_format, errors='coerce')

if pd.isnull(start_date) or pd.isnull(end_date):
    print("Error: Start or end date is missing or invalid in the Excel file.")
    formatted_start_date = "Unavailable"
    formatted_end_date = "Unavailable"
else:
    formatted_start_date = start_date.strftime('%B %d, %Y')
    formatted_end_date = end_date.strftime('%B %d, %Y')
    
category_data = categories_df.iloc[4:14, :2].values.tolist()

# Read 'questions_analytics.xls'
questions_df = pd.read_excel(base_path + 'questions_analytics.xls')
questions_data = [
    [row[0], row[1], pd.to_datetime(row[2]).strftime('%B %d, %Y')] for row in questions_df.iloc[4:14, [1, 3, 8]].values
]
questions_data.insert(0, ['Article', 'Views', 'Last Updated'])  # Insert headers at the top of the data list

# Read 'searches.xls'
searches_df = pd.read_excel(base_path + 'searches.xls', header=None)
# Clean the search queries and split into individual words
searches_df[1] = searches_df[1].apply(clean_query)
# Flatten the list of lists into a single list containing all words
all_words = [word for sublist in searches_df[1] for word in sublist]
# Count the occurrences of each word using a Counter
word_counts = Counter(all_words)
# Get the 5 most common words and their counts
most_common_words = word_counts.most_common(5)
# Generate the search data for the report, converting counts to integers
search_data = [[word, int(count)] for word, count in most_common_words]
search_data.insert(0, ['Query', 'Searches'])  # Insert headers at the top of the data list

# Read 'users_analytics.xls'
users_df = pd.read_excel(base_path + 'users_analytics.xls', header=None)
users_df[3] = pd.to_numeric(users_df[3], errors='coerce')
top_users = users_df.iloc[4:].nlargest(10, 3)
user_data = top_users[[0, 3]].values.tolist()
user_data = [[name, int(views)] for name, views in user_data]  # Convert views to integers
user_data.insert(0, ['User', 'Views'])  # Insert headers at the top of the data list

# Title and Date Range (Centered)
title_style.alignment = 1  # 1 for center alignment
date_style = ParagraphStyle('dateStyle', parent=styles['Normal'], alignment=1)  # New style for date range

story.append(Paragraph('Wiki KPI Report', title_style))
story.append(Paragraph(f'Data from {formatted_start_date} to {formatted_end_date}', date_style))
story.append(Spacer(1, 12))

# Categories Table with Headers
category_data.insert(0, ['Categories', 'Views'])  # Insert headers at the top of the data list


# Categories Table
story.append(Paragraph('Top 10 Most Viewed Categories', title_style))
cat_table = Table(category_data, colWidths=[200, 100], spaceBefore=12, spaceAfter=12)
cat_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('GRID', (0, 0), (-1, -1), 1, colors.black)
]))
story.append(cat_table)

# Questions Table
story.append(Paragraph('Top 10 Most Viewed Articles', title_style))
column_widths = [300, 75, 95]  # Increase the width for the 'Article' column

quest_table = Table(questions_data, colWidths=column_widths, spaceBefore=12, spaceAfter=12)
quest_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('WORDWRAP', (0, 0), (-1, -1), 'LTR'),  # Set word wrap for all cells
    ('FONTSIZE', (0, 0), (-1, -1), 10),  # Optional: adjust font size for better fit
]))
story.append(quest_table)

# Searches Table
story.append(Paragraph('Top 5 Most Used Search Queries', title_style))
search_table = Table(search_data, colWidths=[200, 100], spaceBefore=12, spaceAfter=12)
search_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('GRID', (0, 0), (-1, -1), 1, colors.black)
]))
story.append(search_table)

# Users Table
story.append(Paragraph('Top 10 Users by Articles Viewed', title_style))
user_table = Table(user_data, colWidths=[200, 100], spaceBefore=12, spaceAfter=12)
user_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('GRID', (0, 0), (-1, -1), 1, colors.black)
]))
story.append(user_table)

# Build PDF
doc.build(story)
print("PDF report generated successfully!")

