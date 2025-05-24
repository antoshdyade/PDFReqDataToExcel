import re
import csv
import PyPDF2

# Open and read the PDF
file_path = 'candidate_compressed.pdf'
with open(file_path, 'rb') as file:
    reader = PyPDF2.PdfReader(file)
    all_text = ''
    for page in reader.pages:
        all_text += page.extract_text()

# Regular expression pattern including AIRANK
pattern = re.compile(
    r"(?P<slno>\d+)\s+Class IX\s+(?P<roll>\d+)\s+(?P<appno>\d+)\s+(?P<name>MR\..*?)\s+"
    r"(?P<gender>Male|Female)\s+(?P<cat>SC|[^ ]+)\s+(?P<domicile>[A-Z ]+?)\s+"
    r"(?P<marks>\d+)\s+(?P<result>Qualified.*?|Not Qualified.*?)\s+(?P<airank>\d+)",
    re.IGNORECASE
)

# Extract and filter records
filtered_data = []
for match in pattern.finditer(all_text):
    if (match.group('cat').strip() == 'SC' and
        match.group('domicile').strip().upper() == 'MAHARASHTRA' and
        match.group('gender').strip().lower() == 'male'):
        
        filtered_data.append({
            'SL. NO.': match.group('slno'),
            'ROLL': match.group('roll'),
            'APPNO': match.group('appno'),
            'NAME': match.group('name'),
            'GENDER': match.group('gender'),
            'CATEGORY': match.group('cat'),
            'DOMICILE': match.group('domicile'),
            'MARKS': match.group('marks'),
            'RESULT': match.group('result'),
            'AIRANK': match.group('airank'),
        })

# Write to CSV
csv_file = 'filtered_candidates.csv'
with open(csv_file, mode='w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=filtered_data[0].keys())
    writer.writeheader()
    writer.writerows(filtered_data)

print(f"Filtered records saved to {csv_file}")
