import argparse
import csv
from docx import Document

# parse command line arguments to get CSV filename
parser = argparse.ArgumentParser(description='Import CSV file.')
parser.add_argument('csvfile', type=str, help='path to CSV file')
args = parser.parse_args()

# open Word template and replace fields with CSV values

def process_file(template_file, header_info, output_filename, row_data):
        for i, data in enumerate(row_data[1:], 1):
            # loop through each paragraph in template and replace fields with CSV values
            for paragraph in template_file.paragraphs:
                for fieldname in header_info:
                    if fieldname in paragraph.text:
                        # replace field with value from CSV file
                        value = row_data[header_info.index(fieldname)]
                        paragraph.text = paragraph.text.replace(fieldname, value)

        # save modified document as a new file with last name and index of data
        template_file.save(output_filename)

if __name__ == '__main__':
    with open(args.csvfile, 'r') as f:
        reader = csv.reader(f)

        # read CSV header row to get fieldnames
        headers = next(reader)
        for row_idx, row in enumerate(reader, 2):  # start from row 2
            template = Document('template.docx')
            filename = f'Letter_{row[0]}.docx'
            process_file(template, headers, filename, row)
