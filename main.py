

from src.errors import FileReadError
from src.model.models.AARR_model import AARR_Model
from src.model.parsers.excel_model_parser import ExcelModelParser
from src.model.parsers.html_model_parser import HTMLModelParser
from src.model.parsers.pdf_model_parser import PDFModelParser
from src.model.readers.file_model_reader import FileReader


if __name__ == "__main__":
    reader = FileReader()
    reader.set_target("eprivacy.ecix.tech.html")
    # reader.set_target("RAT_HOSPITAL UNIVERSITARIO HLA MONCLOA.pdf")
    data = reader.read()

    if isinstance(data, FileReadError):
        print(f"Error reading file: {data.message}")
        exit(1)


    # parser = PDFModelParser(AARR_Model.FONT_SIZES_MAPPING)
    parser = HTMLModelParser()
    parsed_file = parser.parse(data)


    




# import time

# import pandas as pd
# import pdfplumber

# PDF_PATH = 'RAT_HOSPITAL UNIVERSITARIO HLA MONCLOA.pdf'
# OUTPUT_PATH = 'extracted_table_pdfplumber.xlsx'

# tables: list[dict[int, str]] = []
# start_time = time.time()

# with pdfplumber.open(PDF_PATH) as pdf, pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
#     for page_n in range(0, len(pdf.pages)):
#         page = pdf.pages[page_n]
#         table = page.extract_table(table_settings={"vertical_strategy": "lines",
#                                                 "horizontal_strategy": "lines", 
#                                                 "snap_tolerance": 10,
#                                                 "intersection_tolerance": 10,
#                                                 "join_tolerance": 10,
#                                                 "edge_min_length": 8,})
#         if table:
#             df = pd.DataFrame(table[1:], columns=table[0])
#             df = df.map(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x)
#             df = df.dropna(how='all')
#             tables.append({"page": page_n, "data": df})

#     composed_table: None | pd.DataFrame = None
#     for table in tables:
#         has_header = table["data"].columns.str.contains('TRATAMIENTO|PRECALIF.|FINALIDADES|NIV.RIESGO').all()
#         if has_header:
#             composed_table = table["data"]
#             continue  # skip header rows and concat to previous table
#             print("hello")

#         is_page_number_ok = table["page"] > 1 and table["page"] < 10
#         contains_4_fields = len(table["data"].columns) == 4
#         if not contains_4_fields:
#             break

#         is_fourth_field_numeric = table["data"].iloc[:, 3].apply(lambda x: x is None or x == '-' or float(x) > 0 ).all()
#         second_field_is_alto_or_bajo = table["data"].iloc[:, 1].isin(['ALTO', 'BAJO', '']).all()
        
#         if is_page_number_ok and contains_4_fields and is_fourth_field_numeric\
#             and second_field_is_alto_or_bajo:

#             if len(table["data"].columns) != 0:
#                 # Treat table's column headers as a data row
#                 header_row = pd.DataFrame([table["data"].columns], columns=composed_table.columns)
#                 data_fixed = table["data"].copy()
#                 data_fixed.columns = composed_table.columns
#                 composed_table = pd.concat([composed_table, header_row, data_fixed], ignore_index=True)
#             else:
#                 table["data"].columns = composed_table.columns
#                 composed_table = pd.concat([composed_table, table["data"]], ignore_index=True)


#             # clean rows with empty columns
#             composed_table = composed_table.dropna(how='all')
#             # reset index
#             composed_table = composed_table.reset_index(drop=True)

#     # Forward fill the first column if it's empty but other columns have data
#     for i in range(1, len(composed_table)):
#         first_col = composed_table.iat[i, 0]
        
#         if (pd.isna(first_col) or str(first_col).strip() == '') \
#             and composed_table.iloc[i, 1:].apply(lambda x: not (pd.isna(x) or str(x).strip() == '')).any():
            
#             composed_table.iat[i, 0] = composed_table.iat[i - 1, 0]

#     composed_table.to_excel(writer, sheet_name="page_" + str(table["page"]), index=False)

# end_time = time.time()
# print(f'Table extraction completed in {end_time - start_time:.2f} seconds.')
