import camelot
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle,Spacer
from reportlab.pdfgen import canvas
import pandas as pd
import os

path = "ScopeTaggingDocXBRL_03042018.pdf"

tables = camelot.read_pdf(path, pages='9-18')

all_tables = [table.df for table in tables]

combined_df = pd.concat(all_tables, ignore_index=True)



concatenated_rows = combined_df.applymap(lambda x: ''.join(x.splitlines()) if isinstance(x, str) else x)


