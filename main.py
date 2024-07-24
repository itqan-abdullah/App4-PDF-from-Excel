import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name="Sheet 1")
    
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    date = filename.split("-")[1]
    
    pdf = FPDF(orientation="P",unit="mm",format="A4")
    pdf.add_page()
    pdf.set_font(family="Times",size=16,style="B")
    
    pdf.cell(w=50,h=8,txt=f"Invoice nr.{invoice_nr}",align='L',ln=1)
    pdf.cell(w=50,h=8,txt=f"Date {date}",align='L',ln=1)
    
    columns = list(df.columns)
    columns = [column.replace("_"," ").title() for column in columns] 
    pdf.set_font(family="Times",size=10,style='B')
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30,h=8,txt=str(columns[0]),border=True)
    pdf.cell(w=50,h=8,txt=str(columns[1]),border=True)
    pdf.cell(w=30,h=8,txt=str(columns[2]),border=True)
    pdf.cell(w=30,h=8,txt=str(columns[3]),border=True)
    pdf.cell(w=30,h=8,txt=str(columns[4]),ln=1,border=True)
    
    
    
    for index, row in df.iterrows():
        pdf.set_font(family="Times",size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30,h=8,txt=str(row["product_id"]),border=True)
        pdf.cell(w=50,h=8,txt=str(row["product_name"]),border=True)
        pdf.cell(w=30,h=8,txt=str(row["amount_purchased"]),border=True)
        pdf.cell(w=30,h=8,txt=str(row["price_per_unit"]),border=True)
        pdf.cell(w=30,h=8,txt=str(row["total_price"]),ln=1,border=True)
    
    
    pdf.set_font(family="Times",size=12,style='B')
    pdf.set_text_color(80,80,80)
    
    pdf.cell(w=30,h=8,txt=str(' '),border=True)
    pdf.cell(w=50,h=8,txt=str(" "),border=True)
    pdf.cell(w=30,h=8,txt=str(" "),border=True)
    pdf.cell(w=30,h=8,txt=str(" "),border=True)
    pdf.cell(w=30,h=8,txt=str(sum(df["total_price"])),ln=1,border=True)
        
    pdf.cell(w=0,h=8,txt=f"The total due amount is {sum(df['total_price'])} Euros.",ln=1)
    pdf.cell(w=25,h=8,txt=f"PythonHow")
    pdf.image('pythonhow.png',w=10)
    
    pdf.output(f"pdfs/{filename}.pdf")
    