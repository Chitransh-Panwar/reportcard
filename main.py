import pandas as pd
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import io, os

# --- Helper Functions ---
def safe_value(val):
    if pd.isna(val):
        return ""
    return str(val)

def safe_num(val):
    if pd.isna(val):
        return 0
    if isinstance(val, str):
        stripped = val.strip().lower()
        # FIX Bug 2: handle 'aa' as absent
        if stripped in ("absent", "aa"):
            return 0
        try:
            return int(float(stripped))
        except (ValueError, TypeError):
            return 0
    try:
        return int(val)
    except (ValueError, TypeError):
        return 0

def format_mark(val):
    if pd.isna(val):
        return ""
    if isinstance(val, str):
        stripped = val.strip()
        # FIX Bug 2: handle 'AA' as absent marker
        if stripped.lower() in ("absent", "aa","a","ab"):
            return "AB"
        try:
            return str(int(float(stripped)))
        except (ValueError, TypeError):
            return ""
    try:
        return str(int(val))
    except (ValueError, TypeError):
        return ""

# FIX Bug 3: PAN formatted properly (no scientific notation)
def format_pan(val):
    if pd.isna(val):
        return ""
    try:
        return str(int(float(val)))
    except (ValueError, TypeError):
        return str(val)

# FIX Bug 1: class normalizer handles '11.0' → '11'
def normalize_class(val):
    if pd.isna(val):
        return ""
    try:
        # Convert float like 11.0 → int → str → '11'
        return str(int(float(str(val)))).strip().lower()
    except (ValueError, TypeError):
        return str(val).strip().lower()

# --- Load Excel ---
try:
    df = pd.read_excel("11,12sci.xlsx")
except FileNotFoundError:
    print("❌ Error: 11 B5.xlsx not found.")
    exit()

df.columns = df.columns.str.strip()
os.makedirs("reports", exist_ok=True)
template_file = "temp.pdf"

print(f"✅ Loaded {len(df)} student records.")
print("Starting PDF generation...\n")

merger = PdfMerger()

for index, row in df.iterrows():
    output_pdf_writer = PdfWriter()

    # --- Format DOB ---
    if "DOB" in row:
        date_value = row["DOB"]
        if pd.notna(date_value):
            formatted_date = pd.to_datetime(date_value).strftime("%d-%m-%Y")
        else:
            formatted_date = ""
    else:
        formatted_date = ""

    # --- Load Template ---
    try:
        existing_pdf_reader = PdfReader(open(template_file, "rb"))
    except Exception as e:
        print(f"Error reading template PDF: {e}")
        exit()

    edited_page = None

    # --- Detect Stream (based on filled subjects) ---
    science_subjects = [
        "A Phy T", "A Chem T", "A Phy E", "A Chem E",
        "HY Phy T", "HY Chem T", "HY Phy E", "HY Chem E"
    ]
    # Updated Commerce subjects based on your images
    commerce_subjects = [
        "A ACC T", "A B.ST T", "A ECO T", "HY ACC T", "HY B.ST T",
        "A ACC E", "A B.ST E", "A ECO E", "HY ACC E", "HY B.ST E",
        "HY MATH/ECO T", "A MATH/ECO T", "HY MATH/ECO E", "A MATH/ECO E"
    ]
    # New Arts subjects based on your images
    arts_subjects = [
        "A CIV T", "A HIS/GEO T", "A DRW/ECO T", "HY CIV T", "HY HIS/GEO T", "HY DRW/ECO T",
        "A CIV E", "A HIS/GEO E", "A DRW/ECO E", "HY CIV E", "HY HIS/GEO E", "HY DRW/ECO E"
    ]

    science_count = sum(1 for s in science_subjects if safe_num(row.get(s)) > 0)
    commerce_count = sum(1 for s in commerce_subjects if safe_num(row.get(s)) > 0)
    arts_count = sum(1 for s in arts_subjects if safe_num(row.get(s)) > 0)

    # Determine stream based on highest count
    if science_count == 0 and commerce_count == 0 and arts_count == 0:
        stream = "Unknown"
    elif science_count >= commerce_count and science_count >= arts_count:
        stream = "Science"
    elif commerce_count >= science_count and commerce_count >= arts_count:
        stream = "Commerce"
    else:
        stream = "Arts"


    # --- Normalize Class value ---
    class_val = str(safe_value(row["Class"])).strip().lower()

    # --- Loop through template pages ---
    for page_num, page in enumerate(existing_pdf_reader.pages):
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)

        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(width, height))
        can.setFont("Helvetica", 12)

        # ---------------------------------------
        # CLASS 11–12 SECTION
        # ---------------------------------------
        if class_val in ['11', '12', 'xi', 'xii']:

            # ✅ SCIENCE STREAM
            if stream == "Science" and page_num == 1:
                can.drawString(180, height - 117, str(safe_num(row["PAN"])))
                can.drawString(159, height - 170, safe_value(row["NAME OF STUDENTS"]))
                can.drawString(159, height - 203, safe_value(row["FATHERS NAME"]))
                can.drawString(440, height - 170, safe_value(row["MOTHERS NAME"]))
                can.drawString(445, height - 117, safe_value(row['Class']))
                can.drawString(175, height - 156, safe_value(row["S.R.N"]))
                can.drawString(455, height - 135, safe_value(row['Section']))
                can.drawString(440, height - 203, formatted_date)
                can.drawString(470, height - 153, safe_value(row["R.N"]))
                
                can.drawString(140, height - 236, safe_value(row["ADDRESS"]))

                # Half-Yearly Marks - using format_mark for display
                can.drawString(174, height - 355, format_mark(row["HY Hindi T"]))
                can.drawString(208, height - 355, format_mark(row["HY Hindi E"]))
                can.drawString(174, height - 380, format_mark(row["HY English T"]))
                can.drawString(208, height - 380, format_mark(row["HY English E"]))
                can.drawString(174, height - 405, format_mark(row["HY Math/BIO T"]))
                can.drawString(208, height - 405, format_mark(row["HY Math/BIO E"]))
                can.drawString(174, height - 430, format_mark(row["HY Phy T"]))
                can.drawString(208, height - 430, format_mark(row["HY Phy E"]))
                can.drawString(174, height - 452, format_mark(row["HY Chem T"]))
                can.drawString(208, height - 452, format_mark(row["HY Chem E"])) 
                can.drawString(243, height - 405, format_mark(row["HY BIO PR"]))
                can.drawString(243, height - 430, format_mark(row["HY PHY PR"]))
                can.drawString(243, height - 452, format_mark(row["HY CHEM PR"]))
                
                # Half-Yearly Totals - using safe_num for calculations
                can.drawString(275, height - 355, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))
                can.drawString(275, height - 380, str(safe_num(row["HY English T"]) + safe_num(row["HY English E"])))
                can.drawString(275, height - 405, str(safe_num(row["HY Math/BIO T"]) +safe_num(row["HY BIO PR"])+ safe_num(row["HY Math/BIO E"])))
                can.drawString(275, height - 430, str(safe_num(row["HY Phy T"]) + safe_num(row["HY PHY PR"]) + safe_num(row["HY Phy E"])))
                can.drawString(275, height - 452, str(safe_num(row["HY Chem T"]) + safe_num(row["HY CHEM PR"]) + safe_num(row["HY Chem E"])))
                can.drawString(174, height - 495, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY English T"]) +
                                                  safe_num(row["HY Math/BIO T"]) + safe_num(row["HY Phy T"]) +
                                                  safe_num(row["HY Chem T"]) ))
                can.drawString(205, height - 495, str(safe_num(row["HY Hindi E"]) + safe_num(row["HY English E"]) +
                                                  safe_num(row["HY Math/BIO E"]) + safe_num(row["HY Phy E"]) +
                                                  safe_num(row["HY Chem E"]) ))
                can.drawString(272, height - 495, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY English T"]) +
                                                  safe_num(row["HY Math/BIO T"])+ safe_num(row["HY Phy T"])+safe_num(row["HY BIO PR"])+
                                                  safe_num(row["HY Chem T"]) + safe_num(row["HY CHEM PR"]) + safe_num(row["HY PHY PR"]) +
                                                  safe_num(row["HY Hindi E"]) + safe_num(row["HY English E"]) +
                                                  safe_num(row["HY Math/BIO E"]) + safe_num(row["HY Phy E"]) +
                                                  safe_num(row["HY Chem E"])))

                # Annual Exam - using format_mark for display
                can.drawString(340, height - 355, format_mark(row["A Hindi T"]))
                can.drawString(380, height - 355, format_mark(row["A Hindi E"]))
                can.drawString(340, height - 380, format_mark(row["A English T"]))
                can.drawString(380, height - 380, format_mark(row["A English E"]))
                can.drawString(340, height - 405, format_mark(row["A Math/BIO T"]))
                can.drawString(380, height - 405, format_mark(row["A Math/BIO E"]))
                can.drawString(340, height - 430, format_mark(row["A Phy T"]))
                can.drawString(380, height - 430, format_mark(row["A Phy E"]))
                can.drawString(340, height - 452, format_mark(row["A Chem T"]))
                can.drawString(380, height - 452, format_mark(row["A Chem E"]))
            
                # Practicals
                can.drawString(415, height - 405, format_mark(row["A BIO PR"]))
                can.drawString(415, height - 430, format_mark(row["A PHY PR"]))
                can.drawString(415, height - 452, format_mark(row["A CHEM PR"]))

                # Annual Totals
                can.drawString(450, height - 355, str(safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"])))
                can.drawString(450, height - 380, str(safe_num(row["A English T"]) + safe_num(row["A English E"])))
                can.drawString(450, height - 405, str(safe_num(row["A Math/BIO T"]) +safe_num(row["A BIO PR"])+ safe_num(row["A Math/BIO E"])))
                can.drawString(450, height - 430, str(safe_num(row["A Phy T"]) + safe_num(row["A PHY PR"]) + safe_num(row["A Phy E"])))
                can.drawString(450, height - 452, str(safe_num(row["A Chem T"]) + safe_num(row["A CHEM PR"]) + safe_num(row["A Chem E"])))
                can.drawString(340, height - 495, str(safe_num(row["A Hindi T"]) + safe_num(row["A English T"]) +
                                                  safe_num(row["A Math/BIO T"]) + safe_num(row["A Phy T"]) +
                                                  safe_num(row["A Chem T"])))
                can.drawString(380, height - 495, str(safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                                  safe_num(row["A Math/BIO E"]) + safe_num(row["A Phy E"]) +
                                                  safe_num(row["A Chem E"])))
                can.drawString(450, height - 495, str(safe_num(row["A PHY PR"]) + safe_num(row["A CHEM PR"]) +
                                                  safe_num(row["A Hindi T"]) + safe_num(row["A English T"]) +
                                                  safe_num(row["A Math/BIO T"]) + safe_num(row["A Phy T"]) +
                                                  safe_num(row["A Chem T"])  +safe_num(row["A BIO PR"])+
                                                  safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                                  safe_num(row["A Math/BIO E"]) + safe_num(row["A Phy E"]) +
                                                  safe_num(row["A Chem E"])))


                # --- GRAND TOTAL (only if at least one Annual mark exists) ---
                annual_fields = [
                        "A Hindi T", "A Hindi E", "A English T", "A English E", "A Math/BIO T", "A Math/BIO E",
                        "A Phy T", "A Phy E", "A Chem T", "A Chem E", "A PHY PR", "A CHEM PR"
                ]
                if any(pd.notna(row.get(col, None)) for col in annual_fields):
                        can.drawString(525, height - 355, str(safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"]) +
                                                      safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))
                        can.drawString(525, height - 380, str(safe_num(row["A English T"]) + safe_num(row["A English E"]) +
                                                      safe_num(row["HY English T"]) + safe_num(row["HY English E"])))
                        can.drawString(525, height - 405, str(safe_num(row["A Math/BIO T"]) +safe_num(row["HY BIO PR"])+ safe_num(row["A Math/BIO E"]) +
                                                      safe_num(row["HY Math/BIO T"]) +safe_num(row["A BIO PR"])+ safe_num(row["HY Math/BIO E"])))
                        can.drawString(525, height - 430, str(safe_num(row["A Phy T"]) + safe_num(row["A PHY PR"]) +
                                                      safe_num(row["A Phy E"]) + safe_num(row["HY PHY PR"])+ safe_num(row["HY Phy T"]) +
                                                      safe_num(row["HY Phy E"])))
                        can.drawString(525, height - 452, str(safe_num(row["A Chem T"]) + safe_num(row["A CHEM PR"]) +
                                                      safe_num(row["A Chem E"]) + safe_num(row["HY Chem T"]) +
                                                      safe_num(row["HY Chem E"])+ safe_num(row["HY CHEM PR"])))
                
                        can.drawString(525, height - 495, str( 
                                                      safe_num(row["A Chem T"])+ safe_num(row["HY CHEM PR"]) + safe_num(row["A CHEM PR"]) +
                                                      safe_num(row["A Chem E"]) + safe_num(row["HY Chem T"]) +
                                                      safe_num(row["HY Chem E"]) + safe_num(row["A Phy T"]) +safe_num(row["HY BIO PR"])+
                                                      safe_num(row["A PHY PR"]) + safe_num(row["A Phy E"]) +safe_num(row["A BIO PR"])+
                                                      safe_num(row["HY Phy T"])+ safe_num(row["HY PHY PR"]) + safe_num(row["HY Phy E"]) +
                                                      safe_num(row["A Math/BIO T"]) + safe_num(row["A Math/BIO E"]) +
                                                      safe_num(row["HY Math/BIO T"]) + safe_num(row["HY Math/BIO E"]) +
                                                      safe_num(row["A English T"]) + safe_num(row["A English E"]) +
                                                      safe_num(row["HY English T"]) + safe_num(row["HY English E"]) +
                                                      safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"]) +
                                                      safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))   
                                        
                
                can.save()
                packet.seek(0)
                new_pdf_reader = PdfReader(packet)
                overlay_page = new_pdf_reader.pages[0]
                page.merge_page(overlay_page)
                edited_page = page

            # ✅ COMMERCE STREAM (Updated with new subjects)
            elif stream == "Commerce" and page_num == 0:
                can.drawString(180, height - 117, str(safe_num(row["PAN"])))
                can.drawString(159, height - 170, safe_value(row["NAME OF STUDENTS"]))
                can.drawString(159, height - 203, safe_value(row["FATHERS NAME"]))
                can.drawString(440, height - 170, safe_value(row["MOTHERS NAME"]))
                can.drawString(445, height - 117, safe_value(row['Class']))
                can.drawString(175, height - 156, safe_value(row["S.R.N"]))
                can.drawString(455, height - 135, safe_value(row['Section']))
                can.drawString(440, height - 203, formatted_date)
                can.drawString(470, height - 153, safe_value(row["R.N"]))
                
                can.drawString(140, height - 236, safe_value(row["ADDRESS"]))
                    
                # Half-Yearly Marks with format_mark - Updated subjects
                can.drawString(200, height - 350, format_mark(row["HY Hindi T"]))
                can.drawString(236, height - 350, format_mark(row["HY Hindi E"]))
                can.drawString(200, height - 373, format_mark(row["HY English T"]))
                can.drawString(236, height - 373, format_mark(row["HY English E"]))
                # Accountancy (replacing ACC/CIV)
                can.drawString(200, height - 400, format_mark(row["HY ACC T"]))
                can.drawString(236, height - 400, format_mark(row["HY ACC E"]))
                # Business Studies (replacing B.ST/GEO/HIS)
                can.drawString(200, height - 428, format_mark(row["HY B.ST T"]))
                can.drawString(236, height - 428, format_mark(row["HY B.ST E"]))
                # Maths/Eco (replacing Math/DRG)
                can.drawString(200, height - 449, format_mark(row["HY MATH/ECO T"]))
                can.drawString(236, height - 449, format_mark(row["HY MATH/ECO E"])) 
 
                # Half-Yearly Totals - Updated for new subjects
                can.drawString(270, height - 350, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))
                can.drawString(270, height - 373, str(safe_num(row["HY English T"]) + safe_num(row["HY English E"])))
                can.drawString(270, height - 400, str(safe_num(row["HY ACC T"]) + safe_num(row["HY ACC E"])))
                can.drawString(270, height - 428, str(safe_num(row["HY B.ST T"]) + safe_num(row["HY B.ST E"])))
                can.drawString(270, height - 449, str(safe_num(row["HY MATH/ECO T"]) + safe_num(row["HY MATH/ECO E"])))
                can.drawString(200, height - 495, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY English T"]) +
                                                  safe_num(row["HY MATH/ECO T"]) + safe_num(row["HY B.ST T"]) +
                                                  safe_num(row["HY ACC T"])  ))
                can.drawString(236, height - 495, str(safe_num(row["HY Hindi E"]) + safe_num(row["HY English E"]) +
                                                  safe_num(row["HY MATH/ECO E"]) + safe_num(row["HY B.ST E"]) +
                                                  safe_num(row["HY ACC E"])  ))
                can.drawString(270, height - 495, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY English T"]) +
                                                  safe_num(row["HY MATH/ECO T"]) + safe_num(row["HY B.ST T"]) +
                                                  safe_num(row["HY ACC T"]) +safe_num(row["HY Hindi E"]) + safe_num(row["HY English E"]) +
                                                  safe_num(row["HY MATH/ECO E"]) + safe_num(row["HY B.ST E"]) +
                                                  safe_num(row["HY ACC E"]) ))
                
                # Annual Exam with format_mark - Updated subjects
                can.drawString(355, height - 350, format_mark(row["A Hindi T"]))
                can.drawString(390, height - 350, format_mark(row["A Hindi E"]))
                can.drawString(355, height - 373, format_mark(row["A English T"]))
                can.drawString(390, height - 373, format_mark(row["A English E"]))
                # Accountancy
                can.drawString(355, height - 400, format_mark(row["A ACC T"]))
                can.drawString(390, height - 400, format_mark(row["A ACC E"]))
                # Business Studies
                can.drawString(355, height - 428, format_mark(row["A B.ST T"]))
                can.drawString(390, height - 428, format_mark(row["A B.ST E"]))
                # Maths/Eco
                can.drawString(355, height - 449, format_mark(row["A MATH/ECO T"]))
                can.drawString(390, height - 449, format_mark(row["A MATH/ECO E"]))
                
                # Annual Totals - Updated for new subjects
                can.drawString(435, height - 350, str(safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"])))
                can.drawString(435, height - 373, str(safe_num(row["A English T"]) + safe_num(row["A English E"])))
                can.drawString(435, height - 400, str(safe_num(row["A ACC T"]) + safe_num(row["A ACC E"])))
                can.drawString(435, height - 428, str(safe_num(row["A B.ST T"]) + safe_num(row["A B.ST E"])))
                can.drawString(435, height - 449, str(safe_num(row["A MATH/ECO T"])  + safe_num(row["A MATH/ECO E"])))
                can.drawString(355, height - 495, str(safe_num(row["A Hindi T"]) + safe_num(row["A English T"]) +
                                                  safe_num(row["A MATH/ECO T"])  +
                                                  safe_num(row["A ACC T"])+ safe_num(row["A B.ST T"])))
                can.drawString(390, height - 495, str(safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                                  safe_num(row["A MATH/ECO E"]) +
                                                  safe_num(row["A ACC E"])+ safe_num(row["A B.ST E"])))
                can.drawString(435, height - 495, str(safe_num(row["A Hindi T"]) + safe_num(row["A English T"]) +
                                                  safe_num(row["A MATH/ECO T"])  +
                                                  safe_num(row["A ACC T"])+ safe_num(row["A B.ST T"])+safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                                  safe_num(row["A MATH/ECO E"])  +
                                                  safe_num(row["A ACC E"])+ safe_num(row["A B.ST E"])))
                
                annual_fields = [
                        "A Hindi T", "A Hindi E", "A English T", "A English E", "A MATH/ECO T", "A MATH/ECO E",
                        "A ACC T", "A ACC E", "A B.ST T", "A B.ST E"
                ]
                if any(pd.notna(row.get(col, None)) for col in annual_fields):
                        can.drawString(525, height - 350, str(safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"]) +
                                                      safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))
                        can.drawString(525, height - 373, str(safe_num(row["A English T"]) + safe_num(row["A English E"]) +
                                                      safe_num(row["HY English T"]) + safe_num(row["HY English E"])))
                        can.drawString(525, height - 400, str(safe_num(row["A ACC T"]) + safe_num(row["A ACC E"]) +
                                                      safe_num(row["HY ACC T"]) + safe_num(row["HY ACC E"])))
                        can.drawString(525, height - 428, str(safe_num(row["A B.ST T"]) +
                                                      safe_num(row["A B.ST E"]) + safe_num(row["HY B.ST T"]) +
                                                      safe_num(row["HY B.ST E"])))
                        can.drawString(525, height - 449, str(safe_num(row["A MATH/ECO T"]) + safe_num(row["HY MATH/ECO T"]) +
                                                      safe_num(row["A MATH/ECO E"])  +
                                                      safe_num(row["HY MATH/ECO E"])))
                        
                
                        can.drawString(525, height - 495, str( 
                                                      +safe_num(row["A MATH/ECO T"]) + safe_num(row["HY MATH/ECO T"]) +
                                                      safe_num(row["A MATH/ECO E"])  +
                                                      safe_num(row["HY MATH/ECO E"])+safe_num(row["A B.ST T"]) +
                                                      safe_num(row["A B.ST E"]) + safe_num(row["HY B.ST T"]) +
                                                      safe_num(row["HY B.ST E"])+safe_num(row["A ACC T"]) + safe_num(row["A ACC E"]) +
                                                      safe_num(row["HY ACC T"]) + safe_num(row["HY ACC E"])+safe_num(row["A English T"]) + safe_num(row["A English E"]) +
                                                      safe_num(row["HY English T"]) + safe_num(row["HY English E"])+safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"]) +
                                                      safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))   
               

                can.save()
                packet.seek(0)
                new_pdf_reader = PdfReader(packet)
                overlay_page = new_pdf_reader.pages[0]
                page.merge_page(overlay_page)
                edited_page = page

            # ✅ ARTS STREAM (New - similar structure to Science/Commerce)
            elif stream == "Arts" and page_num == 4:  # Assuming Arts uses page 2, adjust if needed
                can.drawString(180, height - 117, str(safe_num(row["PAN"])))
                can.drawString(159, height - 170, safe_value(row["NAME OF STUDENTS"]))
                can.drawString(159, height - 203, safe_value(row["FATHERS NAME"]))
                can.drawString(440, height - 170, safe_value(row["MOTHERS NAME"]))
                can.drawString(445, height - 117, safe_value(row['Class']))
                can.drawString(175, height - 156, safe_value(row["S.R.N"]))
                can.drawString(455, height - 135, safe_value(row['Section']))
                can.drawString(440, height - 203, formatted_date)
                can.drawString(470, height - 153, safe_value(row["R.N"]))
                
                can.drawString(140, height - 236, safe_value(row["ADDRESS"]))

                # Half-Yearly Marks - Arts subjects
                can.drawString(170, height - 355, format_mark(row["HY Hindi T"]))
                can.drawString(204, height - 355, format_mark(row["HY Hindi E"]))
                can.drawString(170, height - 380, format_mark(row["HY English T"]))
                can.drawString(204, height - 380, format_mark(row["HY English E"]))
                # Civics
                can.drawString(170, height - 405, format_mark(row["HY CIV T"]))
                can.drawString(204, height - 405, format_mark(row["HY CIV E"]))
                # Hist/Geo
                can.drawString(170, height - 430, format_mark(row["HY HIS/GEO T"]))
                can.drawString(204, height - 430, format_mark(row["HY HIS/GEO E"]))
                # Draw/Eco
                can.drawString(170, height - 452, format_mark(row["HY DRW/ECO T"]))
                can.drawString(204, height - 452, format_mark(row["HY DRW/ECO E"])) 
                
                can.drawString(240, height - 430, format_mark(row["HY GEO PR"]))
                
                # Half-Yearly Totals
                can.drawString(270, height - 355, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))
                can.drawString(270, height - 380, str(safe_num(row["HY English T"]) + safe_num(row["HY English E"])))
                can.drawString(270, height - 405, str(safe_num(row["HY CIV T"]) + safe_num(row["HY CIV E"])))
                can.drawString(270, height - 430, str(safe_num(row["HY HIS/GEO T"]) + safe_num(row["HY HIS/GEO E"])+safe_num(row["HY GEO PR"])))
                can.drawString(270, height - 452, str(safe_num(row["HY DRW/ECO T"]) + safe_num(row["HY DRW/ECO E"])))
                
                # Theory Total
                can.drawString(170, height - 495, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY English T"]) +
                                                  safe_num(row["HY CIV T"]) + safe_num(row["HY HIS/GEO T"]) +
                                                  safe_num(row["HY DRW/ECO T"]) ))
                # Practical/Internal Total
                can.drawString(200, height - 495, str(safe_num(row["HY Hindi E"]) + safe_num(row["HY English E"]) +
                                                  safe_num(row["HY CIV E"]) + safe_num(row["HY HIS/GEO E"]) +
                                                  safe_num(row["HY DRW/ECO E"]) ))
                # Grand Total
                can.drawString(270, height - 495, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY English T"]) +
                                                  safe_num(row["HY CIV T"]) + safe_num(row["HY HIS/GEO T"]) +
                                                  safe_num(row["HY DRW/ECO T"]) +
                                                  safe_num(row["HY Hindi E"]) + safe_num(row["HY English E"]) +
                                                  safe_num(row["HY CIV E"]) + safe_num(row["HY HIS/GEO E"]) +safe_num(row["HY GEO PR"])+
                                                  safe_num(row["HY DRW/ECO E"])))

                # Annual Exam - Arts subjects
                can.drawString(340, height - 355, format_mark(row["A Hindi T"]))
                can.drawString(378, height - 355, format_mark(row["A Hindi E"]))
                can.drawString(340, height - 380, format_mark(row["A English T"]))
                can.drawString(378, height - 380, format_mark(row["A English E"]))
                # Civics
                can.drawString(340, height - 405, format_mark(row["A CIV T"]))
                can.drawString(378, height - 405, format_mark(row["A CIV E"]))
                # Hist/Geo
                can.drawString(340, height - 430, format_mark(row["A HIS/GEO T"]))
                can.drawString(378, height - 430, format_mark(row["A HIS/GEO E"]))
                # Draw/Eco
                can.drawString(340, height - 452, format_mark(row["A DRW/ECO T"]))
                can.drawString(378, height - 452, format_mark(row["A DRW/ECO E"]))
                
                can.drawString(400, height - 430, format_mark(row["A GEO PR"]))

                # Annual Totals
                can.drawString(450, height - 355, str(safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"])))
                can.drawString(450, height - 380, str(safe_num(row["A English T"]) + safe_num(row["A English E"])))
                can.drawString(450, height - 405, str(safe_num(row["A CIV T"]) + safe_num(row["A CIV E"])))
                can.drawString(450, height - 430, str(safe_num(row["A HIS/GEO T"]) + safe_num(row["A HIS/GEO E"])+safe_num(row["A GEO PR"])))
                can.drawString(450, height - 452, str(safe_num(row["A DRW/ECO T"]) + safe_num(row["A DRW/ECO E"])))
                
                # Theory Total
                can.drawString(340, height - 495, str(safe_num(row["A Hindi T"]) + safe_num(row["A English T"]) +
                                                  safe_num(row["A CIV T"]) + safe_num(row["A HIS/GEO T"]) +
                                                  safe_num(row["A DRW/ECO T"])))
                # Practical/Internal Total
                can.drawString(373, height - 495, str(safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                                  safe_num(row["A CIV E"]) + safe_num(row["A HIS/GEO E"]) +
                                                  safe_num(row["A DRW/ECO E"])))
                # Grand Total
                can.drawString(447, height - 495, str(safe_num(row["A Hindi T"]) + safe_num(row["A English T"]) +
                                                  safe_num(row["A CIV T"]) + safe_num(row["A HIS/GEO T"]) +
                                                  safe_num(row["A DRW/ECO T"]) +
                                                  safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                                  safe_num(row["A CIV E"]) + safe_num(row["A HIS/GEO E"])+safe_num(row["A GEO PR"]) +
                                                  safe_num(row["A DRW/ECO E"])))


                # --- GRAND TOTAL (only if at least one Annual mark exists) ---
                annual_fields = [
                        "A Hindi T", "A Hindi E", "A English T", "A English E", 
                        "A CIV T", "A CIV E", "A HIS/GEO T", "A HIS/GEO E", 
                        "A DRW/ECO T", "A DRW/ECO E"
                ]
                if any(pd.notna(row.get(col, None)) for col in annual_fields):
                        can.drawString(525, height - 355, str(safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"]) +
                                                      safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))
                        can.drawString(525, height - 380, str(safe_num(row["A English T"]) + safe_num(row["A English E"]) +
                                                      safe_num(row["HY English T"]) + safe_num(row["HY English E"])))
                        can.drawString(525, height - 405, str(safe_num(row["A CIV T"]) + safe_num(row["A CIV E"]) +
                                                      safe_num(row["HY CIV T"]) + safe_num(row["HY CIV E"])))
                        can.drawString(525, height - 430, str(safe_num(row["A HIS/GEO T"]) + safe_num(row["A HIS/GEO E"]) +safe_num(row["HY GEO PR"])+
                                                      safe_num(row["HY HIS/GEO T"]) + safe_num(row["HY HIS/GEO E"])+safe_num(row["A GEO PR"])))
                        can.drawString(525, height - 452, str(safe_num(row["A DRW/ECO T"]) + safe_num(row["A DRW/ECO E"]) +
                                                      safe_num(row["HY DRW/ECO T"]) + safe_num(row["HY DRW/ECO E"])))
                
                        can.drawString(525, height - 495, str( 
                                                      safe_num(row["A DRW/ECO T"]) + safe_num(row["A DRW/ECO E"]) +
                                                      safe_num(row["HY DRW/ECO T"]) + safe_num(row["HY DRW/ECO E"])+
                                                      safe_num(row["A HIS/GEO T"]) + safe_num(row["A HIS/GEO E"]) +safe_num(row["A GEO PR"])+
                                                      safe_num(row["HY HIS/GEO T"]) + safe_num(row["HY HIS/GEO E"])+safe_num(row["HY GEO PR"])+
                                                      safe_num(row["A CIV T"]) + safe_num(row["A CIV E"]) +
                                                      safe_num(row["HY CIV T"]) + safe_num(row["HY CIV E"])+
                                                      safe_num(row["A English T"]) + safe_num(row["A English E"]) +
                                                      safe_num(row["HY English T"]) + safe_num(row["HY English E"])+
                                                      safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"]) +
                                                      safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))   
                                        
                
                can.save()
                packet.seek(0)
                new_pdf_reader = PdfReader(packet)
                overlay_page = new_pdf_reader.pages[0]
                page.merge_page(overlay_page)
                edited_page = page

        # ---------------------------------------
        # CLASS 6–8 SECTION
        # ---------------------------------------
        elif class_val in ['6', '7', '8', 'vi', 'vii', 'viii']:
            if page_num == 2:
                can.drawString(175, height - 123, str(safe_num(row["PAN"])))
                can.drawString(159, height - 174, safe_value(row["NAME OF STUDENTS"]))
                can.drawString(159, height - 190, safe_value(row["FATHERS NAME"]))
                can.drawString(440, height - 174, safe_value(row["MOTHERS NAME"]))
                can.drawString(470, height - 121, safe_value(row['Class']))
                can.drawString(175, height - 158, safe_value(row["S.R.N"]))
                can.drawString(474, height - 138, safe_value(row['Section']))
                can.drawString(440, height - 190, formatted_date)
                can.drawString(470, height - 155, safe_value(row["R.N"]))
                can.drawString(140, height - 206, safe_value(row["ADDRESS"]))
                
                # Half-Yearly Marks with format_mark
                can.drawString(200, height - 315, format_mark(row["HY Hindi T"]))
                can.drawString(230, height - 315, format_mark(row["HY Hindi E"]))
                can.drawString(200, height - 340, format_mark(row["HY English T"]))
                can.drawString(230, height - 340, format_mark(row["HY English E"]))
                can.drawString(200, height - 363, format_mark(row["HY Math T"]))
                can.drawString(230, height - 363, format_mark(row["HY Math E"]))
                can.drawString(200, height - 385, format_mark(row["HY SCI T"]))
                can.drawString(230, height - 385, format_mark(row["HY SCI E"]))
                can.drawString(200, height - 410, format_mark(row["HY SST T"]))
                can.drawString(230, height - 410, format_mark(row["HY SST E"]))
                can.drawString(200, height - 433, format_mark(row["HY COMP/COMM T"]))
                can.drawString(230, height - 433, format_mark(row["HY COMP/COMM E"]))
                can.drawString(200, height - 455, format_mark(row["HY DRG T"]))
                can.drawString(230, height - 455, format_mark(row["HY DRG E"]))
                can.drawString(200, height - 475, format_mark(row["HY SANS T"]))
                can.drawString(230, height - 475, format_mark(row["HY SANS E"]))
                
                # Half-Yearly Totals
                can.drawString(264, height - 315, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))
                can.drawString(264, height - 340, str(safe_num(row["HY English T"]) + safe_num(row["HY English E"])))
                can.drawString(264, height - 363, str(safe_num(row["HY Math T"]) + safe_num(row["HY Math E"])))
                can.drawString(264, height - 385, str(safe_num(row["HY SCI T"]) + safe_num(row["HY SCI E"])))
                can.drawString(264, height - 410, str(safe_num(row["HY SST T"]) + safe_num(row["HY SST E"])))
                can.drawString(264, height - 433, str(safe_num(row["HY COMP/COMM T"]) + safe_num(row["HY COMP/COMM E"])))
                can.drawString(264, height - 455, str(safe_num(row["HY DRG T"]) + safe_num(row["HY DRG E"])))
                can.drawString(264, height - 475, str(safe_num(row["HY SANS T"]) + safe_num(row["HY SANS E"])))
                can.drawString(200, height - 502, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY English T"]) +
                                     safe_num(row["HY Math T"]) + safe_num(row["HY SCI T"]) +
                                     safe_num(row["HY SST T"])+ safe_num(row["HY COMP/COMM T"])+ safe_num(row["HY DRG T"]) + safe_num(row["HY SANS T"])))
                can.drawString(230, height - 502, str(safe_num(row["HY Hindi E"]) + safe_num(row["HY English E"]) +
                                     safe_num(row["HY Math E"]) + safe_num(row["HY SCI E"]) +
                                     safe_num(row["HY SST E"])+ safe_num(row["HY COMP/COMM E"])+ safe_num(row["HY DRG E"]) + safe_num(row["HY SANS E"])))
                can.drawString(259, height - 502, str(safe_num(row["HY Hindi T"]) + safe_num(row["HY English T"]) +
                                     safe_num(row["HY Math T"]) + safe_num(row["HY SCI T"]) +
                                     safe_num(row["HY SST T"])+ safe_num(row["HY COMP/COMM T"])+ safe_num(row["HY DRG T"]) + safe_num(row["HY SANS T"])+safe_num(row["HY Hindi E"]) + safe_num(row["HY English E"]) +
                                     safe_num(row["HY Math E"]) + safe_num(row["HY SCI E"]) +
                                     safe_num(row["HY SST E"])+ safe_num(row["HY COMP/COMM E"])+ safe_num(row["HY DRG E"]) + safe_num(row["HY SANS E"]) ))
                
                # Annual Exam with format_mark
                can.drawString(364, height - 315, format_mark(row["A Hindi T"]))
                can.drawString(396, height - 315, format_mark(row["A Hindi E"]))
                can.drawString(364, height - 340, format_mark(row["A English T"]))
                can.drawString(396, height - 340, format_mark(row["A English E"]))
                can.drawString(364, height - 363, format_mark(row["A Math T"]))
                can.drawString(396, height - 363, format_mark(row["A Math E"]))
                can.drawString(364, height - 385, format_mark(row["A SCI T"]))
                can.drawString(396, height - 385, format_mark(row["A SCI E"]))
                can.drawString(364, height - 410, format_mark(row["A SST T"]))
                can.drawString(396, height - 410, format_mark(row["A SST E"]))
                can.drawString(364, height - 433, format_mark(row["A COMP/COMM T"]))
                can.drawString(396, height - 433, format_mark(row["A COMP/COMM E"]))
                can.drawString(364, height - 455, format_mark(row["A DRG T"]))
                can.drawString(396, height - 455, format_mark(row["A DRG E"]))
                can.drawString(364, height - 475, format_mark(row["A SANS T"]))
                can.drawString(396, height - 475, format_mark(row["A SANS E"]))
                
                # Annual Totals
                can.drawString(430, height - 315, str(safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"])))
                can.drawString(430, height - 340, str(safe_num(row["A English T"]) + safe_num(row["A English E"])))
                can.drawString(430, height - 363, str(safe_num(row["A Math T"]) + safe_num(row["A Math E"])))
                can.drawString(430, height - 385, str(safe_num(row["A SCI T"])  + safe_num(row["A SCI E"])))
                can.drawString(430, height - 410, str(safe_num(row["A SST T"]) + safe_num(row["A SST E"])))
                can.drawString(430, height - 433, str(safe_num(row["A COMP/COMM T"]) + safe_num(row["A COMP/COMM E"])))
                can.drawString(430, height - 455, str(safe_num(row["A DRG T"]) + safe_num(row["A DRG E"])))
                can.drawString(430, height - 475, str(safe_num(row["A SANS T"]) + safe_num(row["A SANS E"])))
                can.drawString(364, height - 502, str(safe_num(row["A Hindi T"]) + safe_num(row["A English T"]) +
                                     safe_num(row["A Math T"]) + safe_num(row["A SCI T"]) +
                                     safe_num(row["A SST T"])+ safe_num(row["A COMP/COMM T"])+ safe_num(row["A DRG T"]) + safe_num(row["A SANS T"])))
                can.drawString(396, height - 502, str(safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                     safe_num(row["A Math E"]) + safe_num(row["A SCI E"]) +
                                     safe_num(row["A SST E"])+ safe_num(row["A COMP/COMM E"])+ safe_num(row["A DRG E"]) + safe_num(row["A SANS E"])))
                can.drawString(426, height - 502, str( safe_num(row["A Hindi T"]) + safe_num(row["A English T"]) +
                                     safe_num(row["A Math T"]) + safe_num(row["A SCI T"]) +
                                     safe_num(row["A SST T"])+ safe_num(row["A COMP/COMM T"])+ safe_num(row["A DRG T"]) + safe_num(row["A SANS T"])+safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                     safe_num(row["A Math E"]) + safe_num(row["A SCI E"]) +
                                     safe_num(row["A SST E"])+ safe_num(row["A COMP/COMM E"])+ safe_num(row["A DRG E"]) + safe_num(row["A SANS E"])))
                
                # --- GRAND TOTAL (only if at least one Annual mark exists) ---
                annual_fields = [
                    "A Hindi T", "A Hindi E", "A English T", "A English E", "A Math T", "A Math E",
                    "A SCI T", "A SCI E", "A SST T", "A SST E"
                ]
                if any(pd.notna(row.get(col, None)) for col in annual_fields):
                        can.drawString(520, height - 315, str(safe_num(row["A Hindi T"]) + safe_num(row["A Hindi E"]) +
                                          safe_num(row["HY Hindi T"]) + safe_num(row["HY Hindi E"])))
                        can.drawString(520, height - 340, str(safe_num(row["A English T"]) + safe_num(row["A English E"]) +
                                          safe_num(row["HY English T"]) + safe_num(row["HY English E"])))
                        can.drawString(520, height - 363, str(safe_num(row["A Math T"]) + safe_num(row["A Math E"]) +
                                          safe_num(row["HY Math T"]) + safe_num(row["HY Math E"])))
                        can.drawString(520, height - 385, str(safe_num(row["A SCI T"]) +
                                          safe_num(row["A SCI E"]) + safe_num(row["HY SCI T"]) +
                                          safe_num(row["HY SCI E"])))
                        can.drawString(520, height - 410, str(safe_num(row["A SST T"])  +
                                          safe_num(row["A SST E"]) + safe_num(row["HY SST T"]) +
                                          safe_num(row["HY SST E"])))
                        can.drawString(520, height - 433, str(safe_num(row["A COMP/COMM T"])  +
                                          safe_num(row["A COMP/COMM E"]) + safe_num(row["HY COMP/COMM T"]) +
                                          safe_num(row["HY COMP/COMM E"])))
                        can.drawString(520, height - 455, str(safe_num(row["A DRG T"])  +
                                          safe_num(row["A DRG E"]) + safe_num(row["HY DRG T"]) +
                                          safe_num(row["HY DRG E"])))
                        can.drawString(520, height - 475, str(safe_num(row["A SANS T"])  +
                                          safe_num(row["A SANS E"]) + safe_num(row["HY SANS T"]) +
                                          safe_num(row["HY SANS E"])))
                        can.drawString(520, height - 502, str(
                                    safe_num(row["A Hindi T"]) + safe_num(row["A English T"]) +
                                    safe_num(row["A Math T"]) + safe_num(row["A SCI T"]) +
                                    safe_num(row["A SST T"])+ safe_num(row["A COMP/COMM T"])+ safe_num(row["A DRG T"]) + safe_num(row["A SANS T"])+safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                    safe_num(row["A Math E"]) + safe_num(row["A SCI E"]) +
                                    safe_num(row["A SST E"])+ safe_num(row["A COMP/COMM E"])+ safe_num(row["A DRG E"]) + safe_num(row["A SANS E"])+safe_num(row["HY Hindi T"]) + safe_num(row["HY English T"]) +
                                    safe_num(row["HY Math T"]) + safe_num(row["HY SCI T"]) +
                                    safe_num(row["HY SST T"])+ safe_num(row["HY COMP/COMM T"])+ safe_num(row["HY DRG T"]) + safe_num(row["HY SANS T"])+safe_num(row["HY Hindi E"]) + safe_num(row["HY English E"]) +
                                    safe_num(row["HY Math E"]) + safe_num(row["HY SCI E"]) +
                                    safe_num(row["HY SST E"])+ safe_num(row["HY COMP/COMM E"])+ safe_num(row["HY DRG E"]) + safe_num(row["HY SANS E"])))

                can.save()
                packet.seek(0)
                new_pdf_reader = PdfReader(packet)
                overlay_page = new_pdf_reader.pages[0]
                page.merge_page(overlay_page)
                edited_page = page
                
        elif class_val in ['9','10','ix','x'] :
            if page_num == 3 :
                
                can.drawString(159, height - 117, str(safe_num(row["PAN"])))
                can.drawString(159, height - 190, safe_value(row["NAME OF STUDENTS"]))
                can.drawString(159, height - 223, safe_value(row["FATHERS NAME"]))
                can.drawString(440, height - 190, safe_value(row["MOTHERS NAME"]))
                can.drawString(445, height - 117, safe_value(row['Class']))
                can.drawString(175, height - 153, safe_value(row["S.R.N"]))
                can.drawString(455, height - 135, safe_value(row['Section']))
                can.drawString(440, height - 223, formatted_date)
                can.drawString(470, height - 153, safe_value(row["R.N"]))
                
                # Half-Yearly with format_mark
                can.drawString(220, height - 347, format_mark(row["HY Hindi"]))
                can.drawString(220, height - 370, format_mark(row["HY English"]))
                can.drawString(220, height - 392, format_mark(row["HY Math"]))
                can.drawString(220, height - 413, format_mark(row["HY SCI"]))
                can.drawString(220, height - 435, format_mark(row["HY SST"]))
                can.drawString(220, height - 459, format_mark(row["HY DRG/COMM"]))
                
                can.drawString(220, height - 502, str(safe_num(row["HY DRG/COMM"]) + safe_num(row["HY SST"]) +  safe_num(row["HY SCI"]) + safe_num(row["HY Math"]) + safe_num(row["HY English"])+ safe_num(row["HY Hindi"])))
                
                # Annual with format_mark
                can.drawString(325, height - 347, format_mark(row["A Hindi E"]))
                can.drawString(365, height - 347, format_mark(row["A Hindi PR"]))
                can.drawString(325, height - 370, format_mark(row["A English E"]))
                can.drawString(365, height - 370, format_mark(row["A English PR"]))
                can.drawString(325, height - 392, format_mark(row["A Math E"]))
                can.drawString(365, height - 392, format_mark(row["A Math PR"]))
                can.drawString(325, height - 413, format_mark(row["A SCI E"]))
                can.drawString(365, height - 413, format_mark(row["A SCI PR"]))
                can.drawString(325, height - 435, format_mark(row["A SST E"]))
                can.drawString(365, height - 435, format_mark(row["A SST PR"]))
                can.drawString(325, height - 459, format_mark(row["A DRG/COMM E"]))
                can.drawString(365, height - 459, format_mark(row["A DRG/COMM PR"]))

                # Annual Totals
                can.drawString(400, height - 347, str(safe_num(row["A Hindi PR"]) + safe_num(row["A Hindi E"])))
                can.drawString(400, height - 370, str(safe_num(row["A English PR"]) + safe_num(row["A English E"])))
                can.drawString(400, height - 392, str(safe_num(row["A Math PR"]) + safe_num(row["A Math E"])))
                can.drawString(400, height - 413, str(safe_num(row["A SCI PR"])  + safe_num(row["A SCI E"])))
                can.drawString(400, height - 435, str(safe_num(row["A SST PR"]) + safe_num(row["A SST E"])))
                can.drawString(400, height - 459, str(safe_num(row["A DRG/COMM PR"]) + safe_num(row["A DRG/COMM E"])))
                can.drawString(365, height - 502, str(safe_num(row["A Hindi PR"]) + safe_num(row["A English PR"]) +
                                     safe_num(row["A Math PR"]) + safe_num(row["A SCI PR"]) +
                                     safe_num(row["A SST PR"])+ safe_num(row["A DRG/COMM PR"]) ))
                can.drawString(325, height - 502, str(safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                     safe_num(row["A Math E"]) + safe_num(row["A SCI E"]) +
                                     safe_num(row["A SST E"])+ safe_num(row["A DRG/COMM E"]) ))
                can.drawString(400, height - 502, str( safe_num(row["A Hindi PR"]) + safe_num(row["A English PR"]) +
                                     safe_num(row["A Math PR"]) + safe_num(row["A SCI PR"]) +
                                     safe_num(row["A SST PR"])+  safe_num(row["A DRG/COMM PR"]) +safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                     safe_num(row["A Math E"]) + safe_num(row["A SCI E"]) +
                                     safe_num(row["A SST E"])+ safe_num(row["A DRG/COMM E"])))
                
                annual_fields = [
                    "A Hindi PR", "A Hindi E", "A English PR", "A English E", "A Math PR", "A Math E",
                    "A SCI PR", "A SCI E", "A SST PR", "A SST E","A DRG/COMM PR","A DRG/COMM E"
                ]
                if any(pd.notna(row.get(col, None)) for col in annual_fields):
                        can.drawString(505, height - 347, str(safe_num(row["A Hindi PR"]) + safe_num(row["A Hindi E"]) +
                                          safe_num(row["HY Hindi"])))
                        can.drawString(505, height - 370, str(safe_num(row["A English PR"]) + safe_num(row["A English E"]) +
                                           safe_num(row["HY English"])))
                        can.drawString(505, height - 392, str(safe_num(row["A Math PR"]) + safe_num(row["A Math E"]) +
                                          safe_num(row["HY Math"]) ))
                        can.drawString(505, height - 413, str(safe_num(row["A SCI PR"]) +
                                          safe_num(row["A SCI E"]) + safe_num(row["HY SCI"]) ))
                        can.drawString(505, height - 435, str(safe_num(row["A SST PR"])  +
                                          safe_num(row["A SST E"]) + safe_num(row["HY SST"]) ))
                        can.drawString(505, height - 459, str(safe_num(row["A DRG/COMM PR"])  +
                                          safe_num(row["A DRG/COMM E"]) + safe_num(row["HY DRG/COMM"]) ))
                        
                        can.drawString(505, height - 502, str(
                                    safe_num(row["A Hindi PR"]) + safe_num(row["A English PR"]) +
                                    safe_num(row["A Math PR"]) + safe_num(row["A SCI PR"]) +
                                    safe_num(row["A SST PR"])+ safe_num(row["A DRG/COMM PR"]) + safe_num(row["A Hindi E"]) + safe_num(row["A English E"]) +
                                    safe_num(row["A Math E"]) + safe_num(row["A SCI E"]) +
                                    safe_num(row["A SST E"])+  safe_num(row["A DRG/COMM E"]) +safe_num(row["HY Hindi"]) + safe_num(row["HY English"]) +
                                    safe_num(row["HY Math"]) + safe_num(row["HY SCI"]) +
                                    safe_num(row["HY SST"])+ safe_num(row["HY DRG/COMM"]) ))                              
            
            
                can.save()
                packet.seek(0)
                new_pdf_reader = PdfReader(packet)
                overlay_page = new_pdf_reader.pages[0]
                page.merge_page(overlay_page)
                edited_page = page
            
            

    # ✅ Only add the edited page
    if edited_page:
        output_pdf_writer.add_page(edited_page)

    # --- Save Report ---
    output_filename = f"reports/{safe_value(row['NAME OF STUDENTS'])}_report{index+1}.pdf"
    with open(output_filename, "wb") as f:
        output_pdf_writer.write(f)

    # --- Append to merger ---
    merger.append(output_filename)

    print(f"Generated: {output_filename}")

# --- Write merged output ---
merged_output = "All_Students_Report.pdf"
merger.write(merged_output)
merger.close()

print(f"\n✅ Report cards generated successfully!")
print(f"✅ All reports merged into '{merged_output}'")