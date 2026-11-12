"""Builds the fake test files in this folder (CSV + Excel) from the same 18 students as
Regenerate.gs "Reset + Load Fake Data". Run: python3 make_test_data.py  (needs openpyxl).
python3 make_test_data.py --zip-only  rebuilds only "PC Tracker test data.zip" from the files here."""
import csv, io, os, re, json, sys, zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ZIP_NAME = "PC Tracker test data.zip"

# The download zip has ONE file per data set, in the format the real export comes in
# (owner 2026-11-12): Student Data (ATS) and Parent Square = .csv, NYCSA = .xlsx. Site / Class
# is typed or pasted, so .csv; Your Lists = .xlsx (lists usually come as spreadsheets, and the
# bus-routes one has title rows above its headers). This folder keeps both formats: the tests use them.
ZIP_FORMAT = {"1": ".csv", "2": ".csv", "3": ".xlsx", "4": ".csv", "5": ".xlsx"}

def build_zip():
    files = [f for f in sorted(os.listdir(HERE))
             if f.endswith((".csv", ".xlsx")) and os.path.splitext(f)[1] == ZIP_FORMAT.get(f[0])
             or f.startswith("1c ")]  # the broken ATS file only exists as .csv
    with zipfile.ZipFile(os.path.join(HERE, ZIP_NAME), "w", zipfile.ZIP_DEFLATED) as z:
        for f in files + ["README.md"]:
            z.write(os.path.join(HERE, f), f)
    return files

if "--zip-only" in sys.argv:
    print("zip:", ", ".join(build_zip()))
    sys.exit(0)

from openpyxl import Workbook
from openpyxl.styles import Font
REGEN = open(os.path.join(HERE, "..", "Regenerate.gs"), encoding="utf-8").read()

def rows_of(fn):
    """The data rows written by one loadFake*_ function in Regenerate.gs."""
    body = REGEN[REGEN.index("function " + fn + "(") :]
    body = body[: body.index("\n}\n")]
    return [json.loads(l.strip().rstrip(",")) for l in body.split("\n") if re.match(r"^\s*\[", l) and l.strip().startswith("[\"") or re.match(r'^\s*\["', l)]

ATS_H = ["School Year","DBN","Student ID","Last Name","First Name","Race","Gender","Age","IEP Flag","ELL Flag","504 Flag","NYSAA Eligible","Housing Status","Temporary Housing Flag","Economic Disadvantage Flag","Cohort Year","Grade","Official Class","Home Language","Guardian Name 1","Guardian Relationship 1","Guardian Phone Number 1","Guardian Email 1","Guardian Spoken Language 1","Guardian Written Language 1","Guardian Name 2","Guardian Relationship 2","Guardian Phone Number 2","Guardian Email 2","Guardian Spoken Language 2","Guardian Written Language 2"]
PS_H = ["Name","Students","Email","Phone","Secondary Phone","Record Created","Registered?","SIS Synced"]
NY_H = ["Location Code","School District","School Name","Class Code","Grade Level","Student ID","Student First Name","Student Last Name","Guardian First Name","Guardian Last Name","Another Custodial NYCSA Account Exist","Guardian Home Number","Guardian Work Number","Guardian Cell Number","Relationship","Role","Account Creation Code"]

ats = rows_of("loadFakeATS_")
ps = rows_of("loadFakeParentsSquare_")
ny = rows_of("loadFakeNYCSA_")
assert len(ats) == 18 and len(ps) >= 20 and len(ny) >= 20, (len(ats), len(ps), len(ny))
ats_num = lambda r: [int(v) if isinstance(v, str) and re.fullmatch(r"\d{9,10}", v) else v for v in r]  # Excel: IDs / phones as numbers

def write(name, header, rows, title_rows=None, numbers=False):
    """name.csv (UTF-8, no title rows) and name.xlsx (optional title rows above the header)."""
    with open(os.path.join(HERE, name + ".csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(header); w.writerows(rows)
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    for t in (title_rows or []): ws.append(t)
    ws.append(header)
    for c in ws[ws.max_row]: c.font = Font(bold=True)
    for r in rows: ws.append(ats_num(r) if numbers else r)
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 45)
    wb.save(os.path.join(HERE, name + ".xlsx"))

# 1. Student Data (ATS / Insight "Guardian Contact Information")
write("1 - Student Data (ATS)", ATS_H, ats, [["Guardian Contact Information - 75M094 - exported 11/10/2026"], [""]], numbers=True)

# 1b. Next month: Tony Stark left, 2 new students (tests the build report's new / discharged)
nxt = [r for r in ats if r[2] != "923456789"] + [
  ["2025-2026","75M094","812345670","ROGERS","STEVE","NA","M",11,1,0,0,0,"P - Permanent",0,1,"",5,"Y24","English","SARAH ROGERS","MOTHER","2424242424","SARAH.ROGERS@gmail.com","English","English","","","","","",""],
  ["2025-2026","75M094","823456781","LOPEZ","LUCIA","NA","F",8,1,1,0,1,"P - Permanent",0,1,"",2,"Y41","Spanish","CARMEN LOPEZ","MOTHER","2525252525","","Spanish","Spanish","LUIS LOPEZ","FATHER","2626262626","LUIS.LOPEZ@gmail.com","Spanish","Spanish"]]
write("1b - Student Data (ATS) - next month (2 new, 1 left)", ATS_H, nxt, numbers=True)

# 1c. Broken copy: Student ID and phones cut off by a narrow Excel column (Input Data Mode should warn)
def cut(v):
    return ("%.2E" % int(v)).replace("E+0", "E+") if isinstance(v, str) and re.fullmatch(r"\d{9,10}", v) else v
with open(os.path.join(HERE, "1c - Student Data (ATS) - BROKEN numbers cut off.csv"), "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f); w.writerow(ATS_H); w.writerows([[cut(v) if i in (2, 21, 27) else v for i, v in enumerate(r)] for r in ats])

# 2. Parent Square (Admin -> Data Assistant -> Parents -> Export CSV)
write("2 - Parent Square", PS_H, ps)

# 3. NYCSA ("Students with Non-NYCSA Account")
write("3 - NYCSA (Students with Non-NYCSA Account)", NY_H, ny)

# 4. Site / Class: paste over the "Site / Class" tab in Input Data Mode step 4 (columns A:C)
site_class = [["Y11","Bam","276"],["Y12","","276"],["Y24","Lam","281"],["Y25","Sam","51"],["Y27","Rosa","51"],["Y28","Ortiz","281"],
              ["Y41","Tessa","15"],["Y42","Nina","15"],["Y50","","340"],["Y66","Carl","51"],["X30","Ham","188"],["X31","Lee","188"]]
write("4 - Site and Class", ["Class","Teacher Name","Site (Optional)"], site_class)

# 5. Your Lists. Each is built to show the matching: exact, OSIS, typos, other name order,
#    short names, initials, siblings, a parent's name, unknown people, blank cells.
write("5a - Your Lists - Bus Routes (student list)", ["Student","OSIS","Route","Stop","AM Pickup"], [
  ["Bart Simpson","123456789","R12","Evergreen Terr & 7th","7:05"],
  ["SIMPSON, LISA","","R12","Evergreen Terr & 7th","7:05"],          # LAST, FIRST
  ["Wednesday Addams","234567891","R7","Cemetery Ln","6:50"],
  ["Pugsly Adams","","R7","Cemetery Ln","6:50"],                     # typos: best guess
  ["Dash Incredible","","R3","Elm St","7:20"],
  ["Violet Incredable","","R3","Elm St","7:20"],                     # typo
  ["Peter Parker","678912345","","",""],                             # no route (blank cells)
  ["MORALES MILES","","R9","Brooklyn Ave","6:40"],                   # last name first, no comma
  ["Clark Kent","891234567","R15","Hob's End","7:30"],
  ["Diana P.","","R4","Themyscira Pl","7:10"],                       # last initial
  ["Lou Belcher","","R2","Ocean Ave","7:00"],                        # short name (Louise)
  ["Gene Belcher","967891234","R2","Ocean Ave","7:00"],
  ["Kim Possible","978912345","R15","Middleton Rd","7:30"],
  ["Mulan Fa","989123456","R5","Imperial Way","6:55"],
  ["Mirabel Madrigal","999999999","R6","Encanto Dr","7:15"],         # OSIS not on the Student List
  ["Bruce B.","","R9","Brooklyn Ave","6:40"],                        # two Bruces: initial says Banner
  ["Steve Rogers","","R9","Brooklyn Ave","6:40"],                    # not on the Student List (until 1b)
  ["Nobody Here","","R1","Main St","7:00"]                           # not found
], [["District 75 Transportation - Fall Bus Routes"], ["Printed 11/03/2026"], [""]])

write("5b - Your Lists - PTA Members (guardian list)", ["Parent First","Parent Last","Email","Phone","Child","Role"], [
  ["Homer","Simpson","HOMERSIMPSON@gmail.com","(111) 111-1111","Bart","Member"],
  ["Marj","Simpson","","(222) 222-2222","Lisa","Volunteer"],          # typo + phone
  ["Morticia","Adams","MORTICIA.ADDAMS@gmail.com","","Wednesday","Board"],  # typo, email agrees
  ["Linda","Belcher","LINDA.BELCHER@gmail.com","","Louise, Gene","Board"],
  ["Bob","Belcher","","171-717-1717","Gene","Member"],
  ["Julieta","Madrigal","JULIETA.MADRIGAL@gmail.com","","Mirabel","Volunteer"],
  ["Jefferson","Morales","","(888) 888-8888","Miles","Member"],
  ["Martha","K.","","","Clark","Member"],                             # last initial
  ["Ann","Possible","ANN.POSSIBLE@gmail.com","","Kim","Volunteer"],
  ["Pepper","Potts","PEPPER@STARK.COM","","Tony","Board"],
  ["Gomez","Addams","","","Pugsley","Member"],
  ["Alfred","Pennywise","","(121) 212-1212","Bruce","Volunteer"],     # typo, phone agrees
  ["Natasha","Romanoff","","(555) 000-1234","","Member"]              # not found
])

write("5c - Your Lists - Field Trip Forms (first + last name columns)", ["First Name","Last Name","Class","Form Returned","Paid"], [
  ["Bart","Simpson","Y41","Yes","$10"],
  ["Lisa","Simpson","Y25","Yes",""],
  ["Wednesday","Addams","Y66","",""],                                 # blank: "If blank, write Not returned"
  ["Pugsley","Addams","Y66","Yes","$10"],
  ["Miles","Morales","Y24","Yes","$10"],
  ["Bruce","Wayne","Y50","Yes","$10"],
  ["Bruce","Banner","Y24","",""],
  ["Kim","Possible","X31","Yes","$10"],
  ["Mirabel","Madrigal","Y28","Yes",""],
  ["Tina","Belcher","Y42","Yes",""],                                  # not in ATS (a New/Edit Student override in the fake data)
])

write("5d - Your Lists - After School Sign-up (LAST, FIRST)", ["Name","Grade","Program","Days"], [
  ["SIMPSON, BART","6","Robotics","Mon/Wed"],
  ["BELCHER, LOUISE","3","Art","Tue/Thu"],
  ["BELCHER, EUGENE","5","Music","Tue/Thu"],                           # full name for Gene: best guess
  ["KENT, CLARK","6","Basketball","Mon-Fri"],
  ["STARK, ANTHONY","6","Robotics","Mon/Wed"],                         # Tony = Anthony (nickname)
  ["PRINCE, DIANA","4","Dance","Fri"],
  ["MADRIGAL, MIRABEL","5","Choir","Wed"],
  ["FA, MULAN","4","Dance","Fri"]
])

# The download zip: one file per data set (ZIP_FORMAT), README included
print("zip:", ", ".join(build_zip()))
print("done:", len([f for f in os.listdir(HERE) if f.endswith((".csv", ".xlsx"))]), "files")
