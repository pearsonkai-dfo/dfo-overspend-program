import os
import glob
import openpyxl
import re
import warnings
import unicodedata

from decimal import Decimal

from overspendclasses import Cardholder
from overspendclasses import CostCenter
from overspendclasses import Purchase
from overspendclasses import PurchaseStrand

##### GENERAL CONSTANTS #####

GL_NUMBER_LENGTH = 8
CC_NUMBER_LENGTH = 6



##### SAP EXPORT SHEET CONSTANTS #####

SAP_DOCUMENT_NUMBER_COLUMN_IDENTIFIER = "Document Number"
SAP_CARDHOLDER_COLUMN_IDENTIFIER = "Cardholder Name"
SAP_COST_CENTER_COLUMN_IDENTIFIER = "Cost Center"
SAP_GL_COLUMN_IDENTIFIER = "G/L Account"
SAP_POSTED_AMOUNT_COLUMN_IDENTIFIER = "Transaction Amount"
SAP_MATCHING_AMOUNT_COLUMN_IDENTIFIER = "Matching Amount"

##### BLANKET AUTHORITY TRACKING SHEET CONSTANTS #####

BA_CARDHOLDER_COLUMN_IDENTIFIER = "Cardholder Name" # Column name for the cardholders
BA_LIMIT_COLUMN_IDENTIFIER = "BA Limit" # Column name for the budget limit
BA_CC_COLUMN_IDENTIFIER = "Cost Center" # Column name for the cost centers
BA_GL_COLUMN_IDENTIFIER = "Good or Service" # Column name for the GL Accounts
BA_SUBSEQUENT_COLUMN_IDENTIFIER = "Subsequent Blanket Authority" # Column name for identifying if the cardholder already has a BA with the cost center

ACCOUNT_NUM_IDENTIFIER = "GL Account Code: "  # The text in the BA Tracking Sheet up until the GL Account Code








##### PROGRESS BAR FUNCTIONS #####

def remove_progress_bar():
    print("\r" + " " * 100 + "\r", flush=True, end="")


def place_progress_bar(bar_name: str, bar_segments: int, items_done: int, items_total: int):
    percentage = items_done/items_total
    segments_filled = round(percentage * bar_segments)
    
    bar = "["
    for i in range(segments_filled):
        bar += "#"

    for i in range(bar_segments - segments_filled):
        bar += "-"

    bar += "] (" + str(items_done) + "/" + str(items_total) + " Items) (" + str(round(percentage*100,3)) + "%)"
    print(bar,flush=True,end="")

    
##### FUNCTIONS START #####


def normalize_name(name):
    if not isinstance(name,str) and name is not None:
        return None
    
    name = name.lower()

    # Remove accents
    name = ''.join(
        c for c in unicodedata.normalize('NFD', name)
        if unicodedata.category(c) != 'Mn'
    )

    # Remove apostrophes
    name = name.replace("'", "")

    # Replace separators with spaces
    name = re.sub(r"[-,_/]", " ", name)

    # Remove other junk
    name = re.sub(r"[^a-z\s]", "", name)

    # Collapse spaces
    return " ".join(name.split())

def cheap_match(target, holder):
    t = target.lower().split()
    h = holder.lower().split()

    if len(t) < 2 or len(h) < 2:
        return False

    t_first = t[0]
    h_first = h[0]

    t_last = t[-1]
    h_last = h[-1]

    return (
        t_first[0] == h_first[0]
        and (
            h_last.startswith(t_last)
            or t_last.startswith(h_last)
        )
    )

def medium_match(target, holder):
    t = target.lower().split()
    h = holder.lower().split()

    if len(t) < 2 or len(h) < 2:
        return False

    t_first = t[0]
    h_first = h[0]

    t_last = t[-1]
    h_last = h[-1]

    return (
        (
            h_first.startswith(t_first)
            or t_first.startswith(h_first)
        )
        and (
            h_last.startswith(t_last)
            or t_last.startswith(h_last)
        )
    )

def name_score(target, holder):
    t = target.lower().split()
    h = holder.lower().split()

    if len(t) < 2 or len(h) < 2:
        return 0

    t_first = t[0]
    h_first = h[0]

    t_last = t[-1]
    h_last = h[-1]

    # hard rejection
    if t[0] != h[0]:
        return 0

    if t[0] != h[0]:
        return 0

    score = 0

    # first name
    if t_first == h_first:
        score += 30
    elif (
        t_first.startswith(h_first)
        or h_first.startswith(t_first)
    ):
        score += 20

    # last name
    if t_last == h_last:
        score += 50
    elif (
        t_last.startswith(h_last)
        or h_last.startswith(t_last)
    ):
        score += 35

    # middle names
    t_middle = t[1:-1]
    h_middle = h[1:-1]

    for tm in t_middle:
        for hm in h_middle:
            if (
                hm.startswith(tm)
                or tm.startswith(hm)
            ):
                score += 10
                break

    return score

def find_best_match(target, libraryCardholder):
    if not isinstance(target,str) and target is not None:
        return None
    
    best_match = None
    best_score = -1

    for holder in libraryCardholder:

        # Fast reject
        if not cheap_match(target, holder):
            continue

        # Stronger filter
        if not medium_match(target, holder):
            continue

        # Expensive scoring
        score = name_score(target, holder)

        if score > best_score:
            best_score = score
            best_match = holder

    return best_match


def searchSimilarNames(target, libraryCardholder):
    target_parts = target.lower().split()

    for holder in libraryCardholder:
        holder_parts = holder.lower().split()

        if len(target_parts) == len(holder_parts):
            match = True

            for t, h in zip(target_parts, holder_parts):
                if not h.startswith(t):
                    match = False
                    break

            if match:
                return holder

    return target





def readBASheet(path,libraryCardholder):
    print("Loading workbook.", flush=True)
    BAWorkbook = openpyxl.load_workbook(path)
    print("Workbook loaded.", flush=True)
    BAWorksheet = BAWorkbook.active

    cardholderColumn = None
    limitColumn = None
    costCenterColumn = None
    GLColumn = None
    SubsequentBAColumn = None

    for column in BAWorksheet.iter_cols():
        if column[0].value == BA_CARDHOLDER_COLUMN_IDENTIFIER:
            cardholderColumn = column[0].column - 1
        
        elif column[0].value == BA_LIMIT_COLUMN_IDENTIFIER:
            limitColumn = column[0].column - 1
        
        elif column[0].value == BA_CC_COLUMN_IDENTIFIER:
            costCenterColumn = column[0].column - 1
            
        elif column[0].value == BA_GL_COLUMN_IDENTIFIER:
            GLColumn = column[0].column - 1
        
        elif column[0].value == BA_SUBSEQUENT_COLUMN_IDENTIFIER:
            SubsequentBAColumn = column[0].column - 1
    
    if None in (cardholderColumn, limitColumn, costCenterColumn, GLColumn, SubsequentBAColumn):
        raise Exception("Could not locate all required columns")
        return

    for row in BAWorksheet.iter_rows(min_row=2):
        name = row[cardholderColumn].value
        limit = row[limitColumn].value
        costCenterNum = int(row[costCenterColumn].value)
        subsequentBA = row[SubsequentBAColumn].value == "Yes"
        
        textGL = row[GLColumn].value
        if textGL is None:
            continue

        
        if name not in libraryCardholder:
            holder = Cardholder(name)
            libraryCardholder[normalize_name(name)] = holder
        else:
            holder = libraryCardholder[name]


        costCenter = CostCenter(costCenterNum,limit)
        holder.add_cost_center(costCenter)

        while True:
            pos = textGL.find(ACCOUNT_NUM_IDENTIFIER)

            if pos == -1:
                break

            refGL = textGL[pos + len(ACCOUNT_NUM_IDENTIFIER) : pos + len(ACCOUNT_NUM_IDENTIFIER) + GL_NUMBER_LENGTH] # grab 8 digit GL code
            textGL = textGL[pos + len(ACCOUNT_NUM_IDENTIFIER) + GL_NUMBER_LENGTH :] # delete text up until the next ACCOUNT_NUM_IDENTIFIER

            costCenter.add_approved_gl(refGL)



def readSAPExport(path,libraryCardholder,libraryDocument):
    print("Loading workbook.", flush=True)
    SAPWorkbook = openpyxl.load_workbook(path)
    print("Workbook loaded.", flush=True)
    SAPWorksheet = SAPWorkbook.active

    numPeople = 0
    numKnown = 0
    barUpdatesFreq = max(50,round(SAPWorksheet.max_row/38))

    documentColumn = None
    cardholderColumn = None
    costCenterColumn = None
    GLColumn = None
    postedAmoungColumn = None
    matchingAmountColumn = None

    for column in SAPWorksheet.iter_cols():
        if column[0].value == SAP_DOCUMENT_NUMBER_COLUMN_IDENTIFIER:
            documentColumn = column[0].column - 1

        elif column[0].value == SAP_CARDHOLDER_COLUMN_IDENTIFIER:
            cardholderColumn = column[0].column - 1

        elif column[0].value == SAP_COST_CENTER_COLUMN_IDENTIFIER:
            costCenterColumn = column[0].column - 1
            
        elif column[0].value == SAP_GL_COLUMN_IDENTIFIER:
            GLColumn = column[0].column - 1

        elif column[0].value == SAP_POSTED_AMOUNT_COLUMN_IDENTIFIER:
            postedAmoungColumn = column[0].column - 1

        elif column[0].value == SAP_MATCHING_AMOUNT_COLUMN_IDENTIFIER:
            matchingAmountColumn = column[0].column - 1
    
    for row in SAPWorksheet.iter_rows(min_row=2):
        documentNum = row[documentColumn].value
        cardholderName = row[cardholderColumn].value
        costCenter = row[costCenterColumn].value
        matchingAmount = row[matchingAmountColumn].value
        refGL = row[GLColumn].value

        if any(x is None for x in (documentNum,refGL,matchingAmount,costCenter,cardholderName)) is None or costCenter == "#N/A":
            continue
        if cardholderName == None:
            continue
        refGL = refGL[:GL_NUMBER_LENGTH]
        matchingAmount = int(round(float(matchingAmount)*100))
        costCenter = costCenter[:CC_NUMBER_LENGTH]
        if not isinstance(documentNum,int):
            documentNum = int(documentNum)
        numPeople = numPeople + 1

        

        #suggestedName = "friend"
        #suggestedName = searchSimilarNames(cardholderName.lower(),libraryCardholder)
        suggestedName = find_best_match(normalize_name(cardholderName),libraryCardholder)
        
        if cardholderName.lower() in libraryCardholder or (suggestedName != None and suggestedName in libraryCardholder):
            numKnown = numKnown + 1
        else:
            # flag for cardholder not in BA Tracking
            #print(cardholderName)
            continue

        if row[0].row % barUpdatesFreq == 0:
            remove_progress_bar()
            place_progress_bar("SAP Extract",20,numPeople,SAPWorksheet.max_row)
    
    remove_progress_bar()
    place_progress_bar("SAP Extract",20,SAPWorksheet.max_row,SAPWorksheet.max_row)
    print("\n")
    print("Hit Rate: " + str(round(numKnown*100/numPeople,3)) + "%")
    

        




def main():
    warnings.filterwarnings(
        "ignore",
        category=UserWarning,
        module="openpyxl"
    )

    libraryCardholder = {}
    libraryDocument = {}


    print("Enter Blanket Authority Excel Sheet Directory:")
    path = input()
    readBASheet(path,libraryCardholder)
    
    print("Enter QA Transactions (SAP Export) Excel Sheet Folder Directory:")
    path = input()
    readSAPExport(path,libraryCardholder,libraryDocument)
    





    """
    print("Enter QA Transactions (SAP Export) Excel Sheet Folder Directory:")
    path = input()
    readSAPExport(path,libraryGL,libraryCardholder)"""
 

    """
    script_dir = os.path.dirname(os.path.abspath(__file__))

    for filename in os.listdir(script_dir):
        if filename.endswith(".xlsx"):
            
            
            full_path = os.path.join(script_dir, filename)
            print(full_path)
    


    people = ["Monique Bakker","Piper-Lynn"]
    pipersCenters = libraryCardholder[people[0]].get_all_cost_center()
    for center in pipersCenters:
        print(center.get_number())
        for refGL in center.get_all_gl():
            print("   " + center.get_all_gl()[refGL].get_number())
    """

    

main()
input()