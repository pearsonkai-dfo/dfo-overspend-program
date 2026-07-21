import openpyxl
import warnings
from pathlib import Path 
import math

#from decimal import Decimal

from overspendclasses import Card
from overspendclasses import CostCenter
from overspendclasses import Purchase
from overspendclasses import PurchaseStrand
from overspendclasses import Flag
from overspendclasses import flag_codes

##### GENERAL CONSTANTS #####

GL_NUMBER_LENGTH = 8
CC_NUMBER_LENGTH = 6



##### SAP EXPORT SHEET CONSTANTS #####

SAP_DOCUMENT_NUMBER_COLUMN_IDENTIFIER = "Document Number"
SAP_CARD_COLUMN_IDENTIFIER = "Card Number"
SAP_COST_CENTER_COLUMN_IDENTIFIER = "Cost Center"
SAP_GL_COLUMN_IDENTIFIER = "G/L Account"
SAP_POSTED_AMOUNT_COLUMN_IDENTIFIER = "Transaction Amount"
SAP_MATCHING_AMOUNT_COLUMN_IDENTIFIER = "Matching Amount"

##### BLANKET AUTHORITY TRACKING SHEET CONSTANTS #####

BA_CARD_COLUMN_IDENTIFIER = "Card Number" # Column name for the card number
BA_LOWER_LIMIT_COLUMN_IDENTIFIER = "Lower Limit Request" # Column name for the budget limit
BA_HIGHER_LIMIT_COLUMN_IDENTIFIER = "Higher limit" # Column name for the budget limit
BA_CC_COLUMN_IDENTIFIER = "Cost Center" # Column name for the cost centers
BA_GL_COLUMN_IDENTIFIER = "Good or Service" # Column name for the GL Accounts
BA_SUBSEQUENT_COLUMN_IDENTIFIER = "Subsequent Blanket Authority" # Column name for identifying if the cardholder already has a BA with the cost center

ACCOUNT_NUM_IDENTIFIER = "GL Account Code: "  # The text in the BA Tracking Sheet up until the GL Account Code








##### PROGRESS BAR FUNCTIONS #####

def remove_progress_bar():
    print("\r" + " " * 100 + "\r", flush=True, end="")


def place_progress_bar(bar_name: str, bar_segments: int, items_done: int, items_total: int):
    bar = ""
    if bar_name is not None:
        bar = bar_name + " "
    percentage = items_done/items_total
    segments_filled = round(percentage * bar_segments)
    
    bar += "["
    for i in range(segments_filled):
        bar += "#"

    for i in range(bar_segments - segments_filled):
        bar += "-"

    bar += "] (" + str(items_done) + "/" + str(items_total) + " Items) (" + str(round(percentage*100,3)) + "%)"
    print(bar,flush=True,end="")

    
##### FUNCTIONS START #####

def convert(cents):
    return f"{cents // 100}.{cents % 100:02d}"


def checkNone(set):
    return None in set


def readBASheet(path,libraryCard,libraryFlag):
    print("Loading workbook \"" + path + "\".", flush=True)
    BAWorkbook = openpyxl.load_workbook(path)
    print("Workbook loaded.", flush=True)
    BAWorksheet = BAWorkbook.active

    cardColumn = None
    lowerLimitColumn = None
    higherLimitColumn = None
    costCenterColumn = None
    GLColumn = None
    SubsequentBAColumn = None

    for column in BAWorksheet.iter_cols():
        if column[0].value == BA_CARD_COLUMN_IDENTIFIER:
            cardColumn = column[0].column - 1
        
        elif column[0].value == BA_LOWER_LIMIT_COLUMN_IDENTIFIER:
            lowerLimitColumn = column[0].column - 1
        
        elif column[0].value == BA_HIGHER_LIMIT_COLUMN_IDENTIFIER:
            higherLimitColumn = column[0].column - 1
        
        elif column[0].value == BA_CC_COLUMN_IDENTIFIER:
            costCenterColumn = column[0].column - 1
            
        elif column[0].value == BA_GL_COLUMN_IDENTIFIER:
            GLColumn = column[0].column - 1
        
        elif column[0].value == BA_SUBSEQUENT_COLUMN_IDENTIFIER:
            SubsequentBAColumn = column[0].column - 1
    
    if None in (cardColumn,lowerLimitColumn, higherLimitColumn, costCenterColumn, GLColumn, SubsequentBAColumn):
        print("Could not locate all required columns. Check column identifiers in Python File.")
        return

    for row in BAWorksheet.iter_rows(min_row=2):
        cardNum = row[cardColumn].value
        limit = 20000
        costCenterNum = int(row[costCenterColumn].value)

        if row[higherLimitColumn].value != None:
            limit = row[higherLimitColumn].value
        elif row[lowerLimitColumn].value != None:
            limit = row[lowerLimitColumn].value

        limit *= 100
        subsequentBA = row[SubsequentBAColumn].value == "Yes"
        
        textGL = row[GLColumn].value
        if textGL is None:
            continue

        if cardNum not in libraryCard:
            card = Card(cardNum)
            libraryCard[cardNum] = card
        else:
            card = libraryCard[cardNum]



        costCenter = CostCenter(costCenterNum,limit)
        card.add_cost_center(costCenter)

        while True:
            pos = textGL.find(ACCOUNT_NUM_IDENTIFIER)

            if pos == -1:
                break

            refGL = textGL[pos + len(ACCOUNT_NUM_IDENTIFIER) : pos + len(ACCOUNT_NUM_IDENTIFIER) + GL_NUMBER_LENGTH] # grab 8 digit GL code
            textGL = textGL[pos + len(ACCOUNT_NUM_IDENTIFIER) + GL_NUMBER_LENGTH :] # delete text up until the next ACCOUNT_NUM_IDENTIFIER

            costCenter.add_approved_gl(refGL)



def readSAPExport(path,libraryCard,libraryDocument,libraryFlag):
    print("Loading workbook \"" + path + "\".", flush=True)
    SAPWorkbook = openpyxl.load_workbook(path)
    print("Workbook loaded.", flush=True)
    SAPWorksheet = SAPWorkbook.active

    cardsRead = 0
    cardsMatched = 0
    barUpdatesFreq = max(42,round(6.7 * math.log(SAPWorksheet.max_row, 1.1) - 130))
    
    documentColumn = None
    cardNumberColumn = None
    costCenterColumn = None
    GLColumn = None
    postedAmoungColumn = None
    matchingAmountColumn = None
    
    for column in SAPWorksheet.iter_cols():
        if column[0].value == SAP_DOCUMENT_NUMBER_COLUMN_IDENTIFIER:
            documentColumn = column[0].column - 1

        elif column[0].value == SAP_CARD_COLUMN_IDENTIFIER:
            cardNumberColumn = column[0].column - 1

        elif column[0].value == SAP_COST_CENTER_COLUMN_IDENTIFIER:
            costCenterColumn = column[0].column - 1
            
        elif column[0].value == SAP_GL_COLUMN_IDENTIFIER:
            GLColumn = column[0].column - 1

        elif column[0].value == SAP_POSTED_AMOUNT_COLUMN_IDENTIFIER:
            postedAmoungColumn = column[0].column - 1

        elif column[0].value == SAP_MATCHING_AMOUNT_COLUMN_IDENTIFIER:
            matchingAmountColumn = column[0].column - 1

    if checkNone([documentColumn, cardNumberColumn, costCenterColumn, GLColumn, postedAmoungColumn, matchingAmountColumn]):
        print("Could not locate all required columns. Check column identifiers in Python File.")
        return
    
    for row in SAPWorksheet.iter_rows(min_row=2):
        
        if row[0].row % barUpdatesFreq == 0:
            remove_progress_bar()
            place_progress_bar("SAP Extract",20,cardsRead,SAPWorksheet.max_row - 1)
        
        cardsRead += 1

        documentNum = row[documentColumn].value
        cardNumber = row[cardNumberColumn].value
        costCenter = row[costCenterColumn].value
        matchingAmount = row[matchingAmountColumn].value
        refGL = row[GLColumn].value
        

        if checkNone([documentNum,cardNumber,costCenter,matchingAmount,refGL]):
            # flag for a missing value
            continue
        
        cardNumber = cardNumber.lstrip("X")
        refGL = refGL[:GL_NUMBER_LENGTH]
        costCenter = costCenter[:CC_NUMBER_LENGTH]

        try:
            matchingAmount = int(round(float(matchingAmount) * 100))
        except (ValueError, TypeError):
            # flag for invalid matching amount
            continue
        
        try:
            documentNum = int(documentNum)
        except (ValueError, TypeError):
            # flag for invalid document number
            continue

        if costCenter == "#N/A":
            # flag for a missing cost center
            continue
        
        if documentNum in libraryDocument:
            purchase = libraryDocument[documentNum]
        else:
            purchase = Purchase(documentNum,row[postedAmoungColumn].value,cardNumber)
            libraryDocument[documentNum] = purchase
        
        strand = PurchaseStrand(documentNum,matchingAmount,costCenter,refGL)
        purchase.add_purchase_strand(strand)
        
        if cardNumber not in libraryCard:
            # flag for form 32 check
            continue
        cardsMatched += 1

        card = libraryCard[cardNumber]
        center = card.get_cost_center(costCenter)

        if center is None:
            # flag for invalid cost center attributed with card
            continue

        center.add_purchase_strand(strand)
        #print(str(center.get_all_purchase_strands()[int(len(center.get_all_purchase_strands()) - 1)].get_amount()))

    
    remove_progress_bar()
    place_progress_bar("SAP Extract",20,cardsRead,SAPWorksheet.max_row - 1)
    print("\n")
    #print("Hit Rate: " + str(round(cardsMatched*100/cardsRead,4)) + "%")
    

        




def main():
    warnings.filterwarnings(
        "ignore",
        category=UserWarning,
        module="openpyxl"
    )

    libraryFlag = []
    libraryCard = {}
    libraryDocument = {}


    print("Enter Blanket Authority Excel Sheet Directory:")
    #path = input()
    readBASheet("batracking.xlsx",libraryCard,libraryFlag)
    


    print("Enter QA Transactions (SAP Export) Excel Sheet Folder Directory:")
    path = input()
    files = [str(f) for f in Path(path).rglob("*.xlsx")]

    print("\n")
    for file in files:
        readSAPExport(file,libraryCard,libraryDocument,libraryFlag)




    while True:
        print("Enter Card Number:")
        num = input()

        if num == "quit":
            break

        if num == "all":
            for cardNum in libraryCard:
                card = libraryCard[cardNum]
                for costCenter in card.get_all_cost_centers():
                    if costCenter.get_overbudget():
                        print(str(cardNum))
                        break
                
                for costCenter in card.get_all_cost_centers():
                    if not costCenter.get_overbudget():
                        continue

                    # flag for overbudget cost center
                    print("   " + str(costCenter.get_number()) + " ($" + str(convert(costCenter.get_total_spent())) + "/$" + str(convert(costCenter.get_budget())) + ")")

        if not num.isdigit() :#or num not in libraryDocument:
            print("Invalid Document Number.")
            continue

        card = libraryCard[num]

        for costCenter in card.get_all_cost_centers():
            print("   " + str(costCenter.get_number()) + " ($" + str(convert(costCenter.get_total_spent())) + "/$" + str(convert(costCenter.get_budget())) + ")")

main()