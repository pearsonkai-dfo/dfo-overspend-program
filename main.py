from time import sleep
import numbers
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import warnings
from pathlib import Path 
import math

#from decimal import Decimal

from overspendclasses import Card
from overspendclasses import CostCenter
from overspendclasses import Purchase
from overspendclasses import PurchaseStrand
from overspendclasses import Flag
from overspendclasses import flag_codes as fc


##### GENERAL CONSTANTS #####

GL_NUMBER_LENGTH = 8
CC_NUMBER_LENGTH = 6
PERCENT_OVER = 0.75


##### SAP EXPORT SHEET CONSTANTS #####

SAP_DOCUMENT_NUMBER_COLUMN_IDENTIFIER = "Document Number"
SAP_CARD_COLUMN_IDENTIFIER = "Card Number"
SAP_COST_CENTER_COLUMN_IDENTIFIER = "Cost Center"
SAP_GL_COLUMN_IDENTIFIER = "G/L Account"
SAP_POSTED_AMOUNT_COLUMN_IDENTIFIER = "Transaction Amount"
SAP_MATCHING_AMOUNT_COLUMN_IDENTIFIER = "Matching Amount"
SAP_DEBIT_CREDIT_COLUMN_IDENTIFIER = "Transaction Debit/Credit Code"

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


def resolve_matching_amount_for_transaction_code(transaction_code, matching_amount):
    normalized_code = str(transaction_code).strip().upper() if transaction_code is not None else ""

    if normalized_code == "D":
        return matching_amount
    if normalized_code == "C":
        return -matching_amount
    return None


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

            refGL = str(textGL[pos + len(ACCOUNT_NUM_IDENTIFIER) : pos + len(ACCOUNT_NUM_IDENTIFIER) + GL_NUMBER_LENGTH]) # grab 8 digit GL code
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
    debitCreditColumn = None
    
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

        elif column[0].value == SAP_DEBIT_CREDIT_COLUMN_IDENTIFIER:
            debitCreditColumn = column[0].column - 1

    if checkNone([documentColumn, cardNumberColumn, costCenterColumn, GLColumn, postedAmoungColumn, matchingAmountColumn, debitCreditColumn]):
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
        transactionCode = row[debitCreditColumn].value
        
        if documentNum == None:
            # missing document number
            libraryFlag.append(Flag(101,fc[101],path,row[0].row,documentColumn,""))
            continue

        if cardNumber == None:
            # missing card number
            libraryFlag.append(Flag(102,fc[102],path,row[0].row,cardNumberColumn,documentNum))
            continue

        if costCenter == None or costCenter == "#N/A" or costCenter == "":
            # missing cost center number
            libraryFlag.append(Flag(104,fc[104],path,row[0].row,costCenterColumn,documentNum))
            continue

        if matchingAmount == None:
            # missing matching amount
            libraryFlag.append(Flag(105,fc[105],path,row[0].row,matchingAmountColumn,documentNum))
            continue

        if refGL == None:
            # missing GL number
            libraryFlag.append(Flag(106,fc[106],path,row[0].row,GLColumn,documentNum))
            continue
        
        cardNumber = cardNumber.lstrip("X")

        try:
            refGL = int(str(refGL)[:GL_NUMBER_LENGTH])
        except (ValueError, TypeError):
            # flag for invalid GL number
            libraryFlag.append(Flag(106,fc[106],path,row[0].row,GLColumn,documentNum))
            continue

        try:
            costCenter = int(str(costCenter)[:CC_NUMBER_LENGTH])
        except (ValueError, TypeError):
            # flag for invalid cost center
            libraryFlag.append(Flag(104,fc[104],path,row[0].row,costCenterColumn,documentNum))
            continue

        if str(costCenter)[0] == "6" or str(costCenter)[0] == "7":
            # ignore cost centers starting with 6 or 7; they are CCG
            continue
        
        
        adjustedMatchingAmount = resolve_matching_amount_for_transaction_code(transactionCode, matchingAmount)

        if adjustedMatchingAmount is None:
            libraryFlag.append(Flag(407,fc[407],path,row[0].row,debitCreditColumn,documentNum))
            continue

        try:
            matchingAmount = int(round(float(adjustedMatchingAmount) * 100))
        except (ValueError, TypeError):
            # flag for invalid matching amount
            libraryFlag.append(Flag(401,fc[401],path,row[0].row,matchingAmountColumn,documentNum))
            continue
        
        try:
            documentNum = int(documentNum)
        except (ValueError, TypeError):
            # flag for invalid document number
            libraryFlag.append(Flag(402,fc[402],path,row[0].row,documentColumn,documentNum))
            continue
        
        if documentNum in libraryDocument:
            purchase = libraryDocument[documentNum]
        else:
            purchase = Purchase(documentNum,row[postedAmoungColumn].value,cardNumber)
            libraryDocument[documentNum] = purchase

        

        #if matchingAmount > 10000 * 100:
            # flag for over $10000
            #libraryFlag.append(Flag(303,fc[303],path,row[0].row,matchingAmountColumn,documentNum))
        
        strand = PurchaseStrand(documentNum,matchingAmount,costCenter,refGL)
        purchase.add_purchase_strand(strand)
        
        if cardNumber not in libraryCard:
            # flag for form 32 check
            libraryFlag.append(Flag(203,str(fc[203]) + " (Card Num: " + str(cardNumber) + ")",path,row[0].row,str(cardNumberColumn),str(documentNum)))
            continue
        else:
            cardsMatched += 1

        card = libraryCard[cardNumber]
        center = card.get_cost_center(costCenter)


        if center is None:
            # flag for invalid cost center attributed with card
            libraryFlag.append(Flag(202,str(fc[202]) + " (Card Num: " + str(cardNumber) + ")" + " (CC: " + str(costCenter) + ")",path,row[0].row,costCenterColumn,documentNum))
            continue

        if str(refGL) not in center.get_approved_gls():
            # flag for GL not approved for cost center
            libraryFlag.append(Flag(201,str(fc[201]) + " (Card Num: " + str(cardNumber) + ")" + " (CC: " + str(costCenter) + ")" + " (GL: " + str(refGL) + ")",path,row[0].row,GLColumn,documentNum))
            continue

        center.add_purchase_strand(strand)

    
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

    print("!WARNING! Ensure all relevant files are closed before continuing.\n")
    sleep(2)
    print("Enter Blanket Authority Excel Sheet Directory: ", end="")
    path = input()
    print()
    readBASheet(path,libraryCard,libraryFlag)
    


    print("Enter QA Transactions (SAP Export) Excel Sheet Folder Directory: ", end="")
    path = input()
    print()
    files = [str(f) for f in Path(path).rglob("*.xlsx")]

    for file in files:
        readSAPExport(file,libraryCard,libraryDocument,libraryFlag)

    for purchaseKey in libraryDocument: 
        if libraryDocument[purchaseKey].get_calculated_total() > 10000 * 100:
            # flag for over $10000
            libraryFlag.append(Flag(303,fc[303] + " ($" + str(convert(libraryDocument[purchaseKey].get_calculated_total())) + ")",path,"","",libraryDocument[purchaseKey].get_document()))


    for cardNum in libraryCard:
        card = libraryCard[cardNum]
        
        for costCenter in card.get_all_cost_centers():
            if not costCenter.get_overbudget(PERCENT_OVER):
                continue
            
            allDocuments = ""
            for purchaseStrand in costCenter.get_all_purchase_strands():
                if purchaseStrand.get_gl() in costCenter.get_approved_gls():
                    allDocuments += str(purchaseStrand.get_document()) + " (GL: " + str(purchaseStrand.get_gl()) + ") ($" + str(convert(purchaseStrand.get_amount())) + ")\n"
            libraryFlag.append(Flag(301,fc[301] + " Card: " + str(cardNum) + ", CC: " + str(costCenter.get_number()) + " ($" + str(convert(costCenter.get_total_spent())) + "/$" + str(convert(costCenter.get_budget())) + ")","","","",allDocuments))
            
    
    wb = Workbook()
    ws = wb.active
    ws.title = "output"
    ws.column_dimensions["A"].width = 12  # Flag Code
    ws.column_dimensions["B"].width = 55  # Description
    ws.column_dimensions["C"].width = 65  # File Path
    ws.column_dimensions["D"].width = 10  # Row
    ws.column_dimensions["E"].width = 10  # Column
    ws.column_dimensions["F"].width = 45  # Document

    ws.append(["Flag Code", "Description", "File Path", "Row", "Column", "Document Number"])

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="4472C4"  # Excel blue
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    header_alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment


    for flag in libraryFlag:
        flag.workbook_print(ws)

    for row in ws.iter_rows():
        max_lines = 1

        for cell in row:
            if cell.value:
                # Enable wrapping
                cell.alignment = Alignment(wrap_text=True)

                # Estimate lines based on column width
                col_width = ws.column_dimensions[cell.column_letter].width or 10

                text = str(cell.value)
                estimated_lines = math.ceil(len(text) / col_width)

                max_lines = max(max_lines, estimated_lines)

        # Approx. 16 points per line
        ws.row_dimensions[cell.row].height = max_lines * 16

    ws.auto_filter.ref = ws.dimensions
    wb.save("OverspendOutput.xlsx")






    while True:
        print("Enter Card Number:")
        num = input()

        if num == "quit":
            break

        if num == "many":
            print(str(len(libraryFlag)))
            continue

        if num == "all":
            for cardNum in libraryCard:
                card = libraryCard[cardNum]
                for costCenter in card.get_all_cost_centers():
                    if costCenter.get_overbudget():
                        print(str(cardNum))
                        break
                
                for costCenter in card.get_all_cost_centers():
                    if not costCenter.get_overbudget(PERCENT_OVER):
                        continue

                    # flag for overbudget cost center
                    print("   " + str(costCenter.get_number()) + " ($" + str(convert(costCenter.get_total_spent())) + "/$" + str(convert(costCenter.get_budget())) + ")")

                    for strand in costCenter.get_all_purchase_strands():
                        print("      " + str(strand.get_document()) + " (GL: " + str(strand.get_gl_account()) + ") ($" + str(convert(strand.get_amount())) + ")")
            continue

        if not num.isdigit() :#or num not in libraryDocument:
            print("Invalid Document Number.")
            continue

        card = libraryCard[num]

        for costCenter in card.get_all_cost_centers():
            print("   " + str(costCenter.get_number()) + " ($" + str(convert(costCenter.get_total_spent())) + "/$" + str(convert(costCenter.get_budget())) + ")")

main()