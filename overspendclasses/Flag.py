import openpyxl

flag_codes = {
    100: "",
    101: "Missing Document Number",
    102: "Missing Card Number",
    103: "Missing Cardholder Name",
    104: "Missing Cost Center Number",
    105: "Missing Matching Amount",
    106: "Missing GL Account",
    107: "Authorized Lower and Higher Budget on BA Tracking",

    201: "Active BA, No Matching GL Account",
    202: "Active BA, No Matching Cost Center",
    203: "No BA",

    301: "Overbudget",
    302: "Mismatched Posted And Summed Matching Price",
    303: "Purchase Exceeds $10 000",

    401: "Invalid Matching Amount",
    402: "Invalid Document Number"
}
"""
Flag code first letter id:
1 - Missing information
2 - No Authoritization
3 - General purpose flags
4 - Major issue

"""

class Flag:
    """
    A flag is an object used to identify an issue or important piece of information. 
    
    Flags in this program are used to mark overspending cardholders, issues in a file, purchases over $10000, and other miscellaneous things.
    """
    def __init__(self,flag_num,flag_description,flag_path,flag_row,flag_column,flag_document):
        self.number = flag_num
        self.description = flag_description
        self.path = flag_path
        self.row = flag_row
        self.column = flag_column
        self.document = flag_document # if applicable

    def workbook_print(self, OUTWorksheet):
        OUTWorksheet.append([
            str(self.number),
            str(self.description),
            str(self.path),
            str(self.row),
            str(self.column),
            str(self.document)
        ])