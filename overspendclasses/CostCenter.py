from .PurchaseStrand import PurchaseStrand

class CostCenter:
    """
    Represents a single budget allocation from a blanket authority.
    
    Cost center numbers are unique.
    """

    def __init__(self, number: str, budget: float):
        self
        self.number: int = number
        self.budget: int = budget
        self.approved_gl_accounts: list[int] = []
        self.invoices: list[PurchaseStrand] = []

    def add_approved_gl(self, gl_account: int):
        self.approved_gl_accounts.append(gl_account)
        
    def add_purchase_strand(self, strand):
        self.invoices.append(strand)

    def contains_gl(self, gl_number: str):
        return gl_number in self.approved_gl_accounts

    def get_number(self):
        return self.number

    def get_budget(self):
        return self.budget

    def get_all_gl(self):
        return self.gl_accounts
    
    def get_all_purchase_strands(self):
        return self.invoices
        
    def get_gl(self, gl_number: str):
        return self.gl_accounts.get(gl_number)

    def get_approved_gls(self):
        return self.approved_gl_accounts

    def get_total_spent(self):
        total = 0

        for strand in self.invoices:

            if int(strand.get_gl()) in self.approved_gl_accounts:
                total += int(strand.get_amount())
            elif str(strand.get_gl()) in self.approved_gl_accounts:
                total += int(strand.get_amount())

        return total

    def get_remaining_budget(self, percent_over: float = 1.00):
        return self.budget * percent_over - self.get_total_spent()

    def get_overbudget(self, percent_over: float = 1.00):
        return self.get_remaining_budget(percent_over) <= 0