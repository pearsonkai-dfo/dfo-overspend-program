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

    def contains_gl(self, gl_number: str):
        return gl_number in self.approved_gl_accounts

    def get_number(self):
        return self.number

    def get_budget(self):
        return budget

    def get_all_gl(self):
        return self.gl_accounts
        
    def get_gl(self, gl_number: str):
        return self.gl_accounts.get(gl_number)

    def get_total_spent(self):
        total = 0

        for strand in self.invoices:
            total += strand.get_amount()

        return total

    def get_remaining_budget(self):
        return self.budget - self.get_total_spent()

    def get_overbudget(self):
        return self.get_remaining_budget() < 0