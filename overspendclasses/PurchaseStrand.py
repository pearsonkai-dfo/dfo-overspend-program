class PurchaseStrand:
    """
    Purchase strands are allocation of cost to different GLs from a purchase.

    A purchase requires at least one purchase strand. 
    """
    def __init__(self, document: int, amount: int, cost_center: int, gl: int, name: str, posted_date=None, transaction_date=None, fund: str = "", supplier: str = ""):
        self.document = document
        self.amount: int = amount
        self.cost_center = cost_center
        self.gl = gl
        self.name = name
        self.posted_date = posted_date
        self.transaction_date = transaction_date
        self.fund = fund
        self.supplier = supplier

    def get_document(self):
        return self.document

    def get_amount(self):
        return self.amount

    def get_cost_center(self):
        return self.cost_center

    def get_gl(self):
        return self.gl

    def get_name(self):
        return self.name

    def get_supplier(self):
        return self.supplier

    def get_transaction_date(self):
        return self.transaction_date

    def get_fund(self):
        return self.fund

    def get_posted_date(self):
        return self.posted_date