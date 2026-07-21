class PurchaseStrand:
    """
    Purchase strands are allocation of cost to different GLs from a purchase.

    A purchase requires at least one purchase strand. 
    """
    def __init__(self,document: int, amount: int, cost_center: int, gl: int):
        self.document = document
        self.amount : int = amount
        self.cost_center = cost_center
        self.gl = gl

    def get_document(self):
        return self.document

    def get_amount(self):
        return self.amount

    def get_cost_center(self):
        return self.cost_center

    def get_gl(self):
        return self.gl