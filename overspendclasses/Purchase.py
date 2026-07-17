from .PurchaseStrand import PurchaseStrand

class Purchase:
    """
    Container for storing document number and GL accounts the matching amounts were allocated to.

    The posted amount is the total price of the transaction on the Excel sheet.

    The sum of matching amounts in all purchase strands should be equal to the posted amount. If not, the transaction should be flagged for a manual review.
    """
    def __init__(self, document_number: str, posted_amount: int, cardholder: str):
        self.document_number: int = document_number
        self.posted_amount: int = posted_amount
        self.cardholder: str = cardholder
        self.purchase_strands: list[PurchaseStrand] = []

    def add_purchase_strand(self, purchase_strand):
        self.purchase_strands.append(purchase_strand)

    def get_document(self):
        return self.document_number

    def get_posted_amount(self):
        return self.posted_amount

    def get_cardholder(self):
        return self.cardholder

    def get_all_purchase_strand(self):
        return self.purchase_strands

    def get_calculated_total(self):
        total = 0

        for strand in self.purchase_strands:
            total += strand.get_amount()

        return total

    def get_difference(self):
        return self.posted_amount - self.get_calculated_total()

    def is_balanced(self):
        return abs(self.get_difference()) < 0.01

    def get_all_cost_center(self):
        cost_centers = []

        for strand in self.purchase_strands:
            cost_centers.append(strand.get_cost_center())

        return cost_centers

    def get_all_gl(self):
        gl_numbers = []

        for strand in self.purchase_strands:
            gl_numbers.append(strand.get_gl())

        return gl_numbers