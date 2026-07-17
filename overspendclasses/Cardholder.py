from .CostCenter import CostCenter

class Cardholder:
    """
    Cardholders are employees with a budget allocation with a cost center from the blanket authority sheet.

    A cardholder may have multiple seperate budget allocations from the same cost center.
    A cardholder may have multiple budget allocations from different cost centers.

    A cardholder should be flagged by the program for a manual review if the sum of the matching amounts from GLs in one of their cost center allocations exceeds its budget.
    """
    def __init__(self, name: str):
        self.name: str = name
        self.cost_centers: list[CostCenter] = []

    def add_cost_center(self, cost_center: CostCenter):
        self.cost_centers.append(cost_center)

    def get_name(self):
        return self.name
    
    def get_all_cost_center(self):
        return self.cost_centers

    def get_cost_center(self, cost_center_number: int):
        for center in self.cost_centers:
            if center.get_number() == cost_center_number:
                return center
        
        return None

    def get_overbudget_cost_center(self):
        overbudget_cost_centers = []

        for center in self.cost_centers:
            if center.is_over_budget():
                overbudget_cost_centers.append(center)

        return overbudget_cost_centers

    def get_total_spent(self):
        total = 0

        for cost_center in self.cost_centers:
            total += cost_center.get_total_spent()

        return total

    def has_cost_center(self,cost_center_number: int):
        for center in self.cost_centers:
            if center.get_number() == cost_center_number:
                return True
        
        return False