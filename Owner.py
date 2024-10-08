import Pet
from openpyxl import Workbook, load_workbook
class Owner:
    def __init__(self):
        self.attributes = []
        self.link = None

    def add_attribute(self, datum):

        self.attributes.append(datum)

    def link_to_pet(self, other):
        self.link = other
        
        
        
    