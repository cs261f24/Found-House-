class Pet:
    def __init__(self):
        self.attributes = []
        self.link = None

    def add_attribute(self, datum):
        self.attributes.append(datum)

    def link_to_owner(self, other):
        self.link = other
