class Attribute:
    def __init__(self):
        self.name = "new attribute"
        self.value = None

    def set_name(self, neo):
        self.name = neo

    def get_name(self):
        if (self.name != "new attribute"):
            return self.name
        else:
            print("ERROR: tried to get the name of an attribute that wasn't "
                  "yet given a name")
            return "!!!!!!!"

    def set_value(self, new_thing):
        self.value = new_thing

    def get_value(self):
        if (self.value is not None):
            return self.value
        else:
            print("Just an alert: " + self.name + " has no value.")
            return None

# Frazee's