from .PDF import PDF

class Company:

    def __init__(self,name, planParticipants = 0, assets = 0, active = True):
        self.name = name
        self.pdfs = []
        self.planParticipants = planParticipants
        self.assets = assets
        self.active = active

    def add_pdf(self, pdf):
        self.pdfs.append(pdf)
    
    def __str__(self):
        output = f"{self.name}: " 
        for pdf in self.pdfs:
            output= output + pdf.__str__() 
        return output