class Company:
    def __init__(self,name, planParticipants = 0, assets = 0):
        self.name = name
        self.pdfs = []
        self.planParticipants = planParticipants
        self.assets = assets


    def add_pdf_link(self, pdf_link):
        self.pdfs.append(pdf_link)
    
    def __str__(self):
        return f"{self.name}: {', '.join(self.pdfs)}"