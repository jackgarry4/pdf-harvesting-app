from classes.PDF import PDF

class Company:

    def __init__(self, name, planParticipants = 0, assets = 0):
        self.name = name
        self.pdfs = set()
        self.planParticipants = planParticipants
        self.assets = assets

    def add_pdf(self, pdfURL, pdfTitle):
        if pdfURL not in (pdf.url for pdf in self.pdfs):
            newPDF = PDF(url = pdfURL, title = pdfTitle)
            self.pdfs.add(newPDF)
            
    
    def __str__(self):
        output = f"{self.name}: " 
        for pdf in self.pdfs:
            output= output + pdf.__str__() 
        return output