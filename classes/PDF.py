class PDF:
    
    def __init__ (self, url, title="", fileLocation="", source ="TransAmerica"):
        self.url = url
        self.title = title
        self.fileLocation= fileLocation
        self.source = source

    def __str__(self):
        return f"{self.title}: {self.url} \n" 
