from googleapiclient import discovery
import re
from datetime import datetime, date


def stringify(value):
    """ Convert values to strings and remove leading/trailing blanks """
    # Let's normalize everything to strings/unicode strings
    if isinstance(value, (int, long, float, bool)):
        value = '%s' % value
    if isinstance(value, bool):
        value = '1' if value else '0'
    elif isinstance(value, (datetime, date)):
        value = ('%s' % value)[0:10]
    return value.strip()

class GSheet:
    """Makes a Google spreadsheet easier to deal with:"""
    
    # We return a class with items that match the column labels as normalized
    class GSheetRow:
        def __init__(self, names, values):
            for (name, value) in zip(names, values):
                self.__dict__[name] = value
    

    
    
    def __init__(self, sheetid, apikey):
        if '/' in sheetid:
            # Have a whole URL; get the key
            sheetid = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', sheetid).groups()[0]
            
        service = discovery.build('sheets', 'v4', developerKey=apikey)
        request = service.spreadsheets().values().get(spreadsheetId=sheetid, range='a1:zz999')
        self.values = request.execute()['values']
        self.rownum = 0
        
    
        # Save the unmodified labels
        self.origlabels = self.values[0]
    
        # Now, build the label lookup dictionary
        self.lookup = {}
        self.labels = []
        colnum = 0
        for item in self.origlabels:
            # Normalize the label name: lowercase, and remove non-alphamerics
            item = re.sub('[\W_]+', '', item.lower())
            self.lookup[item] = colnum
            self.lookup[colnum] = item
            self.labels.append(item)
            colnum += 1
    
            
            
    def __iter__(self):
        return self
    
            
    def next(self):
        self.rownum += 1
        if self.rownum >= len(self.values):
            raise StopIteration
        else:
            self.row = self.values[self.rownum]
            self.strings = [stringify(item) for item in self.row]
            return self.GSheetRow(self.labels, self.strings)
            