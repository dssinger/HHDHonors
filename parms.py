#!/usr/bin/env python3
""" Read the honors.yaml file and set up appropriate parameters """
import yaml
class Parms:
    def __init__(self):
        self.contents = []
        with open('honors.yaml', 'r') as f:
            parms = yaml.load(f, Loader=yaml.FullLoader)
        # Promote values to attributes
        for item in parms:
            self.contents.append(item)
            setattr(self, item, parms[item])
    
if __name__ == '__main__':
    parms = Parms()
    for item in parms.contents:
        print('%s = %s' % (item, getattr(parms,item)))
    
