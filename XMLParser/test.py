from xml.dom import minidom

xmldoc = minidom.parse('/Users/enricoschmidt/Downloads/IS185989_0.xml')
itemlist = xmldoc.getElementsByTagName('item') 
print(len(itemlist))
print(itemlist[0].attributes['name'].value)
for s in itemlist :
    print(s.attributes['name'].value+":",
            s.childNodes[0].nodeValue)