import xlrd
from lxml import etree
import sys

# NOTE: Need to pass the Spreadsheet file name as a command line
# argument to this script.

# Settings
worksheet_name = sys.argv[1]  # Spreadsheet to read
latitude_column = 'latitude'
longitude_column = 'longitude'
label_column = 'label'
description_column = 'stn_name'

# Create the root kml element and document
kml_element = etree.Element('kml', xmlns="http://www.opengis.net/kml/2.2")
kml_doc = etree.ElementTree(kml_element)
document_element = etree.SubElement(kml_element, 'Document')
document_name_element = etree.SubElement(document_element, 'name')
document_name_element.text = worksheet_name

# different pushpin styles for each type of climate file

# default yellow pushpin for TMY2
style_element = etree.SubElement(document_element, 'Style', id="tmy2")

# Green Pushpin for TMY3 Class II
style_element = etree.SubElement(document_element, 'Style', id="tmy3-II")
icon_style = etree.SubElement(style_element, 'IconStyle')
icon = etree.SubElement(icon_style, 'Icon')
icon.text = 'http://maps.google.com/mapfiles/kml/pushpin/grn-pushpin.png'

# Red Pushpin for TMY3 Class III
style_element = etree.SubElement(document_element, 'Style', id="tmy3-III")
icon_style = etree.SubElement(style_element, 'IconStyle')
icon = etree.SubElement(icon_style, 'Icon')
icon.text = 'http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png'

# Open the Excel worksheet and extract the header row
worksheet = xlrd.open_workbook(worksheet_name).sheet_by_index(0)
sheet_keys = worksheet.row_values(0)

# Loop through the content rows in the worksheet and create KML placemarks
for row_number in range(1, worksheet.nrows):
    values = worksheet.row_values(row_number)
    row_data = dict(zip(sheet_keys, values))
    try:
        latitude = float(row_data[latitude_column])
        longitude = float(row_data[longitude_column])

        placemark_element = etree.SubElement(document_element, 'Placemark')

        placemark_name_element = etree.SubElement(placemark_element, 'name')
        placemark_name_element.text = str(row_data[label_column])

        snippet_element = etree.SubElement(placemark_element, 'snippet')
        snippet_element.text = str(row_data[description_column])

        description_element = etree.SubElement(placemark_element, 'description')
        description_element.text = str(row_data[description_column])

        styleurl_element = etree.SubElement(placemark_element, 'styleUrl')
        tmy_type = row_data['tmy_type']
        tmy_class = row_data['tmy_class']
        styleurl_element.text = '#tmy2' if tmy_type==2 else '#tmy3-II' if tmy_type==3 and tmy_class=='II' else '#tmy3-III'

        point_element = etree.SubElement(placemark_element, 'Point')

        coordinates_element = etree.SubElement(point_element, 'coordinates')
        coordinates_element.text = str(longitude) + ',' + str(latitude)
    except ValueError:
        print "Invalid Coords: %s" % row_data

# Write the KML file
with open(worksheet_name + '.kml', 'w') as f:
    kml_doc.write(f, xml_declaration=True, encoding='utf-8')
