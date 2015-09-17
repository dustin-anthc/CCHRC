import xlrd
from lxml import etree

# Settings
worksheet_name = 'NOAA_1981-2010_Annual_temps.xlsx'
latitude_column = 'LATITUDE'
longitude_column = 'LONGITUDE'
#label_column = 'Annual Tavg'
label_column = 'DJF Tavg'
description_column = 'STATION_NAME'

# Create the root kml element and document
kml_element = etree.Element('kml', xmlns="http://www.opengis.net/kml/2.2")
kml_doc = etree.ElementTree(kml_element)
document_element = etree.SubElement(kml_element, 'Document')
document_name_element = etree.SubElement(document_element, 'name')
document_name_element.text = worksheet_name
style_element = etree.SubElement(document_element,'Style', id="style1")

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
        styleurl_element.text = '#style1'

        point_element = etree.SubElement(placemark_element, 'Point')

        coordinates_element = etree.SubElement(point_element, 'coordinates')
        coordinates_element.text = str(longitude) + ',' + str(latitude)
    except ValueError:
        print "Invalid Coords: %s" % row_data

# Write the KML file
with open(worksheet_name + '.kml', 'w') as f:
    kml_doc.write(f, xml_declaration=True, encoding='utf-8')
