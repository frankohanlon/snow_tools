#!/usr/bin/python3
"""
Simple tool to extract magnaprobe data from Campbell Scientific CR800 data file
and use it to create a kml type file for google earth.

Currently the output kml uses icons created using imagemagick

"""
import simplekml
import argparse
import sys



parser = argparse.ArgumentParser(
  formatter_class=argparse.ArgumentDefaultsHelpFormatter)
parser.add_argument('-s', '--source', type=str, required=True,
                    help="Input source .dat file from MagnaProbe e.g. /path/to/data/SN8898_10.dat")
parser.add_argument('-o', '--output_file', type=str, required=True,
                    help='Output full file name and path for kml.  e.g. /path/magna_output.kml')
parser.add_argument('-u', '--alternate_url', type=str, required=False,
                    help='Alternate url for the icons, default is "http://ngeedata.iarc.uaf.edu/data/icons/"')
parser.add_argument('-i', '--icon_group', type=str, required=False,
                    help='icon group e.g. "number", "blank", or "swe".  default is "number"')

##########################################################################################################################
##########################################################################################################################
##########################################################################################################################
##########################################################################################################################

# read in command line arguments
args = parser.parse_args()
input_file = args.source
output_file = args.output_file

if args.alternate_url != None :
    url = args.alternate_url
else:
    url = 'http://ngeedata.iarc.uaf.edu/data/icons/'

if args.icon_group != None :
    icons = args.icon_group
else:
    icons = 'number'
print('\n\n##########################################################################')
print('##########################################################################')
print ('Input File (.dat) = ', input_file )
print ('Output File (.kml) = ', output_file  )
print ('Alternate URL = ', url  )
print ('Icon Type = ', icons )
print ('\n\n')


# Setup the KML
kml = simplekml.Kml()

# set up some colors
sharedstyle_0 = simplekml.Style()
sharedstyle_0.labelstyle.scale = 0
sharedstyle_0.iconstyle.icon.href = url + icons + '_0.png'

sharedstyle_10 = simplekml.Style()
sharedstyle_10.labelstyle.scale = 0
sharedstyle_10.iconstyle.icon.href = url + icons + '_10.png'

sharedstyle_20 = simplekml.Style()
sharedstyle_20.labelstyle.scale = 0
sharedstyle_20.iconstyle.icon.href = url + icons + '_20.png'

sharedstyle_30 = simplekml.Style()
sharedstyle_30.labelstyle.scale = 0
sharedstyle_30.iconstyle.icon.href = url + icons + '_30.png'

sharedstyle_40 = simplekml.Style()
sharedstyle_40.labelstyle.scale = 0
sharedstyle_40.iconstyle.icon.href = url + icons + '_40.png'

sharedstyle_50 = simplekml.Style()
sharedstyle_50.labelstyle.scale = 0
sharedstyle_50.iconstyle.icon.href = url + icons + '_50.png'

sharedstyle_60 = simplekml.Style()
sharedstyle_60.labelstyle.scale = 0
sharedstyle_60.iconstyle.icon.href = url + icons + '_60.png'

sharedstyle_70 = simplekml.Style()
sharedstyle_70.labelstyle.scale = 0
sharedstyle_70.iconstyle.icon.href = url + icons + '_70.png'

sharedstyle_80 = simplekml.Style()
sharedstyle_80.labelstyle.scale = 0
sharedstyle_80.iconstyle.icon.href = url + icons + '_80.png'

sharedstyle_90 = simplekml.Style()
sharedstyle_90.labelstyle.scale = 0
sharedstyle_90.iconstyle.icon.href = url + icons + '_90.png'

sharedstyle_100 = simplekml.Style()
sharedstyle_100.labelstyle.scale = 0
sharedstyle_100.iconstyle.icon.href = url + icons + '_100.png'

sharedstyle_110 = simplekml.Style()
sharedstyle_110.labelstyle.scale = 0
sharedstyle_110.iconstyle.icon.href = url + icons + '_110.png'

sharedstyle_120 = simplekml.Style()
sharedstyle_120.labelstyle.scale = 0
sharedstyle_120.iconstyle.icon.href = url + icons + '_120.png'

sharedstyle_130 = simplekml.Style()
sharedstyle_130.labelstyle.scale = 0
sharedstyle_130.iconstyle.icon.href = url + icons + '_130.png'

# Done with KML file set up, now read in the magnaprobe data and add to the KML.

try:
    input_data_file = input_file
    dfile = open( input_data_file , 'r')
    all_input_data = dfile.readlines()
    dfile.close()
except:
    print('Input file path error.  File not found.\n\n')
    sys.exit()

fol = kml.newfolder(name='Snow Depth Data')

for line in all_input_data[5:] :
    samples = line.split(',')
    depth = samples[3]
    fl_depth = float(depth)

    if (fl_depth < 6999.) :
        date = samples[0]
        recnum = samples[1]
        counter = samples [2]
        text = "Counter: " + counter + "\n" + date
        latitude = float(samples[5]) + float(samples[14])
        longitude = float(samples[7]) + float(samples[15])
        pnt = fol.newpoint(name=text, coords=[(longitude,latitude)], description=depth)
        if (fl_depth <5.) :
            pnt.style = sharedstyle_0
        elif (fl_depth >=5. and fl_depth < 15.) :
            pnt.style = sharedstyle_10
        elif (fl_depth >=15. and fl_depth < 25.) :
            pnt.style = sharedstyle_20
        elif (fl_depth >=25. and fl_depth < 35.) :
            pnt.style = sharedstyle_30
        elif (fl_depth >=25. and fl_depth < 45.) :
            pnt.style = sharedstyle_40
        elif (fl_depth >=45. and fl_depth < 55.) :
            pnt.style = sharedstyle_50
        elif (fl_depth >=55. and fl_depth < 65.) :
            pnt.style = sharedstyle_60
        elif (fl_depth >=65. and fl_depth < 75.) :
            pnt.style = sharedstyle_70
        elif (fl_depth >=75. and fl_depth < 85.) :
            pnt.style = sharedstyle_80
        elif (fl_depth >=85. and fl_depth < 95.) :
            pnt.style = sharedstyle_90
        elif (fl_depth >=95. and fl_depth < 105.) :
            pnt.style = sharedstyle_100
        elif (fl_depth >=105. and fl_depth < 115.) :
            pnt.style = sharedstyle_110
        elif (fl_depth >=115. and fl_depth < 125.) :
            pnt.style = sharedstyle_120
        elif (fl_depth >=125. and fl_depth < 6999.) :
            pnt.style = sharedstyle_130

try:
    kml.save(output_file)
    print('Program successful.  Now view the kml with the program of your choice.')
except:
    print('Perhaps incorrect path specified for location of output kml?\n\n')
    sys.exit()

