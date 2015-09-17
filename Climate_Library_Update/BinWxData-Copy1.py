#!python3
'''
Class to create bin weather data for use by the cooling module in AkWarm
using TMY3 data, avaialable from 
http://rredc.nrel.gov/solar/old_data/nsrdb/1991-2005/tmy3/by_state_and_city.html
For information about the results produced, see documentation of the 'results' 
method of the BinWxData class below.
Main script processes all TMY3 files in a directory and produces the bin 
weather data from them.  
'''

import os, sys

# Map tmy file names to main weather city IDs
tmyToCity = {
'700260TY.csv': 4,   # BARROW W POST-W ROGERS ARPT [NSA - ARM]
'701330TY.csv': 14,  # KOTZEBUE RALPH WEIN MEMORIAL
'701740TY.csv': 20,  # BETTLES FIELD
'702000TY.csv': 6,   # NOME MUNICIPAL ARPT
'702190TY.csv': 5,   # BETHEL AIRPORT
'702310TY.csv': 15,  # MCGRATH ARPT
'702510TY.csv': 17,  # TALKEETNA STATE ARPT
'702610TY.csv': 2,   # FAIRBANKS INTL ARPT
'702710TY.csv': 9,   # GULKANA INTERMEDIATE FIELD
'702730TY.csv': 1,   # ANCHORAGE INTL AP
'702740TY.csv': 22,  # PALMER MUNICIPAL, AkWarm city is Wasilla
'703080TY.csv': 16,  # ST PAUL ISLAND ARPT
'703160TY.csv': 11,  # COLD BAY ARPT
'703260TY.csv': 13,  # KING SALMON ARPT
'703410TY.csv': 12,  # HOMER ARPT
'703500TY.csv': 8,   # KODIAK AIRPORT
'703610TY.csv': 19,  # YAKUTAT STATE ARPT
'703810TY.csv': 3,   # JUNEAU INT'L ARPT
'703950TY.csv': 23,  # KETCHIKAN INTL AP
'704540TY.csv': 21,  # ADAK NAS
}

from collections import Counter

class BinWxData:

   # Width of the temperature bins in degrees F.  Should be an even 
   # number so midpoint is an integer
   BIN_WIDTH = 2

   # Minimum temperature below which no bins are created.  Should be an
   # integer.
   MIN_TEMP = 50

   def __init__(self):
      self.reset()

   def reset(self):
      '''Resets the counters and accumulators to start a new binning
      process.
      '''
      self.hours = 0         # tracks total hours of data analyzed
      self.totSolar = 0.0    # tracks total solar for all hours
      self.temp = Counter()  # number of hours in each temperature bin
      self.solar = Counter() # solar occuring in each temperature bin
   
   def add(self, tempVal, solarVal):
      '''Adds an hour of temp/solar data to the analysis.
      'tempVal' is the temperature during the hour, deg F
      'solarVal' is the solar during the hours, any units
      '''
      self.hours += 1
      self.totSolar += solarVal

      # do not bin up if below the minimum temperature
      minT = BinWxData.MIN_TEMP
      if tempVal < minT:
         return

      # determine the temperature bin, and add a count to that
      # bin; add the solar to the solar Counter for that bin.
      bw = BinWxData.BIN_WIDTH
      bin = int((tempVal - minT) // bw) * bw + minT + bw // 2
      self.temp[bin] += 1
      self.solar[bin] += solarVal

   def results(self):
      '''Returns the results of the binning process in the form
      of a string.  In the string, each bin is separated by a space
      and information within the one bin is separated by commas. An
      example is:  '71,23,2.560 73,14,2.923'
      which gives information for two bins.  For each bin, the first
      value is the midpoint temperature of the bin, e.g. for a bin
      width of 2 deg F, a value of 71 means the bin is  
            70 F <= temperature < 72 F.
      The second number describing the bin is the number of hours that
      fall into the indicated bin.  The third number describing the bin
      is the ratio of solar power (W or Btu/hour) occuring for the bin
      hours relative to the average solar power for the time period 
      analyzed, averaging all of the hours in the time period, including 
      those that fall outside of the bins.
      '''
      avgSolar = self.totSolar / self.hours
      tempBins = sorted(self.temp)
      result = ''
      for b in tempBins:
         if len(result):
            result += " "
         binCount = self.temp[b]
         result += '%s,%s,%.3f' % (b, binCount, self.solar[b] / binCount / avgSolar)
      return result

def test():
   '''
   Test Routine to exercise BinWxData class.
   '''
   bw = BinWxData()
   bw.add(70, 120)
   bw.add(23.0, 10)
   bw.add(71.5, 140)
   bw.add(80, 200)
   print(bw.results())
   bw.reset()
   bw.add(73, 400)
   bw.add(45, 0)
   print(bw.results())

def tmyList():
   fout = open('list.txt', 'w')
   for f in glob('tmy/*.csv'):
      # open file and read in the two header rows
      fin = open(f)
      hdr1 = fin.readline()

      # extract the station name out of the first header row and 
      # print it to the screen and output file.
      flds = hdr1.split(',')
      stn = flds[1][1:-1]
      baseFname = os.path.basename(f)

      print("'%s': ,  # %s" % (baseFname, stn), file=fout)
   

if __name__=="__main__":
   
   from glob import glob

   fout = open('results.txt', 'w')
   flibOut = open('lib_xfer.txt', 'w')

   # Read all the TMY3 files in the 'tmy' directory and process
   for f in glob('tmy/*.csv'):

      # extract out base file name and then determine AkWarm weather
      # city ID
      baseFname = os.path.basename(f)
      akWxID = tmyToCity[baseFname]

      # open file and read in the two header rows
      fin = open(f)
      hdr1 = fin.readline()
      hdr2 = fin.readline()

      # extract the station name out of the first header row and 
      # print it to the screen and output file.
      flds = hdr1.split(',')
      stn = flds[1][1:-1]
      print(stn)
      print('\n%s, %s' % (stn, f), file=fout)

      # create a bin weather analyzer
      bw = BinWxData()

      line = fin.readline() # start the process by reading the first data line
      lastMo = 1            # tracks month of last data line
      while line:
         flds = line.split(',')
         mo = int(flds[0][:2])    # month of the this hour, 1 - 12
         if mo != lastMo:

            # have moved to a new month, output results from completed month
            print('Month %d\t%s' % (lastMo, bw.results()), file=fout)

            # add to the AkWarm library transfer file
            print('%s\t%s\t%s' % (akWxID, lastMo, bw.results()), file=flibOut)

            bw.reset()
            lastMo = mo
         
         dbTemp = float(flds[31]) * 1.8 + 32.0

         # For solar, we weight direct normal solar by 40% and diffuse
         # horizontal by 50% (vertical surfaces have 50% view of sky). 
         # The direct solar weighting ought to reflect the % of direct 
         # solar received by a typical fenestration in the building, less
         # than 100% because windows are not directly pointed at the sun
         # all of the time.  These two percentages do NOT need to add to 100%,
         # since the ultimate solar value resulting from this process is a ratio
         # of solar power in a particular bin to the average solar power.
         # These factors are a judgement call, but the results do not change
         # tremendously varying the direct normal weighting from 0.2 to 0.8.
         solar = 0.4 * float(flds[7]) + 0.5 * float(flds[10])
   
         # add the data to the analyzer
         bw.add(dbTemp, solar)

         # read the next line
         line = fin.readline()

      # print out the month just completed
      print('Month %d\t%s' % (lastMo, bw.results()), file=fout)

      # add to the AkWarm library transfer file as well
      print('%s\t%s\t%s' % (akWxID, lastMo, bw.results()), file=flibOut)
