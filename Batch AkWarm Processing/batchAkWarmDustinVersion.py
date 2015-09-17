import os
import subprocess
from lxml import etree
import csv

AkWarmExe = r'C:\Program Files (x86)\AHFC\AkWarm\AkWarmCL.exe'
inputDirectory = r'C:\Users\dustin\Dropbox\AKWarm Documentation\02. Work\02. Task X - Build AKWarm Rating Comparison Tool\Randomly Sampled Files\Sample File Set - 4-16-2015'
outputDirectory = r'C:\Users\dustin\Dropbox\AKWarm Documentation\02. Work\02. Task X - Build AKWarm Rating Comparison Tool\Randomly Sampled Files\OutputXML'
libraryToUse = 'BestMatch'  # valid choices are Newest, BestMatch, or a filename


class CalculationResults(object):

    def __init__(self, filename):
        self.filename = os.path.basename(filename)
        input_file = filename
        output_file = os.path.join(outputDirectory, os.path.splitext(self.filename)[1] + '.xml')
        self.exitCode = subprocess.call([AkWarmExe, input_file, output_file, libraryToUse])
        if self.exitCode in [0, 1, 2]:
            try:
                self.tree = etree.parse(output_file)
            except:
                self.tree = None
        else:
            self.tree = None

    @staticmethod
    def as_keys():
        return ['FileName',
                'Result',
                'RatingPoints',
                'EnergyCost',
                'ElectricUse']

    def as_dict(self):
        return {'FileName': self.filename,
                'Result': self.exit_code_text(),
                'RatingPoints': self.rating_points(),
                'EnergyCost': self.energy_cost(),
                'ElectricUse': self.electric_use()}

    def exit_code_text(self):
        exit_code_lookup = {0: 'Success',
                            1: 'CalculatedWithValidationErrors',
                            2: 'CalculatedWithValidationWarnings',
                            10: 'CalculationError',
                            20: 'InvalidInputFile',
                            21: 'InvalidLibrary',
                            22: 'ErrorCreatingOutputFile',
                            29: 'OtherProcessingError'}
        if self.exitCode in exit_code_lookup:
            return exit_code_lookup[self.exitCode]
        else:
            return self.exitCode

    def rating_points(self):
        if self.tree:
            try:
                return self.tree.xpath('//RateResults/RatingPoints')[0].text
            except:
                pass
        return None

    def energy_cost(self):
        if self.tree:
            try:
                return self.tree.xpath('//EnrgResults/EnergyCost')[0].text.replace(',', '')
            except:
                pass
        return None

    def electric_use(self):
        if self.tree:
            try:
                return self.tree.xpath('//AnnualQuantityFuel/Electric')[0].text.replace(',', '')
            except:
                pass
        return None

with open(os.path.join(outputDirectory, 'Results.csv'), 'wb+') as outputCSV:
    outputWriter = csv.DictWriter(outputCSV, CalculationResults.as_keys())
    outputWriter.writeheader()

    for path, dirs, files in os.walk(inputDirectory):
        for fileName in files:
            fileRoot, fileExtension = os.path.splitext(fileName)
            if fileExtension in ['.hm2', '.hom', '.hmc']:
                results = CalculationResults(os.path.join(path, fileName))
                outputWriter.writerow(results.as_dict())
