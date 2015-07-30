#!/usr/bin/env python2

"""
Address parser. This program is highly reliant on zipcodes to determine what kind of parsing it applies.

This is how the program deals with the addresses.
    1. Splits the address into two where the semicolon is.
    2. Preserves the front part and checks to see if the 1st section is an address. If it contains traces of "State",
       "City", etc, it runs the first part through the regional parse method.
    3. The second part of the address is run through the regional parse method. This will separate out the State, City,
       Country, and other regional aspects of the address, and add that data to a dictionary, which will be returned to
       the main method.
    4. The street address is added into the dictionary, and the dictionary's information is parsed into an array and
       dropped into the CSV writer.
"""

import datetime
import os
import usaddress
import xlrd
import csv
import intaddress
from fuzzywuzzy import fuzz
import re


def main():
    fo = Folder(os.path.expanduser('~') + "/Desktop/", "Source/", "Target/", "Problem/")
    cycle(fo)


# Cycles through every single folder in the path, converting each file to an excel file.
def cycle(folder):
    file_names = folder.seek()
    print('Cycle start')
    for name in file_names:
        data = read("%s%s" % (folder.source, name))
        stuff, problem = rewrite(data)
        write(stuff, folder.target, os.path.splitext(name)[0])
        write(problem, folder.problem, "Problem - %s" % os.path.splitext(name)[0])
        print("Cycle finished")

    for file_data in file_names:
        os.remove("%s%s" % (folder.source, file_data))


# Rewrites the data held within the read array.
def rewrite(stuff):
    # The array containing all of the good data
    data = [stuff[0]]
    # The array holding all the data that needs to be checked later
    problem = []

    # Cycle through each of the rows in the data set
    for irow in range(1, len(stuff)):
        # Creates an Address object every time this process gets cycled
        addr = Address(stuff, irow)

        # Case 1: The Address Lane is blank
        if addr.addr == '':
            data.append(stuff[irow])
        # Case 2: The Address Lane is populated with some kind of data.
        else:
            try:
                new = addr.parse_it()
                # Mirror whatever data is already held inside the original row.
                temp = stuff[irow]
                while len(temp) < 8:
                    temp.append('')
                # Special cases that will only get called in special cases, hence the name, special cases
                if new['City'] == '' and 'Recipient' in new:
                    new['City'] = new['Recipient']

                if new['City'] == '' and 'LandmarkName' in new:
                    new['City'] = new['LandmarkName']
                elif 'LandmarkName' in new:
                    new['Street'] = "%s %s" % (new["LandmarkName"], new['Street'])

                # Fields that will later be parsed into the csv file.
                temp[2] = new['Street']
                temp[4] = new['City']
                temp[5] = new['State']
                temp[7] = new['Zipcode']

                if 'RegionName' in new:
                    temp[5] = new['RegionName']

                if 'Country' in new:
                    temp[6] = new['Country']

                # Special case that adds on the PO Number detected on the other half of the address semicolon.
                if 'PO' in new:
                    temp[3] = ("%s %s" % (temp[3], new['PO'])).strip()
                if 'Address2' in new:
                    temp[3] = ("%s %s" % (temp[3], new['Address2'])).strip()


                # Error counter: It tosses out any files that contain a suspiciously low amount of data.
                original_length = len(addr.addr.replace(' ', ''))
                cumulative_length = 0
                for str_index in range(2, len(temp)):
                    try:
                        cumulative_length += len(temp[str_index])
                    except TypeError:
                        pass

                if cumulative_length < 0.7 * original_length:
                    problem.append(stuff[irow])
                else:
                    data.append(temp)
            except ValueError:
                problem.append(stuff[irow])

    return data, problem


# Address object, contains everything that I do with an address, including parsing it.
class Address:
    # Init just creates a variable that holds the address being parsed.
    def __init__(self, data, index):
        self.addr = data[index][1]

    # Main function that parses the address held inside.
    def parse_it(self):
        # Rend separates the file based on semicolons. If there is only one side, or if there are no semicolons, it
        # tries to categorize it and returns the objects in the places it deems appropriate.
        street, region = self.rend()

        # Regional parse. Separates out the city, state, etc.
        data = self.region_parse(region)

        data['Street'] = "%s %s" % (data.setdefault('Street', ''), street)

        return data

    """Region parse takes the 2nd part of an address (as determined by the 'rend' method), and runs it through a
    series of tests to figure out if it is a domestic or international address. Each time the address passes a test,
    a little bit of address is pulled out of the main address string, and is appended to a dictionary that will be
    passed along. This makes the parser more accurate, since there is less "odd floating text".

    What it will first do is run the address through a filter that pulls out all PO Box numbers (which then get slotted
    into the dictionary). Next, it scans to see if there are any 'Dublin' zip codes (It's a special case, so I didn't
    include it in the main regex), and does the same thing it did to the PO Box numbers. Then, it runs the address
    through a regex that searches for US zip codes. If a domestic zip code is located, it is passed through the
    domestic parser. Else, the address is run through the international zip code regex, which pulls out the
    international zip codes. If there are no international zip codes located, it is simply run through the domestic
    parser."""
    def region_parse(self, region):
        # Scans for international zip codes, using the regex forms stored in the array below.
        regex_base = ['NSW \d{4}','\W(GIR|[A-Z]\d[A-Z\d]??|[A-Z]{2}\d[A-Z\d]??)[ ]??(\d[A-Z]{2})\W',
                      '\W((?:0[1-46-9]\d{3})|(?:[1-357-9]\d{4})|(?:[4][0-24-9]\d{3})|(?:[6][013-9]\d{3}))\W',
                      '\W([ABCEGHJKLMNPRSTVXY]\d[ABCEGHJKLMNPRSTVWXYZ])\ {0,1}(\d[ABCEGHJKLMNPRSTVWXYZ]\d)\W',
                      '\W(F-)?((2[A|B])|[0-9]{2})[0-9]{3}\W', '\W(V-|I-)?[0-9]{5}\W', '\W[^\W\d_]{2}-\d{4}\W',
                      '\W\d{6}\W',
                      '\W(0[289][0-9]{2})|([1345689][0-9]{3})|(2[0-8][0-9]{2})|(290[0-9])|(291[0-4])|(7[0-4][0-9]{2})|(7[8-9][0-9]{2})\W',
                      '\W[1-9][0-9]{3}\s?([a-zA-Z]{2})?\W', '\W([1-9]{2}|[0-9][1-9]|[1-9][0-9])[0-9]{3}\W',
                      '\W([D-d][K-k])?( |-)?[1-9]{1}[0-9]{3}\W', '\W(s-|S-){0,1}[0-9]{3}\s?[0-9]{2}\W',
                      '\W[1-9]{1}[0-9]{3}\W',
                      '\W[^\W\d_]{2}\d \d[^\W\d_]{2}\W', '\W[^\W\d_]{2}/d[^\W\d_] /d[^\W\d_]{2}\W', '\W\d{3}-\d{4}\W',
                      '\W\d{3}-\d{3}\W', '\W[^\W\d_]{2}\d \d[^\W\d_]\W',
                      '\W[^\W\d_]\d{6}\W', '\W[^\d_]{2}\d\W', '\W[^\W\d_]{2}-\d{4}\W',
                      '\W[^\W\d_]{2}\d \d[^\W\d_]{2}\W', '\W[^\W\d_]{3} \d{4}\W', '\W[^\W\d_]-\d{4}\W',
                      '\W[^\W\d_]\d{2} [^\W\d_]{3}\W']

        # Initialize this with blank values but valid keys, so I can just break out of this method while returning the
        # appropriate data structure if I need to.
        filtered_region = {'City': '',
                           'State': '',
                           'Zipcode': ''}

        # Initialized as an empty variable so that the method can also use it as a boolean if it is never called
        zippy = None

        # Must go before basic parsing, or else blank spaces will throw it off.
        if region is None:
            return filtered_region

        region = ' '.join(region.split())

        # Scanning for PO numbers in the latter part of the address. If one is found, the same procedure as
        # the zipcodes is followed.
        try:
            holder = usaddress.tag(region)
            if 'USPSBoxType' in holder[0] and 'USPSBoxID' in holder[0]:
                po = "%s %s" % (holder[0]['USPSBoxType'], holder[0]['USPSBoxID'])
                region = region.replace(po, '')
                filtered_region['PO'] = po
        except usaddress.RepeatedLabelError:
            pass

        dublin = re.search("dublin \d(\d)?", region.lower())
        if dublin:
            dubb = dublin.group(0)
            dub = dubb.split(' ')
            filtered_region['City'] = dub[0].capitalize()
            filtered_region['Zipcode'] = dub[1]
            region = ((region.lower()).replace(dubb, ',')).capitalize()
            return self.international(region, filtered_region)

        result = re.search("\W\d{5}([\-]?\d{4})?\W", " %s " % region)

        if not result:
            # Finds the international zipcodes, and then strips them from the region string
            for reg in regex_base:
                print(region)
                region = ' '.join(region.split())
                results = re.search(reg, " %s " % region)
                if results:
                    zippy = (results.group(0)).strip()
                    filtered_region['Zipcode'] = zippy
                    # print(zippy)
                    region = region.replace(zippy, ',')
                    break
            if not zippy:
                # Back to American Parsing. These addresses that make it down here should all be domestic.
                filtered_region = self.domestic(region, filtered_region)
                return filtered_region
        else:
            # American Parsing. These addresses that make it down here should all be domestic.
            filtered_region = self.domestic(region, filtered_region)
            return filtered_region

        # International Parsing. Special cases for them.
        if zippy:
            return self.international(region, filtered_region)

        return filtered_region

    """Parses out international regional codes, once the addresses have been identified as an international address.
    This function cuts to the chase by passing the address through the international address parser, which separates
    the address into parts. The first part will scan for a misplaced address, similar to what the 'domestic' method
    does, and then it will place it into a dictionary if the street address exists. This method does the same for
    "state", "city", "country", and "region"."""
    @staticmethod
    def international(region, filtered_region):
        try:
            inter_parse = intaddress.tag(region)

            address = [inter_parse.setdefault('AddressNumber', ''),
                       inter_parse.setdefault('StreetNamePreDirectional', ''),
                       inter_parse.setdefault('StreetNamePreModifier', ''), inter_parse.setdefault('StreetName', ''),
                       inter_parse.setdefault('StreetNamePostType', ''),
                       inter_parse.setdefault('StreetNamePostDirectional', '')]
            if address is not None:
                address = ' '.join(address)
                address = ' '.join(address.split())
                filtered_region['Street'] = address

            if 'CityName' in inter_parse:
                if filtered_region['City'] == '':
                    filtered_region['City'] = inter_parse['CityName']
                elif filtered_region['State'] == '':
                    filtered_region['State'] = inter_parse['CityName']
                else:
                    filtered_region['Address2'] = inter_parse['CityName']

            if 'StateName' in inter_parse:
                filtered_region['State'] = inter_parse['StateName']
            if 'CountryName' in inter_parse:
                filtered_region['Country'] = inter_parse['CountryName']
            if 'RegionName' in inter_parse:
                filtered_region['RegionName'] = inter_parse['RegionName']
        except:
            raise ValueError
        return filtered_region


    """Parsing for domestic elements. First scans to see if there are any spaces or commas in the address field. If
    neither of those exist, the element is most likely a city name. Next, it scans to see if Washington D.C. is in the
    string, or anything similar to that. If the similarity ratio exceeds 90%, then the program replaces the DC string
    with a comma and stores DC as the city name. After that, it simply places the address into the address parser,
    which separates the address into its elements and is parsed into a dictionary. Recipient and LandmarkName are both
    key values that I include in the dictionary, because sometime this program mistakes address elements for those.
    This way, this catches as much information as possible."""
    @staticmethod
    def domestic(region, filtered_region):
        region = region.strip()
        if not ' ' in region and not ',' in region:
            filtered_region['City'] = region
            return filtered_region

        citi = None
        # Because Washington DC has some odd formatting and weird state issues (in maryland, but not technically in
        # maryland), I just created a special case that deals with this string. Uses fuzzy string reading to ID.
        if fuzz.partial_ratio(region, 'Washington D.C') > 90:
            citi = 'Washington, D.C.'
            region = region.replace('Washington', '')
            region = region.replace('DC', '')
            region = region.replace('D.C.', '')

        # The meat of the US address parsing. If the parsing returns an error, nothing is returned and it will end up
        # in the problem file stack.
        try:
            region = region.strip(',')
            clean = usaddress.tag(region)
        except usaddress.RepeatedLabelError:
            return filtered_region

        # Street addresses returned by the parser may contain some or none of these elements, which is why they are
        # initially placed into an array to be combined later. This snippet of code is here, in addition to the first
        # section so that it can catch any misplaced addresses.
        address = [clean[0].setdefault('AddressNumber', ''),
                   clean[0].setdefault('StreetNamePreDirectional', ''),
                   clean[0].setdefault('StreetNamePreModifier', ''), clean[0].setdefault('StreetName', ''),
                   clean[0].setdefault('StreetNamePostType', ''),
                   clean[0].setdefault('StreetNamePostDirectional', '')]
        if address is not None:
            address = ' '.join(address)
            address = ' '.join(address.split())
            filtered_region['Street'] = address

        # Reassigns variable values to the dictionary. If these values didn't exist in the firm place within the
        # parser, nothing is done. The setdefault function only works with dictionaries.
        filtered_region['State'] = clean[0].setdefault('StateName', '')
        filtered_region['Zipcode'] = clean[0].setdefault('ZipCode', '')

        # Enters Washington DC
        if citi:
            filtered_region['City'] = citi
        else:
            filtered_region['City'] = clean[0].setdefault('PlaceName', '')

        # Adds on the PO box stuff into the dictionary.
        if 'USPSBoxType' in clean[0] and 'USPSBoxID' in clean[0]:
            filtered_region['PO'] = "%s %s" % (clean[0]['USPSBoxType'], clean[0]['USPSBoxID'])

        if 'CountryName' in clean[0]:
            filtered_region['Country'] = clean[0]['CountryName']

        if 'Recipient' in clean[0]:
            filtered_region['Recipient'] = clean[0]['Recipient']

        if 'LandmarkName' in clean[0]:
            filtered_region['LandmarkName'] = clean[0]['LandmarkName']

        return filtered_region

    """Splits the address by the semicolon in the middle. If the semicolon does not exist, the entire address is just
    assumed to be a street address. Alternatively, this has a case catch that will scan to see if there are any
    misplaced elements and it will attempt to place them in the right place."""
    def rend(self):
        if not self.addr.find(';') == -1:
            temp = self.addr.split(';')
            if temp[1] == '':
                try:
                    if 'PlaceName' not in usaddress.tag(temp[0])[0] or 'StateName' not in usaddress.tag(temp[0])[0]:
                        return temp[0], None
                    else:
                        return '', temp[0]
                except usaddress.RepeatedLabelError:
                    return temp[0], ''
            return temp[0], temp[1]
        else:
            return self.addr, ''


# Reads the Excel file. Currently only works for Excel, if support is needed for txt files, tsv, etc, I need to add
# in another snippet later.
def read(file_data):
    data = []
    book = xlrd.open_workbook(file_data)
    for wsnum in range(0, book.nsheets):
        ws = book.sheet_by_index(wsnum)
        if wsnum == 0:
            start = 0
        else:
            start = 1

        for rows in range(start, ws.nrows):
            data.append(ws.row_values(rows))
    return data


# Redoes the encoding for each element in the table. Turns it into Unicode text, so the CSV reader can handle it.
def uniform(dictionary_input):
    if isinstance(dictionary_input, dict):
        return {uniform(key): uniform(value) for key, value in dictionary_input.iteritems()}
    elif isinstance(dictionary_input, list):
        return [uniform(element) for element in dictionary_input]
    elif isinstance(dictionary_input, unicode):
        return dictionary_input.encode('utf-8')
    else:
        return dictionary_input


# Function that writes the CSV.
def write(data, pathway, name):
    data = uniform(data)
    now = datetime.date.today().strftime("%m.%d.%y")

    with open("%s/%s (%s).txt" % (pathway, name, now), 'w') as f:
        writer = csv.writer(f, delimiter='|')
        writer.writerows(data)


# Folder class. Deals with everything that has to do with the file locations.
class Folder:
    def __init__(self, home, source, target, problem):
        self.home = home
        self.source = "%s%s" % (home, source)
        self.target = self.direct(target)
        self.problem = self.direct(problem)

    # Creates a new path to store the created excel files in.
    @staticmethod
    def direct(new_folder):
        home = os.path.expanduser('~')
        pathway = "%s/Desktop/%s" % (home, new_folder)
        if not os.path.isdir(pathway):
            os.mkdir(pathway)
        return pathway

    def seek(self):
        files = os.listdir(self.source)
        return files


# Runs the main, after establishing that this is not a library.
if __name__ == "__main__":
    main()
