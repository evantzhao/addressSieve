#!/usr/bin/env python2
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

    for x in range(0, len(file_names)):
        print('Cycle start')
        for name in file_names:
            data = read("%s%s" % (folder.source, name))
            stuff, problem = rewrite(data)
            # write(stuff, folder.target, os.path.splitext(name)[0])
            write(stuff, folder.target, 'cross')
            write(problem, folder.problem, os.path.splitext(name)[0])
            print("Cycle fin")

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
                    temp[3] = "%s %s" % (temp[3], new['PO'])


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
                print("HI")
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

    # Will output another dictionary, combines with other one using update method.
    def address_parse(self, address):
        stuff = {'Street': address}

        if address is None:
            return stuff
        try:
            parser = usaddress.tag(address)

            address = [parser[0].setdefault('AddressNumber', ''), parser[0].setdefault('StreetNamePreDirectional', ''),
                       parser[0].setdefault('StreetNamePreModifier', ''), parser[0].setdefault('StreetName', ''),
                       parser[0].setdefault('StreetNamePostType', ''), parser[0].setdefault('StreetNamePostDirectional', '')]
            if address is not None:
                address = ' '.join(address)
                address = ' '.join(address.split())

            stuff['Street'] = address
        except usaddress.RepeatedLabelError:
            return stuff

        print(parser)
        return stuff



    # Does most of the parsing work.
    def region_parse(self, region):
        # Scans for international zip
        # codes, using the regex forms stored in the array below.

        regex_base = ['\W(GIR|[A-Z]\d[A-Z\d]??|[A-Z]{2}\d[A-Z\d]??)[ ]??(\d[A-Z]{2})',
                      '\W((?:0[1-46-9]\d{3})|(?:[1-357-9]\d{4})|(?:[4][0-24-9]\d{3})|(?:[6][013-9]\d{3}))',
                      '\W([ABCEGHJKLMNPRSTVXY]\d[ABCEGHJKLMNPRSTVWXYZ])\ {0,1}(\d[ABCEGHJKLMNPRSTVWXYZ]\d)',
                      '\W(F-)?((2[A|B])|[0-9]{2})[0-9]{3}', '\W(V-|I-)?[0-9]{5}',
                      '\W(0[289][0-9]{2})|([1345689][0-9]{3})|(2[0-8][0-9]{2})|(290[0-9])|(291[0-4])|(7[0-4][0-9]{2})|(7[8-9][0-9]{2})',
                      '\W[1-9][0-9]{3}\s?([a-zA-Z]{2})?', '([1-9]{2}|[0-9][1-9]|[1-9][0-9])[0-9]{3}',
                      '\W([D-d][K-k])?( |-)?[1-9]{1}[0-9]{3}', '(s-|S-){0,1}[0-9]{3}\s?[0-9]{2}', '[1-9]{1}[0-9]{3}$',
                      '\W[^\W\d_]{2}\d \d[^\W\d_]{2}', '[^\W\d_]{2}/d[^\W\d_] /d[^\W\d_]{2}', '\d{3}-\d{4}', '\d{3}-\d{3}'
                      '\d{6}', '[^\W\d_]{2}-\d{4}', '[^\W\d_]\d{6}', '\W[^\d_]{2}\d', '\W[^\W\d_]{2}-\d{4}',
                      '\W[^\W\d_]{2}\d \d[^\W\d_]{2}']

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

        result = re.search("\W\d{5}([\-]?\d{4})?", region)

        if not result:
            # Finds the international zipcodes, and then strips them from the region string
            for reg in regex_base:
                results = re.search(reg, " %s" % region)
                if results:
                    zippy = results.group(0)
                    filtered_region['Zipcode'] = zippy
                    # print(zippy)
                    region = region.replace(zippy, ',')
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

    @staticmethod
    def international(region, filtered_region):
        try:
            whee = intaddress.tag(region)

            address = [whee.setdefault('AddressNumber', ''), whee.setdefault('StreetNamePreDirectional', ''),
                       whee.setdefault('StreetNamePreModifier', ''), whee.setdefault('StreetName', ''),
                       whee.setdefault('StreetNamePostType', ''), whee.setdefault('StreetNamePostDirectional', '')]
            if address is not None:
                address = ' '.join(address)
                address = ' '.join(address.split())
            if address:
                filtered_region['Street'] = address
            if 'CityName' in whee:
                filtered_region['City'] = whee['CityName']
            if 'StateName' in whee:
                filtered_region['State'] = whee['StateName']
            if 'CountryName' in whee:
                filtered_region['Country'] = whee['CountryName']
            if 'RegionName' in whee:
                filtered_region['RegionName'] = whee['RegionName']

        except:
            print("HIEJ")
            raise ValueError

        print(filtered_region)

        return filtered_region

    @staticmethod
    def domestic(region, filtered_region):
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

        # Reassigns variable values to the dictionary. If these values didn't exist in the firm place within the
        # parser, nothing is done. The setdefault function only works with dictionaries.
        filtered_region['State'] = clean[0].setdefault('StateName', '')
        filtered_region['Zipcode'] = clean[0].setdefault('ZipCode', '')

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

    def rend(self):
        if not self.addr.find(';') == -1:
            temp = self.addr.split(';')
            if temp[1] == '':
                try:
                    if 'PlaceName' not in usaddress.tag(temp[0])[0] or 'StateName' not in usaddress.tag(temp[0])[0]:
                        return temp[0], None
                    else:
                        return None, temp[0]
                except usaddress.RepeatedLabelError:
                    return temp[0], ''
            return temp[0], temp[1]
        else:
            return self.addr, ''

    @staticmethod
    def merge(arr, index, end):
        string = ''
        while index <= end:
            string = string + ' ' + arr[index]
            index += 1
        string = string.strip()
        return string


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


def uniform(dictionary_input):
    if isinstance(dictionary_input, dict):
        return {uniform(key): uniform(value) for key, value in dictionary_input.iteritems()}
    elif isinstance(dictionary_input, list):
        return [uniform(element) for element in dictionary_input]
    elif isinstance(dictionary_input, unicode):
        return dictionary_input.encode('utf-8')
    else:
        return dictionary_input


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
