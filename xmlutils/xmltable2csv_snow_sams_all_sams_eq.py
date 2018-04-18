"""
    xmltable2csv.py
    Yigal Lazarev, http://yig.al
    May 2015
    
    License:        MIT License
    Documentation:    http://nadh.in/code/xmlutils.py
"""

import os
import codecs
import xml.etree.ElementTree as ETree

import pdb

class xmltable2csv:
    """
    This class is intended to convert tables formatted as XML document, to a
    comma-separated value lines (CSV) file.

    This is a bit different than the xml2csv tool, which tries to convey the XML hierarchy
    into a CSV file - it keeps descending to the selected tags child nodes and translates these as well.

    Example for the expected input to this converter class:
    =======================================================

    A table of the following form:

    Header 1           Header 2
    Value R1C1         Value R1C2
    Value R2C1         Value R2C2

    Will be formatted something along the lines of the following XML in Microsoft Excel:

     <?xml version="1.0" encoding="utf-8"?>
     <?mso-application progid="Excel.Sheet"?>
     <ss:Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
     <ss:Table ss:ExpandedColumnCount="11" ss:FullColumns="1" ss:ExpandedRowCount="28" ss:FullRows="1">
      <ss:Row>
        <ss:Cell ss:StyleID="HeaderStyle">
          <ss:Data ss:Type="String">Header 1</ss:Data>
        </ss:Cell>
        <ss:Cell ss:StyleID="HeaderStyle">
          <ss:Data ss:Type="String">Header 2</ss:Data>
        </ss:Cell>
      </ss:Row>
      <ss:Row>
        <ss:Cell>
          <ss:Data ss:Type="String">Value R1C1</ss:Data>
        </ss:Cell>
        <ss:Cell>
          <ss:Data ss:Type="String">Value R1C2</ss:Data>
        </ss:Cell>
      </ss:Row>
      <ss:Row>
        <ss:Cell>
          <ss:Data ss:Type="String">Value R2C1</ss:Data>
        </ss:Cell>
        <ss:Cell>
          <ss:Data ss:Type="String">Value R2C2</ss:Data>
        </ss:Cell>
      </ss:Row>
     </ss:Table>
     </ss:Workbook>

    This might be a bit different in later versions, but the general form is the same. Notice that
    the tags are namespaced, and this namespacing might be somewhat obfuscated, in the form of a xmlns
    property in the containing 'Workbook' tag.

    This class converts simple (not tested with XLSX sheets containing formulas etc) XML-formatted tables
    to csv, regardless of the specific tagging and hierarchy structure.

    Tested with some XLSX files and worked fine even for files that wouldn't convert in tools
    such as dilshod's xlsx2csv.
    """

    def __init__(self, input_file, output_file, encoding='utf-8'):
        """Initialize the class with the paths to the input xml file
        and the output csv file

        Keyword arguments:
        input_file -- input xml filename
        output_file -- output csv filename
        encoding -- character encoding
        """

        self.output_filename = output_file
        self.input_directory = input_file
        self.output_buffer = []
        self.output_dict = [{}]
        self.item_titles = {}
        self.output = None

        self._init_output_dict()

        # open the xml file for iteration
        # self.context = ETree.iterparse(input_file, events=("start", "end"))

        # output file handle
        try:
            pass
            #self.output = codecs.open(output_file, "w", encoding=encoding)
        except:
            print("Failed to open the output file")
            raise

    def convert(self, tag="Data", delimiter=",", noheader=False,
                limit=-1, buffer_size=1000):

        """Convert the XML table file to CSV file

            Keyword arguments:
            tag -- the record tag that contains a single entry's text. eg: Data (Microsoft XLSX)
            delimiter -- csv field delimiter
            limit -- maximum number of records to process
            buffer -- number of records to keep in buffer before writing to disk

            Returns:
            number of records converted
        """
        file_ctr = 0
        item_ctr = 0
        for dirName, subdirList, fileList in os.walk(self.input_directory):
            print('Found directory: %s' % dirName)
            for fname in fileList:
                print('\t%s' % fname)
                # open the xml file for iteration
                if not fname.endswith(".xml"):
                    continue
                #pdb.set_trace()
                
                input_file = dirName + "/" + fname
                self.context = ETree.iterparse(input_file, events=("start", "end"))

                # iterate through the xml
                items = [{}]

                depth = 0
                min_depth = 0
                row_depth = -1
                n = 0
                for event, elem in self.context:
                    if event == "start":
                        depth += 1
                        continue
                    else:
                        depth -= 1
                        if depth < min_depth:
                            min_depth = depth

                    if depth < row_depth and items:
                        if noheader:
                            noheader = False
                        else:
                            # new line
                            self.output_buffer.append(items)
                        items = []
                        # flush buffer to disk
                        if len(self.output_buffer) > buffer_size:
                         self._write_buffer(delimiter)

                    plain_tag = elem.tag
                    last_delim = max(elem.tag.rfind('}'), elem.tag.rfind(':'))
                    if 0 < last_delim < len(elem.tag) - 1:
                        plain_tag = elem.tag[last_delim + 1:]
                    if tag == plain_tag:
                        if n == 0:
                            min_depth = depth
                        elif n == 1:
                            row_depth = min_depth
                        n += 1
                        if 0 < limit < n:
                            break
                        elem_name = elem.get("name")
                        if elem_name in self.output_dict[0].keys():
                            # if  elem_name == 'SamS.ArchivedURL':
                            #     if hash(elem.text) in self.item_titles.keys() and self.item_titles[hash(elem.text)] == elem.text:
                            #         #item is repetative
                            #         self.output_dict[item_ctr]={}
                            #         #item_ctr-=1
                            #         break
                            #     else:
                            #         self.item_titles[hash(elem.text)] = elem.text
                            self.output_dict[item_ctr][elem_name]= elem.text and elem.text.encode('utf8') or ''

                # #if (len(self.output_dict[item_ctr]) > 0 ) :
                # if ('SamS.ArchivedURL' in self.output_dict[item_ctr]):
                #    item_ctr+=1
                #    self.output_dict.append({})
                # else:
                #    self.output_dict[item_ctr] = {}
                item_ctr+=1
                self.output_dict.append({})
                file_ctr+=1 #next row in the dictionary array
                print "processing file no ", file_ctr, " item no", item_ctr

        pdb.set_trace()
        self._drop_repetition()
        self._write_buffer(delimiter)  # write rest of the buffer to file

        return n

    def _drop_repetition(self):
        """
        Drop any second item which has exactly the Sams field as a 
        a previous on.
        """
        no_repeat_list = [self.output_dict[0]]
        sams_hash_table =  {}
        for cur_item in self.output_dict[1:]:
            #make the string to be hashed by concatinating all sams values
            sams_hash_input = ""
            for cur_element in cur_item:
                if cur_element[:len("SamS.")]=="SamS.":
                    sams_hash_input += cur_item[cur_element]

            #check if we have seen this item already
            if hash(sams_hash_input) in  sams_hash_table:
                print "found repeated record: ", cur_item
            else:
                no_repeat_list.append(cur_item)
                sams_hash_table[hash(sams_hash_input)] = 1

        print "from ", len(self.output_dict) - 1, "items , ", len(no_repeat_list) - 1, " have unique collective Sams fields value"
        self.output_dict = no_repeat_list
            
    def _write_buffer(self, delimiter):
        """Write records from buffer to the output file"""
        import csv

        keys = self.output_dict[0].keys()
        with open(self.output_filename, 'wb') as output_file:
            dict_writer = csv.DictWriter(output_file, keys)
            dict_writer.writeheader()
            dict_writer.writerows(self.output_dict)

        #self.output.write('\n'.join([delimiter.join(e) for e in self.output_buffer]) + '\n')
        #self.output_buffer = []

    def _init_output_dict(self):
        self.output_dict[0] = {'Language':'', 'Encoding':'', 'URL':'', 'Source':'', 'SourceFile':'', 'FileFormat':'', 'NumPages':'', 'SamS.ArchivedURL':'', 'SamS.Classification':'', 'SamS.Creator':'', 'SamS.Creator^Office':'', 'SamS.DateCreated':'', 'SamS.DateCreated^Circa':'', 'SamS.DatePublished':'', 'SamS.Description':'', 'SamS.Distribution':'', 'SamS.Genre':'', 'SamS.Publication Link':'', 'SamS.Publisher':'', 'SamS.Redactions':'', 'SamS.Relation':'', 'SamS.Reporter':'', 'SamS.Subject':'', 'SamS.Surveillance Program':'', 'SamS.Target':'', 'SamS.Title':'', 'Identifier':'', 'SamS.Relation':'', 'ex.File.FileName':''}
