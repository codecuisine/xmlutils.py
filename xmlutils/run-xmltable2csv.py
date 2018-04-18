import argparse
from xml2csv import xml2csv
from xmltable2csv_snow_sams_all_sams_eq import xmltable2csv

def run_xml2csv():
	print """xml2csv by Kailash Nadh (http://nadh.in)
	--help for help

	"""

	# parse arguments
	parser = argparse.ArgumentParser(description='Convert an xml file to csv format.')
	parser.add_argument('--input', dest='input_file', required=True, help='input xml filename')
	parser.add_argument('--output', dest='output_file', required=True, help='output csv filename')
	parser.add_argument('--tag', dest='tag', required=True, help='the record tag. eg: item')
	parser.add_argument('--delimiter', dest='delimiter', default=',', help='delimiter character. (default=,)')
	parser.add_argument('--ignore', dest='ignore', default='', nargs='+', help='list of tags to ignore')
	parser.add_argument('--noheader', dest='noheader', action='store_true', help='exclude csv header (default=False)')
	parser.add_argument('--encoding', dest='encoding', default='utf-8', help='character encoding (default=utf-8)')
	parser.add_argument('--limit', type=int, dest='limit', default=-1, help='maximum number of records to process')
	parser.add_argument('--buffer_size', type=int, dest='buffer_size', default='1000',
						help='number of records to keep in buffer before writing to disk (default=1000)')
	parser.add_argument('--noquotes', dest='noquotes', action='store_true', help='no quotes around values')

	args = parser.parse_args()

	converter = xml2csv(args.input_file, args.output_file, args.encoding)
	num = converter.convert(tag=args.tag, delimiter=args.delimiter, ignore=args.ignore,
							noheader=args.noheader, limit=args.limit, buffer_size=args.buffer_size,
							quotes=not args.noquotes)

	print "\n\nWrote", num, "records to", args.output_file

def run_xmltable2csv():
	print """xmltable2csv by Yigal Lazarev (http://yig.al)
	--help for help

	"""

	# parse arguments
	parser = argparse.ArgumentParser(description='Convert an xml file to csv format.')
	parser.add_argument('--input', dest='input_file', required=True, help='input xml filename')
	parser.add_argument('--output', dest='output_file', required=True, help='output csv filename')
	parser.add_argument('--tag', dest='tag', required=True, help='the record tag. eg: Data')
	parser.add_argument('--delimiter', dest='delimiter', default=',', help='delimiter character. (default=,)')
	parser.add_argument('--noheader', dest='noheader', action='store_true', help='exclude csv header (default=False)')
	parser.add_argument('--encoding', dest='encoding', default='utf-8', help='character encoding (default=utf-8)')
	parser.add_argument('--limit', type=int, dest='limit', default=-1, help='maximum number of records to process')
	parser.add_argument('--buffer_size', type=int, dest='buffer_size', default='1000',
						help='number of records to keep in buffer before writing to disk (default=1000)')

	args = parser.parse_args()

	converter = xmltable2csv(args.input_file, args.output_file, args.encoding)
	num = converter.convert(tag=args.tag, delimiter=args.delimiter,
							noheader=args.noheader, limit=args.limit, buffer_size=args.buffer_size)

	print "\n\nWrote", num, "records to", args.output_file

if __name__ == "__main__":
    run_xmltable2csv()

