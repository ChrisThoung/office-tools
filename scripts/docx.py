#!/usr/bin/python3


import argparse


from officetools import __version__ as version


# Create top-level parser
parser = argparse.ArgumentParser(
    description='Microsoft Word document handler.')
parser.add_argument(
    '-V', '--version',
    action='version',
    version=version)

# Add sub-parsers
subparsers = parser.add_subparsers(
    title='commands',
    dest='command')

# Extract
parser_extract = subparsers.add_parser(
    'extract',
    help='extract elements from a document')
parser_extract.add_argument(
    '-o', '--output',
    nargs=1,
    metavar='OUTPUT',
    default=None,
    type=str,
    required=False,
    help='set output CSV file')
parser_extract.add_argument(
    'file',
    nargs=1,
    metavar='FILE',
    type=str,
    help='path of Word document to process')


# Main
if __name__ == '__main__':
    args = parser.parse_args()
    if args.command == 'extract':
        from officetools.docx import extract
        comments = extract.extract_comments(args.file[0])
        if args.output is not None:
            comments.to_csv(args.output[0])
