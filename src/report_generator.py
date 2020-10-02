import argparse
import api
import logging

# args
parser = argparse.ArgumentParser(
    description='A utility to generate a nicely formatted DICOM PDF from a DICOM SR report using'
                ' a template document (like word etc)')
parser.add_argument('dicom_sr_file', type=str, help='actual DICOM SR report file')
parser.add_argument('config_file', type=str, help='actual json config file being used to produce the sr. refer to the manual')
parser.add_argument('--log_level', type=int, default=logging.WARN,
                    help='log level (CRITICAL = 50, ERROR = 40, WARNING = 30, INFO = 20, DEBUG = 10, NOTSET = 0')
parser.add_argument('--log_file', type=str, default=None, help='log file')

args = parser.parse_args()
api.generate_report(args.dicom_sr_file, args.config_file, args.log_level, args.log_file)
