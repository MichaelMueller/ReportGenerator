import argparse
import api
import logging

# args
parser = argparse.ArgumentParser(
    description='Create the installer package')
parser.add_argument('--log_level', type=int, default=logging.INFO,
                    help='log level (CRITICAL = 50, ERROR = 40, WARNING = 30, INFO = 20, DEBUG = 10, NOTSET = 0')
parser.add_argument('--log_file', type=str, default=None, help='log file')

args = parser.parse_args()
api.create_installer(args.log_level, args.log_file)
