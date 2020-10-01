import argparse
import api
import logging
# args
parser = argparse.ArgumentParser(description='Report Generator')
parser.add_argument('--log_level', type=int, default=logging.INFO, help='the log level')
parser.add_argument('--log_file', type=str, default=None, help='the log file')

args = parser.parse_args()
api.setup_logging(args.log_level, args.log_file)

print("hello world")
