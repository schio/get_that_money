import json
import argparse

from dotmap import DotMap

def get_config_from_json():
	parser = argparse.ArgumentParser(description=__doc__)
	parser.add_argument(dest="config",
						metavar="config",
						default=None,
						help="json file for config")

	args = parser.parse_args()

	with open(args.config, 'r') as config_file:
		config_dict = json.load(config_file)

	return DotMap(config_dict)