import json
import argparse
import os

import asyncio
import aiosparkapi
import questionbot

parser = argparse.ArgumentParser()
default_config_file = '{}/.config/questionbot/config.json'.format(os.path.expanduser('~'))
parser.add_argument(
    '--config',
    '-c',
    default=default_config_file,
    help='Path to configuration file. Default: {}'.format(default_config_file)
)

args = parser.parse_args()

with open(args.config, 'r') as fd:
    config = json.load(fd)

loop = asyncio.get_event_loop()
bot = questionbot.Bot(config, loop)
bot.run(loop)
