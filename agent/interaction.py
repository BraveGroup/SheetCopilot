import argparse, yaml, asyncio,os
from Agent import Agent

os.environ["http_proxy"] = "http://127.0.0.1:7890"
os.environ["https_proxy"] = "http://127.0.0.1:7890"

parser = argparse.ArgumentParser(description='Process config.')
parser.add_argument('--config', '-c', type=str, help='path to config file')
args = parser.parse_args()

with open(args.config, 'r') as f:
    config = yaml.load(f, Loader=yaml.Loader)

config['interaction_mode'] = True
agent = Agent(config)
while True:
    instruction = input('Enter your instruction: \n')
    asyncio.run(agent.Instruction2('', instruction))