import argparse, yaml, asyncio
from Agent.agent import Agent

parser = argparse.ArgumentParser(description='Process config.')
parser.add_argument('--config', '-c', type=str, default="./config/config.yaml", help='path to config file')
args = parser.parse_args()

with open(args.config, 'r') as f:
    config = yaml.load(f, Loader=yaml.Loader)

config['interaction_mode'] = True
agent = Agent(config)
while True:
    instruction = input('Enter your instruction: \n')
    asyncio.run(agent.Instruction('', instruction))