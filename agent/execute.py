from utils import ChatGPT
from Agent import Agent
import asyncio
import yaml, json
import time, os, shutil, datetime
import pandas as pd

with open('config/config.yaml', 'r') as f:
    config = yaml.load(f, Loader=yaml.Loader)

# task_list = []
# with open(config['path']['task_path'], 'r', encoding='utf-8-sig') as f:
#     for line in f:
#         task_list.append(json.loads(line))

task_df = pd.read_excel(config['path']['task_path'], header=0)

agent = Agent(config['Agent'])

async def worker(index):
    row = task_df.iloc[index-1]
    source_path = os.path.join(config['path']['source_path'],row['Sheet Name']+'.xlsx')
    context = row['Context']
    instructions = row['Instructions']
    print('Context:', context)
    print('Instructions:', instructions)
    if agent._backend.activeWB is not None:
        agent._backend.activeWB.Close(SaveChanges=False)
    success, res = await agent.Instruction2(context, instructions, source_path)
    print('Success:', success)
    context_log = res.pop('context_log')
    for i, log in enumerate(context_log):
        with open(f'context_log/{index}_{i+1}.yaml', 'w') as f:
            f.write(yaml.dump(log))
    print(yaml.dump(res))
        
while True:
    index = int(input('index: '))
    asyncio.run(worker(index))