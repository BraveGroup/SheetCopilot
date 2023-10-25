from utils import ChatGPT
from Agent import Agent
import asyncio
import yaml, json
import time, os, shutil, datetime
import pandas as pd
import argparse

parser = argparse.ArgumentParser(description='Process config.')
parser.add_argument('--config', '-c', type=str, help='path to config file')
args = parser.parse_args()

with open(args.config, 'r') as f:
    config = yaml.load(f, Loader=yaml.Loader)

if not os.path.exists(config['path']['save_path']):
    os.makedirs(config['path']['save_path'])

index_path = os.path.join(config['path']['save_path'], 'index.txt')
if not os.path.exists(index_path):
    with open(index_path, 'w') as f:
        f.write('')
with open(os.path.join(config['path']['save_path'], 'index.txt'), 'r') as file:
    index_list = file.readlines()
    index_list = [int(line.strip()) for line in index_list]
    print("Loading checkpoint for {}. Tasks {} have been processed.".format(config['path']['save_path'], ', '.join(str(x) for x in index_list)))
task_queue = asyncio.Queue()

# task_list = []
# with open(config['path']['task_path'], 'r', encoding='utf-8-sig') as f:
#     for line in f:
#         task_list.append(json.loads(line))
# task_list = task_list[:]

task_df = pd.read_excel(config['path']['task_path'], header=0)

async def producer():
    for index, row in task_df.iterrows():
        if index+1 in index_list:
            continue
        source_path = os.path.join(config['path']['source_path'],row['Sheet Name']+'.xlsx')
        path = os.path.join(config['path']['save_path'],f"{index+1}_{row['Sheet Name']}")
        if not os.path.exists(path):
            os.makedirs(path)
            destination_path = os.path.join(path, f"{index+1}_{row['Sheet Name']}_source.xlsx")
            shutil.copy(source_path, destination_path)
        await task_queue.put((index+1, os.path.join(path, f"{index+1}_{row['Sheet Name']}_"), row['Context'], row['Instructions']))

async def worker():
    agent = Agent(config)
    while True:
        index, path, context, instructions = await task_queue.get()
       
        source_path = path + 'source'+'.xlsx'

        print("\033[0;36;40mProcessing Task {}: {}\033[0m\n".format(index, source_path))
        
        log = {
            'Source Path': os.path.abspath(source_path),
            'Context': context,
            'Instructions': instructions,
            'Success Response': [],
            'Fail Response': [],
            'Prompt_format': config['Agent']['ChatGPT_1']['prompt_format'],
            'Use oracle API doc': config['Agent']['use_oracle_API_doc'],
        }
        success_count = 0
        for i in range(config['repeat']):
            save_path = path + str(i+1)
            success, res = await agent.Instruction2(context, instructions, source_path, save_path)
            context_log_list = res.pop('context_log')
            if not os.path.exists(save_path):
                os.makedirs(save_path)
            for i, context_log in enumerate(context_log_list):
                with open(os.path.join(save_path, f'context_log_{i+1}.yaml'), 'w') as f:
                    f.write(yaml.dump(context_log))
            if success:
                success_count += 1
                log['Success Response'].append(res)
            else:
                log['Fail Response'].append(res)

        log['Success Count'] = success_count
        log['Total Count'] = config['repeat']
        log['Timestamp'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        print("\033[0;36;40mTask {} have been processed\033[0m\n".format(index))

        # check_consistency(save_path)
        with open(path + 'log.yaml', 'w') as f:
            f.write(yaml.dump(log))
        with open(os.path.join(config['path']['save_path'], 'index.txt'), 'a') as f:
            f.write(str(index)+'\n')
        remaining = task_queue.qsize()
        print(f"Remaining tasks: {remaining} - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        task_queue.task_done()

async def main():
    await producer()
    tasks = []
    for i in range(config['worker']):
        tasks.append(asyncio.create_task(worker()))

    await task_queue.join()

if __name__ == '__main__':
    start = time.time()
    asyncio.run(main())
    print("Time: {:.1f} min".format((time.time()-start) / 60))