from Agent import Agent, xwBackend
import pandas as pd
import os
from utils import compare_workbooks
import threading
import pythoncom
import queue
import json
import time
import yaml
import asyncio

path = '../example_sheets_part1/Business/task_instructions.xlsx'
sheetdir_path = '../example_sheets_part1/Business'
savedir_path = '../example_sheets_part1/Business_Res'

result_queue = queue.Queue()
is_done = False

with open('config/config.yaml', 'r') as f:
    config = yaml.load(f, Loader=yaml.Loader)

async def worker():
    df = pd.read_excel(config['path']['tasks_path'], header=0)
    agent = Agent(config['Agent'])
    for index, row in df.iterrows():
        source_path = os.path.join(config['path']['source_path'],row['Sheet Name']+'.xlsx')
        save_path = os.path.join(config['path']['save_path'],str(index+1)+'_'+row['Sheet Name'])
        if not os.path.exists(save_path):
            os.makedirs(save_path)
        log = {
            'Source Path': os.path.abspath(source_path),
            'Context': row['Context'],
            'Instructions': row['Instructions'],
            'Success Response': [],
            'Fail Response': []
        }
        print('Processing {}...'.format(row['Sheet Name']))
        print('Context: {}'.format(row['Context']))
        print('Instructions: {}'.format(row['Instructions']))
        success_count = 0
        for i in range(config['repeat']):
            success, res = await agent.Instruction(row['Context'], row['Instructions'], source_path, os.path.join(save_path,str(i+1)))
            if success:
                success_count += 1
                result_queue.put(os.path.join(save_path,str(i+1)))
                log['Success Response'].append(res)
            else:
                log['Fail Response'].append(res)

        log['Success Count'] = success_count
        log['Total Count'] = config['repeat']
        # check_consistency(save_path)
        with open(os.path.join(save_path,'log.yaml'), 'w') as f:
            f.write(yaml.dump(log))
        
def check_consistency(dir_path: str):
    files = os.listdir(dir_path)
    count = len(files)
    consisstency = [1 for i in range(count)]
    for i in range(count):
        for j in range(i+1,count):
            files1 = dir_path + "/" + files[i]
            files2 = dir_path + "/" + files[j]
            report, sussess = compare_workbooks(files1, files2)
            if sussess:
                consisstency[i] += 1
                consisstency[j] += 1
            # report = compare_workbooks(ground_truth_file, rpa_processed_file, check_board)
            # print(json.dumps(report, indent=2))
    print(consisstency)

def checker():
    global result_queue, is_done
    while not is_done:
        if not result_queue.empty():
            ground_truth_file = result_queue.get()
            rpa_processed_file = ground_truth_file
            report = compare_workbooks(ground_truth_file, rpa_processed_file)
            print(json.dumps(report, indent=2))
        if is_done:
            break
        time.sleep(1)

lock = threading.Lock()

def process(path):
    pythoncom.CoInitialize()
    agent = Agent(config['Agent'])
    lock.acquire()
    time.sleep(0.01)
    agent._backend.OpenWorkbook(path)
    lock.release()
    for i in range(1, agent._backend.activeWS.UsedRange.Rows.Count+1):
        for j in range(1, agent._backend.activeWS.UsedRange.Columns.Count+1):
            lock.acquire()
            time.sleep(0.01)
            agent._backend.activeWS.Cells(i,j).Value = 'test'+'_'+str(i)+'_'+str(j)
            lock.release()
    pythoncom.CoUninitialize()

async def main():
    await worker()

if __name__ == '__main__':
    # t1 = threading.Thread(target=process, args=('../example_sheets_part1/Business/Invoices.xlsx',))
    # t2 = threading.Thread(target=process, args=('../example_sheets_part1/Business/PricingTable.xlsx',))
    # t1.start()
    # time.sleep(0.5)
    # t2.start()
    # t1.join()
    # t2.join()
    asyncio.run(main())