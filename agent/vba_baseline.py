from utils import ChatGPT
from Agent.xwAPI import xwBackend
import yaml
import pandas as pd
import os, shutil, time, re
import asyncio
import multiprocessing
import win32com.client as win32
import pythoncom
import argparse

parser = argparse.ArgumentParser(description='Process config.')
parser.add_argument('--config', '-c', type=str, help='path to config file')
args = parser.parse_args()

PROMPT = '''{context}
Sheet State: {sheetState}
Instructions: {instructions}
Please write a VBA script to implement the instructions above.
'''

with open(args.config, 'r') as f:
    config = yaml.load(f, Loader=yaml.Loader)

save_path = config['path']['save_path']
os.makedirs(save_path, exist_ok=True)

index_path = os.path.join(save_path, 'index.txt')
if not os.path.exists(index_path):
    with open(index_path, 'w') as f:
        f.write('')
with open(index_path, 'r') as file:
    index_list = file.readlines()
    index_list = [int(line.strip()) for line in index_list]

exe_index_path = os.path.join(save_path, 'exe_index.txt')
if not os.path.exists(exe_index_path):
    with open(exe_index_path, 'w') as f:
        f.write('')
with open(exe_index_path, 'r') as file:
    exe_index_list = file.readlines()
    exe_index_list = [int(line.strip()) for line in exe_index_list]

task_queue = asyncio.Queue()

task_df = pd.read_excel(config['path']['task_path'], header=0)
async def producer():
    for index, row in task_df.iterrows():
        if index+1 in index_list:
            continue
        source_path = os.path.join(config['path']['source_path'],row['Sheet Name']+'.xlsx')
        path = os.path.join(save_path,f"{index+1}_{row['Sheet Name']}")
        if not os.path.exists(path):
            os.makedirs(path)
            destination_path = os.path.join(path, f"{index+1}_{row['Sheet Name']}_source.xlsx")
            shutil.copy(source_path, destination_path)
        await task_queue.put((index+1, os.path.join(path, f"{index+1}_{row['Sheet Name']}_"), row['Context'], row['Instructions']))

async def worker():
    xwBot = xwBackend()
    chatbot = ChatGPT(config['Agent']['ChatGPT_1'])
    while True:
        index, path, context, instructions = await task_queue.get()
        source_path = path + 'source'+'.xlsx'
        log = {
            'Source Path': os.path.abspath(source_path),
            'Context': context,
            'Instructions': instructions,
            'VBA Code': False,
        }
        time.sleep(0.2)
        xwBot.OpenWorkbook(source_path)
        sheetState = xwBot.GetSheetsState()
        prompt = PROMPT.format(context=context, sheetState=sheetState, instructions=instructions)
        response = await chatbot(prompt)
        chatbot.context = []
        log['Prompt'] = prompt
        log['Response'] = (response)
        vba_code = re.findall(r'Sub .*?\(\).*?End Sub', response, re.DOTALL)
        if len(vba_code) == 0:
            print('No VBA code found for task', index)
            log['Error'] = 'No VBA code found'
        else:
            print('VBA code found for task', index)
            vba_code = vba_code[0]
            log['VBA Code'] = True
            with open(path+'vba_code.bas', 'w') as f:
                f.write(vba_code)
            with open(path+'log.yaml', 'w', encoding='utf-8') as f:
                f.write(yaml.dump(log, allow_unicode=True))
        with open(index_path, 'a') as f:
            f.write(str(index)+'\n')
        xwBot.activeWB.Close(SaveChanges=False)


def execute():
    for index, row in task_df.iterrows():
        if index+1 in exe_index_list:
            continue
        task_name = f"{index+1}_{row['Sheet Name']}"
        task_path = os.path.join(save_path, task_name)
        vba_code_path = os.path.join(task_path, f'{task_name}_vba_code.bas')
        if os.path.exists(vba_code_path):
            log_path = os.path.join(task_path, f'{task_name}_log.yaml')
            source_path = os.path.join(task_path, f'{task_name}_source.xlsx')
            run_macro(source_path, vba_code_path, log_path, task_path)
            with open(exe_index_path, 'a') as f:
                f.write(str(index+1)+'\n')

def run_macro(source_path, vba_code_path, log_path, save_path):
    excel = win32.Dispatch('Excel.Application')
    excel.DisplayAlerts = False
    excel.Visible = True
    time.sleep(0.2)
    print('Running macro for', source_path, '...')
    wb = excel.Workbooks.Open(os.path.abspath(source_path))
    with open(log_path, 'r', encoding='utf-8') as f:
        log = yaml.load(f, Loader=yaml.Loader)
    with open(vba_code_path, 'r') as f:
        vba_code = f.read()
    SubName = re.findall(r'Sub (.*?)\(\)', vba_code)[0]
    excelModule = wb.VBProject.VBComponents.Add(1)
    excelModule.CodeModule.AddFromString(vba_code)
    try:
        excel.Run(f'{SubName}')
        log['Success'] = True
    except Exception as e:
        print('Error:', e)
        log['Error'] = str(e)
        log['Success'] = False
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write(yaml.dump(log, allow_unicode=True))
    wb.SaveAs(os.path.abspath(os.path.join(save_path, f'vba.xlsx')))
    wb.Close(SaveChanges=False)


async def generate_vbacode():
    await producer()
    tasks = []
    for i in range(1):
        tasks.append(asyncio.create_task(worker()))

    await task_queue.join()

if __name__ == '__main__':
    start = time.time()
    asyncio.run(generate_vbacode())
    execute() # execute vba code from local files
    print('Time:', time.time()-start)
