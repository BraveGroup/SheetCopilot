# Description: This file is used to execute the excel agent from the log file
from Agent import Agent
import yaml, json
import time, os, shutil, datetime
import pandas as pd

def main():
    with open('config/GPT3.5_oursfull.yaml', 'r') as f:
        config = yaml.load(f, Loader=yaml.Loader)

    agent = Agent(config['Agent'])

    root_path = config['path']['save_path']
    dir_list = ['1_BoomerangSales'] # os.listdir(root_path)
    for each_dir in dir_list:
        dir_path = os.path.join(root_path, each_dir)
        if not os.path.isdir(dir_path):
            continue
        file_list = os.listdir(dir_path)
        for each_file in file_list:
            if each_file.endswith('.yaml'):
                log_path = os.path.join(dir_path, each_file)
                with open(log_path, 'r', encoding='utf-8') as f:
                    log = yaml.load(f, Loader=yaml.Loader)
                for each_response in log['Success Response']:
                    agent._backend.OpenWorkbook(log['Source Path'])
                    for each_refined_actions in each_response['refined response']:
                        success, error = agent.Process(each_refined_actions)
                        if not success:
                            each_response['error'] = error
                            break
                    agent._backend.activeWB.Close(SaveChanges=False)
                    print(error)
                with open(log_path, 'w', encoding='utf-8') as f:
                    f.write(yaml.dump(log, allow_unicode=True))
                print(f"Finished {log_path}")

if __name__ == '__main__':
    main()

