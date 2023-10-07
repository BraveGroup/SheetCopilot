import yaml
import os

with open('config/config.yaml', 'r') as f:
    config = yaml.load(f, Loader=yaml.Loader)

save_path = config['path']['save_path']
task_dir_list = os.listdir(save_path)
success_count = 0
total_count = 0
for task_dir in task_dir_list:
    task_path = os.path.join(save_path, task_dir)
    if not os.path.isdir(task_path):
        continue
    if not os.path.exists(os.path.join(task_path, f'{task_dir}_log.yaml')):
        continue
    with open(os.path.join(task_path, f'{task_dir}_log.yaml'), 'r') as f:
        log = yaml.load(f, Loader=yaml.Loader)
    success_count += log['Success Count']
    total_count += log['Total Count']

print(f'Success Rate: {success_count}/{total_count} = {success_count/total_count:.2f}')