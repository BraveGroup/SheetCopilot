from utils import compare_workbooks
import tqdm
import yaml
import pandas as pd
import os
import time
import numpy as np
from datetime import datetime
import argparse
from collections import defaultdict

USE_NO_AND_SHEETNAME = False

def evaluate(config):
    task_path = config['path']['task_path']
    gt_path = config['path']['gt_path']
    save_path = config['path']['save_path']
    eval_result_path = os.path.join(config['path']['save_path'], 'eval_result.yaml')

    print("Evaluate the results at ", save_path)
    if os.path.exists(eval_result_path):
        with open(eval_result_path, 'r') as f:
            eval_result = yaml.load(f, Loader=yaml.Loader)
    else:
        eval_result = {"check_result_each_repeat": {}}

    task_df = pd.read_excel(task_path, header=0)

    print("\033[0;36;40m========================================================\nEvaluate task result: {}\033[0m\n".format(save_path))
    
    for repeat_id in range(1, config["repeat"] + 1):
        t = time.time()
        if eval_result["check_result_each_repeat"].get(repeat_id, None) is None:
            eval_result["check_result_each_repeat"][repeat_id] = {
                "matched_gt_lst": [],
                "checked_list": [],
                "exec_success_list": [],
                "success_list": [],
                "checked_list_by_cate": defaultdict(list),
                "exec_success_list_by_cate": defaultdict(list),
                "success_list_by_cate": defaultdict(list),
                "gt_min_action_cnt_list": [],
                "action_cnt_list": [],
                "error_log": [],
                "eval_results": {}
            }

        check_result = eval_result["check_result_each_repeat"][repeat_id]

        num_tasks = len([x for x in os.listdir(save_path) if os.path.isdir(os.path.join(save_path, x))])

        remaining_task_cnt = num_tasks - len(check_result["checked_list"])
        assert remaining_task_cnt > 0, "No tasks to be evaluated"

        with tqdm.tqdm(total=remaining_task_cnt, desc=f"Processing the remaining {remaining_task_cnt}/{num_tasks} results of repeat {repeat_id}") as pbar:
            for index, row in task_df.iloc[:].iterrows():
                if index + 1 in check_result["checked_list"]: continue

                # Result file
                if USE_NO_AND_SHEETNAME:
                    task_name = f"{row['No.']}_{row['Sheet Name']}"
                else:
                    task_name = f"{index + 1}_{row['Sheet Name']}"

                task_path = os.path.join(save_path, task_name)
                if not os.path.exists(task_path):
                    continue

                res_path = os.path.join(task_path, f"{task_name}_{repeat_id}.xlsx") #Claude

                # Load the running log of the task
                log_file = os.path.join(task_path, "{}_log.yaml".format(os.path.basename(task_path)))
                if not os.path.exists(log_file): continue
                
                with open(log_file, 'r', encoding='utf-8') as f:
                    log = yaml.load(f, yaml.Loader)

                # Check if the result xlsx file exists
                res_file_exists = os.path.exists(res_path)

                cates = row['Categories'].split(', ')
                if log["Success Count"] > 0 and res_file_exists:
                    check_result["exec_success_list"].append(task_name)
                    for cate in cates:
                        check_result["exec_success_list_by_cate"][cate].append(task_name)

                if os.path.exists(log_file) and res_file_exists:
                    # Compare the result with all reference solutions.
                    # All reference solutions for one sheet is placed under a folder with the same name.

                    # Load GTs
                    gt_folder_this_task = os.path.join(gt_path, row['Sheet Name'], f"{row['No.']}_{row['Sheet Name']}")

                    for gt_file in [x for x in os.listdir(gt_folder_this_task) if x.endswith('.xlsx') and "$" not in x]:
                        gt = os.path.join(gt_folder_this_task, gt_file)
                        check_board = os.path.join(gt_folder_this_task, gt_file.replace(".xlsx", "_check.yaml"))

                        with open(check_board, 'r') as f:
                            check_board = yaml.load(f, Loader=yaml.Loader)

                        if not os.path.exists(gt):
                            check_result["error_log"].append("{} not exists".format(os.path.basename(res_path))) 
                            continue
                        
                        """
                        Comparing.......
                        Comparing..............
                        Comparing.....................
                        """
                        check_res = compare_workbooks(gt, res_path, check_board["check_board"])
                        """
                        Comparing.....................
                        Comparing..............
                        Comparing.......
                        """

                        # If checking is successful
                        if check_res[1] and len(log["Success Response"]) > 0:
                            check_result["success_list"].append(task_name)
                            for cate in cates:
                                check_result["success_list_by_cate"][cate].append(task_name)

                            # Count the number of actions in the generated plan, regardless of execution success or failure
                            num_acts = 0
                            plan = log["Success Response"][repeat_id - 1]["refined response"]
                            for steps in plan:
                                num_acts += len(steps)
                            check_result["action_cnt_list"].append(num_acts)

                            # Count the minimum number of actions among Gts
                            gt_actions = [x for x in row['Atomic actions'].split(',') if "function" not in x]
                            check_result["gt_min_action_cnt_list"].append(len(gt_actions))
                            check_result["matched_gt_lst"].append(gt_file)
                            
                            # Matched GT found. Stop checking for the task
                            break
                    
                    with open(eval_result_path, 'w') as f:
                        yaml.dump(eval_result, f)
                
                check_result["checked_list"].append(task_name)
                for cate in cates:
                    check_result["checked_list_by_cate"][cate].append(task_name)

                pbar.update(1)

        print("\033[0;33;40mEvaluation for Repeat {} has finished. Time elapse: {:.2f}s\033[0m".format(repeat_id, time.time() - t))
        print("Error Log: {}\n".format('\n'.join(x for x in check_result["error_log"])))
        exec_success_cnt, success_cnt, total = len(check_result["exec_success_list"]), len(check_result["success_list"]), len(check_result["checked_list"])
        action_cnt_list, gt_min_action_cnt_list = np.array(check_result["action_cnt_list"]), np.array(check_result["gt_min_action_cnt_list"])

        check_result["eval_results"]["Total"] = total
        check_result["eval_results"]["Exec@1"] = exec_success_cnt / total
        check_result["eval_results"]["Pass@1"] = success_cnt / total
        
        for k, v in check_result["exec_success_list_by_cate"].items():
            cate_total = len(check_result['checked_list_by_cate'][k])
            check_result["eval_results"][f"{k} Exec & Pass"] = "{:d}/{:d} & {:d}/{:d}".format(len(v), cate_total, len(check_result["success_list_by_cate"][k]), cate_total)
        
        # Task status
        check_result["action_cnt_list"] = ', '.join(str(x) for x in check_result["action_cnt_list"])
        check_result["gt_min_action_cnt_list"] = ', '.join(str(x) for x in check_result["gt_min_action_cnt_list"])
        check_result["checked_list"] = ', '.join(str(x) for x in check_result["checked_list"])
        check_result["success_list"] = ', '.join(str(x) for x in check_result["success_list"])
        check_result["exec_success_list"] = ', '.join(str(x) for x in check_result["exec_success_list"])
        check_result['checked_list_by_cate'] = {k: ', '.join(str(x) for x in check_result['checked_list_by_cate'][k]) for k in check_result['checked_list_by_cate'].keys()}
        check_result['success_list_by_cate'] = {k: ', '.join(str(x) for x in check_result['success_list_by_cate'][k]) for k in check_result['success_list_by_cate'].keys()}
        check_result['exec_success_list_by_cate'] = {k: ', '.join(str(x) for x in check_result['exec_success_list_by_cate'][k]) for k in check_result['exec_success_list_by_cate'].keys()}
        
        # Action statistics
        check_result["eval_results"]["A_mean"] = np.mean(action_cnt_list).item()
        check_result["eval_results"]["A50_norm"] = np.median(action_cnt_list / gt_min_action_cnt_list).item()
        check_result["eval_results"]["A90_norm"] = np.percentile(action_cnt_list / gt_min_action_cnt_list, 90).item()
        
        for k, v in check_result["eval_results"].items():
            print("{}: {}".format(k, v))
        
        print("========================================================\n")

        # Save the metrics to the eval_result and save it
        with open(eval_result_path, 'w') as f:
            yaml.dump(eval_result, f)

    print("{} have been evaluated on {}... . Time: {}".format(save_path, gt_path, datetime.now().strftime("%H:%M:%S")))

parser = argparse.ArgumentParser(description='Process config.')
parser.add_argument('--config', '-c', default="./config/config.yaml", type=str, help='path to config file')

args = parser.parse_args()

if __name__ == '__main__':
    with open(args.config, 'r') as f:
        config = yaml.load(f, Loader=yaml.Loader)
    
    evaluate(config)
    print("Evaluate {}".format(config["path"]["save_path"]))
