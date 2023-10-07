import os, yaml
from utils import compare_workbooks
import tqdm
import pandas as pd
import time
import numpy as np
from datetime import datetime

DEBUG = False
USE_NO_AND_SHEETNAME = False

def main():
    with open("./config/vba_config.yaml", 'r') as f:
        config = yaml.load(f, Loader=yaml.Loader)

    task_path = config['path']['task_path']
    gt_path = config['path']['gt_path']
    save_path = config['path']['save_path']
    eval_result_path = os.path.join(config['path']['save_path'], 'eval_result.yaml')

    if os.path.exists(eval_result_path):
        with open(eval_result_path, 'r') as f:
            eval_result = yaml.load(f, Loader=yaml.Loader)
    else:
        eval_result = {"check_result_each_repeat": {}}

    task_df = pd.read_excel(task_path, header=0)

    print("\033[0;36;40mEvaluate task result: {}\033[0m\n".format(save_path))
    
    for repeat_id in range(1, config["repeat"] + 1):
        t = time.time()
        if eval_result["check_result_each_repeat"].get(repeat_id, None) is None:
            eval_result["check_result_each_repeat"][repeat_id] = {
                "matched_gt_lst": [],
                "checked_list": [],
                "exec_success_list": [],
                "success_list": [],
                "gt_min_action_cnt_list": [],
                "check_result_list": [],
                "Code_length_list": [],
                "error_log": [],
                "eval_results": {}
            }

        check_result = eval_result["check_result_each_repeat"][repeat_id]

        remaining_task_cnt = len(task_df.iloc[:]) - len(check_result["checked_list"])

        with tqdm.tqdm(total=remaining_task_cnt, desc=f"Processing the remaining {remaining_task_cnt}/{len(task_df.iloc[:])} results of repeat {repeat_id}") as pbar:
            for index, row in task_df.iloc[:].iterrows():
                if index + 1 in check_result["checked_list"]: continue
                if DEBUG and index % 20 != 0: continue

                # Result file
                if USE_NO_AND_SHEETNAME:
                    task_name = f"{row['No.']}_{row['Sheet Name']}"
                else:
                    task_name = f"{index + 1}_{row['Sheet Name']}"
                
                task_path = os.path.join(save_path, task_name)
                if not os.path.exists(task_path):
                    continue
                res_path = os.path.join(task_path, f"{task_name}_{repeat_id}.xlsx")

                # Load the running log of the task
                log_file = os.path.join(task_path, "{}_log.yaml".format(os.path.basename(task_path)))

                with open(log_file, 'r', encoding="utf-8") as f:
                    log = yaml.load(f, yaml.Loader)

                if log["Success Count"] > 0:
                    if USE_NO_AND_SHEETNAME:
                        check_result["exec_success_list"].append(task_name)
                    else:
                        check_result["exec_success_list"].append(index+1)

                if os.path.exists(log_file) and 'conditional' not in row['Atomic actions'].lower():
                    # Compare the result with all reference solutions.
                    # All reference solutions for one sheet is placed under a folder with the same name.
                    # gt_folder_this_task = os.path.join(gt_path, row['Sheet Name'], f"{row['No.']}_{row['Sheet Name']}")

                    # Load GTs
                    gt_folder_this_task = os.path.join(gt_path, row['Sheet Name'], f"{row['No.']}_{row['Sheet Name']}")

                    for gt_file in [x for x in os.listdir(gt_folder_this_task) if x.endswith('.xlsx') and "$" not in x]:
                        gt = os.path.join(gt_folder_this_task, gt_file)
                        check_board = os.path.join(gt_folder_this_task, gt_file.replace(".xlsx", "_check.yaml"))

                        with open(check_board, 'r') as f:
                            check_board = yaml.load(f, Loader=yaml.Loader)

                        if not os.path.exists(gt) or not os.path.exists(res_path):
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
                        if check_res[1] and log["Success Count"] > 0:
                            if USE_NO_AND_SHEETNAME:
                                check_result["success_list"].append(task_name)
                            else:
                                check_result["success_list"].append(index+1)

                            check_result["Code_length_list"].append(log["Code Length"])
                            check_result["matched_gt_lst"].append(gt_file)
                            break
                    else:
                        print(f"Check error: {index+1}_{row['Sheet Name']}")
                        print(check_res[0])

                    with open(eval_result_path, 'w') as f:
                        yaml.dump(eval_result, f)
                
                check_result["checked_list"].append(task_name if USE_NO_AND_SHEETNAME else index+1)

                pbar.update(1)

        print("Evaluation for VBA Repeat {} has finished. Time elapse: {}s".format(repeat_id, time.time() - t))
        print("Error Log: {}\n".format('\n'.join(x for x in check_result["error_log"])))
        exec_success_cnt, success_cnt, total = len(check_result["exec_success_list"]), len(check_result["success_list"]), len(check_result["checked_list"])
        print("Total:", total)
        print("Excecution Success Cnt:", exec_success_cnt)
        print("Excecution Success Rate:", exec_success_cnt/total)
        print("Pass cnt:", success_cnt)
        print("Pass Rate:", success_cnt/total)

        print("Mean code length:", np.mean(check_result["Code_length_list"]).item())
        print("Median code length:", np.median(check_result["Code_length_list"]).item())
        print("90-percentile code length:", np.percentile(check_result["Code_length_list"], 90).item())

        print("Exec Success list:", check_result["exec_success_list"])
        print("Success list:", check_result["success_list"])

        # Save the metrics to the eval_result and save it
        with open(eval_result_path, 'w') as f:
            check_result["eval_results"]["Exec Success Rate"] = exec_success_cnt/total
            check_result["eval_results"]["Pass@1"] = success_cnt/total
            
            check_result["eval_results"]["Median code length"] = np.median(check_result["Code_length_list"]).item()
            check_result["eval_results"]["Mean code length"] = np.mean(check_result["Code_length_list"]).item()
            check_result["eval_results"]["90-percentile code length"] = np.percentile(check_result["Code_length_list"], 90).item()

            yaml.dump(eval_result, f)

    print("{} have been evaluated ... . Time: {}".format(save_path, datetime.now().strftime("%H:%M:%S")))
    
main()