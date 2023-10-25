import os, openpyxl, tqdm, random
from collections import defaultdict

def main():
    output_file = f"./dataset.xlsx"

    out_wb = openpyxl.Workbook()
    out_ws = out_wb.worksheets[0]

    out_ws['A1'], out_ws['B1'], out_ws['C1'], out_ws['D1'], out_ws['E1'], out_ws['F1'],out_ws['G1'], out_ws['H1'], out_ws['I1'], out_ws['J1'] = \
    "Sheet Name", "Context", "Instructions", "Source", "Categories", "Atomic ations", "Seed task", "Validity", "Reason", "Chosen"

    input_file = f"./SU_adaptation_check/check.xlsx"
    
    wb = openpyxl.load_workbook(input_file)
    ws = wb.worksheets[0]

    data = [[cell.value for cell in row] for row in ws.iter_rows(min_row=2)]
    wb.close()

    for row in data:
        out_ws.append(row)

    sample_dict = defaultdict(list)

    # group row ids according to seed tasks
    for row_id in range(2, out_ws.max_row + 1):
        validity = out_ws[f'H{row_id}'].value
        if validity == 'Invalid': continue

        seed_task = out_ws[f'G{row_id}'].value

        sample_dict[seed_task].append(row_id)

    cnt = 0

    sample_per_seed_task = 6
    with tqdm.tqdm(total=sample_per_seed_task * 67) as pbar:
        for seed_task, row_ids in sample_dict.items():
            chosen_row_ids = random.sample(row_ids, min(len(row_ids), sample_per_seed_task))
            
            for row_id in chosen_row_ids:
                out_ws[f'J{row_id}'] = 1
                cnt += 1
                pbar.update(1)
        
    out_wb.save(output_file)
    out_wb.close()
    print('Saving to ', output_file)

main()

