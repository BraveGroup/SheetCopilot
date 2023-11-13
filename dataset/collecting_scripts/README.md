**SheetCopilot**: Bringing Software Productivity to the Next Level through Large Language Models
========

This folder contains the scripts used to generate our dataset. These scripts are introduced below:

# Step 1: Adapt SuperUser seed tasks for example sheets

```adapt_superuser_tasks.py``` is used to adapt the 67 seed tasks to generate new tasks applied to [the task sheets](../task_sheets/). Specifically, ChatGPT is prompted to recognize the atomic actions used in a seed task and then generate a new task applied to a given spreadsheet using these recognized actions. 

## Dataset Stats
Seed Task: 67

After adaptation, 496 Business, 767 Finance, 229 Physics, 177 Economics (1669 new instructions)

## Used prompt

As an Excel expert, you have been assigned the responsibility of adapting a set of task instructions for specific Excel workbooks. These instructions will be utilized to evaluate the Excel manipulation capabilities of large language models.

Requirements:

1. First, identify individual atomic actions used in the original instructions, then develop new instructions incorporating these actions.

2. Use the detailed descriptions of the provided workbooks to modify the original instructions so that they become compatible with the workbooks. Specifically, you must change the manipulated objects (ranges, sheets, rows, and columns) in the original instructions. You must also change the way you use atomic actions. For instance, if the original instruction sets the data type as accounting, you can change it to other types in the adaptation.

3. Use standard range references, such as 'Sheet2!A1:C9', 'A2:E16', 'A:H', column C, or row 16.

4. Use different phrasing (e.g., various sentence structures and noun/verb forms) to create diverse instructions.

5. Apply domain-specific knowledge to diversify the generated instructions. For instance, use financial knowledge to calculate various metrics, demonstrate trends in product sales from different perspectives, visualize data using various types of charts, and so on.

6. (Important!) The generated instructions must describe realistic tasks and involve the normal use of atomic actions according to the workbook descriptions.

7. In every new instruction, new tables and Pivot tables must be created in a new worksheet and start from A1. Only one new sheet is allowed to be created. Besides, the headers of new columns/rows and the names of new sheets must be set.

8. The instructions after adaptation should look like what a non-expert Excel user would say and should not mention any specific funtions or operations built in Excel.

Here are the atomic actions you can identify within the six categories:
...

Restrictions for the atomic action parameters:

Chart type can only be (stacked/clustered) column chart, (stacked/clustered) bar chart, (3D) pie chart, line chart (with smooth lines), (stacked) area chart, or scatter chart.
Cell value type can only be date, text, time, currency, percentage, number, or general.

I will give you an example first:

Given an Excel workbook:

The sheet 'Sheet1' records the employee working hours of a coffee shop. It has 8 columns (Headers are: A: "location", B: "name", C: "date", D: "hours", E: "ot hours", F: "base pay", G: "ot pay", H: "total pay") and 11 rows (including the header row). The cells in the "location" column can be "capitol hill", "queen anne". The cells in the "name" column can be "Aaron", "Bob", "Blanca", "Frank".

The original instructions to be adapted:

1. I'd like to know if there's a formula (or combination of formulas) to sum Column B ONLY if it equals Column A?

2. Column A contains multiple due dates and cell B1 is 10/31/2022. Append a special character to the cells in column A depending on whether the due date is less or more than the date in B1. If neither applies, leave the cell unchanged.

3. Create a Pivot Table to separate dates by each month similar to the quarterly function. This will show all dates in the last year under the column label. Choose the monthly option under data filters in the columns label under data filters.

4. Freeze A1:B10 so that no matter how I scroll vertically or horizontally this range is always frozen.

5. I have five groups of data each occupying two columns. I'd like to have them plotted all on one bar chart but with the series separated (i.e., not clustered) so that several column charts stick together sharing the same y-axis.


Adaptations compatible with the given workbook (Show the categories involved in the generated instruction and list the atomic actions following the category label):
1. Instruction: In a new column with header "Total pay each location", use a formula to sum the "total pay" (Column H) for each employee ONLY if their "location" (Column A) is "capitol hill". - Categories (atomic actions): A (Update cell value); F (Math functions)

2. Instruction: Create a formula in a new column (Column I) with header "Marked Due Dates" to check if the dates in column C are less or more than 10/31/2022. If the date is less, append a '-' to the cell in the same row in column C; if the date is more, append a '+'; otherwise, leave the cell unchanged. - Categories (atomic actions): A (Update cell value, Autofill); F (Logical functions, Text functions)

3. Instruction: Create a Pivot Table in a new sheet named "Sheet2" based on the data in 'Sheet1' and then summarize sum of hours for each location in this Pivot Table. - Categories (atomic actions): A (Create sheet); E (Create Pivot Table)

4. Instruction: Freeze the range A1:H1 so that the headers remain visible when scrolling vertically or horizontally. - Categories (atomic actions): C (Freeze panes)

5. Instruction: Create a Pivot Table in a new sheet named "Sheet2" to sum the hours for all dates and then plot a line chart to show the trend of hours changing with dates. - Categories (atomic actions): A (Create sheet); E (Create Pivot Table); D (Create chart)

# Check instructions with four criteria

As the generated tasks may be invalid due to vaguenes and incompleteness, we use ```check_task_feasibility.py``` to prompt ChatGPT to determine the validity of all these tasks according to 4 criteria (Realism, Relevance, Clarity, and Completeness). Tasks not fulfilling all criteria will be marked as 'invalid'.

## Check stats
Business: 466 valid and 30 invalid

Finance: 687 valid and 80 invalid

Economics: 166 valid and 11 invalid

Finance: 196 valid and 22 invalid

Total: 1515 valid and 143 invalid

## Used prompt
I will give you a batch of Excel task instructions that are utilized to evaluate the spreadsheet manipulation capabilities of large language models. Please check all instructions according to the criteria and the descriptions of the given workbooks.

Requirements:

1. If an instruction fulfills all of the following four criteria strictly at the same time, it is valid.

A. Realism: Verify that the instructions are representative of real-world Excel problems encountered by expert users. Avoid including contrived or overly complex tasks that would rarely be encountered in practice.

B. Relevance: Verify that the instructions are relevant to the context of the given workbooks. Reject instructions that do not relate to the content of the workbooks or are not applicable in the context of Excel.

C. Clarity: Verify that the instructions are clear and easy to understand. They should be free from grammatical errors and ambiguities.

D. Completeness: Verify that the instructions can be completed using the provided workbook data and Excel features. If additional data or functionality is needed to accomplish a task, reject the instruction.

2. Use your domain-specific knowledge to reason and make decisions. For example, it is unlikely to calculate monthly or yearly averages of some physical values because this task is impractical.

I will give you an example first:

Given an Excel workbook:

My workbook records many invoices made on different dates. Sheet "Sheet1" has 7 columns (Headers are A: "No.", B: "Date", C: "Salesman", D: "Product", E: "Price", F: "Units", G: "Sales") and 19 rows (including the header row). The cells in the "No." column range from 10500.00 to 10505.00. The cells in the "Date" column can be "2011-05-28 00:00:00", "2011-05-25 00:00:00", "2011-05-27 00:00:00". The cells in the "Salesman" column can be "Joe", "Chin", "Moe". The cells in the "Product" column can be "Quad", "Majestic", "Bellen", "Carlota", "Alpine". The cells in the "Price" column range from 22.00 to 32.00. The cells in the "Units" column range from 4.00 to 25.00. The cells in the "Sales" column range from 128.00 to 750.00.

Instructions to check:

1. Find the sales value corresponding to each No. in a new column called "Invoice Lookup".

2. Compare the names in the Salesman column to find the closest match for each name. Put the results in a new column named "Salesman Matched".

3. Create a sheet named "Sheet2" to summarize the total sales for each sales representative in Sheet1.

4. Prepend leading zeros to each number in the "No." column so that they have a fixed length of 6 digits. Append the new results to the corresponding product names in the "Product" column, and put the results in a new column named "Padded No.".

5. Find the corresponding date for each No. in a new column named "Lookup Date" in Sheet1.

6. Find and display the row values in Sheet1 based on the "No." column. Put the results in a new sheet named "Sheet2".

7. Round the values in the "Sales" column to two decimal places in a new column named "Rounded Sales", and display the results with trailing zeros.

8. Merge cells A1 through C1 with cells A2 through D2.

9. Add hyperlinks to each cell in the "No." column that link to its corresponding file.

Check results (Give brief reasons in the comments and tell me if the instruction is valid):

1. Realism: Yes, Relevance: Yes, Clarity: Yes, Completeness: Yes. Comment: The instruction fulfills the 4 criteria, so it is valid. 

2. Realism: No, Relevance: No, Clarity: No, Completeness: No. Comment: This instruction is unrealistic and unclear and does not seem to be relevant to the context of the given workbook, so it is invalid.

3. Realism: Yes, Relevance: Yes, Clarity: Yes, Completeness: Yes. Comment: The instruction fulfills the 4 criteria, so it is valid. 

4. Realism: No, Relevance: Yes, Clarity: Yes, Completeness: Yes. Comment: The instruction appends numbers to product names, which is not a realistic requirement, so it is invalid.

5. Realism: Yes, Relevance: Yes, Clarity: Yes, Completeness: Yes. Comment: The instruction fulfills the 4 criteria, so it is valid.

6. Realism: Yes, Relevance: Yes, Clarity: No, Completeness: No. Comment: The instruction does not specify what values to display, so it is invalid.

7. Realism: Yes, Relevance: Yes, Clarity: Yes, Completeness: Yes. Comment: The instruction fulfills the 4 criteria, so it is valid.

8. Realism: No, Relevance: No, Clarity: Yes, Completeness: Yes. Comment: The instruction merges cells, which destroys the original data and is meaningless in the context of the workbook, so it is invalid.

9. Realism: Yes, Relevance: Yes, Clarity: Yes, Completeness: No. Comment: The instruction does not refer to specific corresponding files, which is incomplete, so it is invalid.

Now it's your turn.

Given an Excel workbook:

{Workbook descriptions}

Instructions to check:

{Instructions}

Check results (Give brief reasons in the comments and tell me if the instruction is valid):

# Select a representative subset (our 221 tasks)

```pick_samples_forall_seed_tasks.py``` is used to select valid representative tasks from the 1669 raw tasks above. After selection, 221 tasks remain.


# Rephrase instructions so that they take on a non-expert tone

```naturalize_SU_adaptations.py``` is used to rephrase each task to make them sound like what a non-expert user says.

## Rephrase stats

Token length before rephrasing (SU adaptations):

| Field     | Min | Max | Avg  |
|-----------|-----|-----|------|
| Business  | 17  | 118 | 46.9 |
| Finance   | 13  | 157 | 47.1 |
| Economics | 12  | 112 | 48.5 |
| Physics   | 16  | 111 | 46.9 |
| Total     | 12  | 157 | 47.1 |

Token length after rephrasing:

| Field     | Min | Max | Avg  |
|-----------|-----|-----|------|
| Business  | 7   | 152 | 32.7 |
| Finance   | 3   | 137 | 33.3 |
| Economics | 8   | 91  | 36.0 |
| Physics   | 8   | 105 | 36.4 |
| Total     | 3   | 152 | 33.8 |

## Used prompt

You have been tasked with paraphrasing a set of instructions for Excel tasks.

Requirements:
1. Paraphrase each given instruction so that it is more like what a non-expert Excel user would say while retaining the original intention and order of actions in the instruction.

2. Do not mention any Excel built-in functions in the paraphrases. If the original instruction mentions the names of specific Excel built-in functions, you should replace these functions with spoken language. For example, "Use CONCATENATE to combine cells" mentions "CONCATENATE" and should be rephrased as "concatenate cells". "Use conditional formatting to change the background color of cells" mentions "conditional formatting" and should be rephrased as "Set cell background color if ...". "Write a formula to copy data" should be rephrased as "Copy data".

3. Do not refer to existing columns by their indices (e.g., column A is not desired); instead, mention columns by their headers. For instance, 'In column A' should be rephrased as 'In the Year column'.

4. Do not add column references in brackets. For instance, "subtracting Total Expenses (Column B) from Revenue (Column A)" should be rephrased as "Subtract Total Expenses from Revenue".

5. When inserting new columns, the column indices must be given to avoid ambiguity. Besides, each new column must be given a header, so you need to keep the column headers mentioned in the original instructions.

6. Use domain-specific knowledge to diversify the expression of the generated instructions and do not use the same expression repeatedly.

I will give you some examples first:

Original instructions:
1. Convert the "Year" column values (Column A) from the format of yyyy.00 to yyyy.
2. In a new column with the header "Combined data", use the CONCATENATE function to combine the data from columns A through Z in each row, including the headers, and then autofill the formula to the last row of data.
3. Use advanced filters in "Sheet1" to display rows with sales data for a specific month (assuming weeks starting from the 1st of each month). For example, filter to show rows with weeks "Week 1" to "Week 4" for the first month.
4. Create a line chart on a new sheet named "Sheet2" with weeks (Column A) on the x-axis and sales (Column B) as the y-axis.
5. Apply conditional formatting to Column H (Net Profit) based on Columns F (Operating Profit) and G (Tax Expense). If the cell value in Column G is not 0 and Column F is greater than 0, display the value in Column H in red.
6. In a new sheet named "Sheet3", use the VLOOKUP function to match each row in "Year" (Column A) in "Sheet1" with the corresponding value in Column B ("Net Sales") from "Sheet1".
7. In a new column (Column G) with the header "Tax Calculation", use a formula to calculate the tax by multiplying the corresponding "Subtotal" (Column D) by 0.1.

Think about the flaws in the original instructions before paraphrasing:

1.
Think: The original instruction refers to columns by indices in brackets, which is undesired.
Paraphrase: Convert the "Year" column format from yyyy.00 to yyyy.

2.
Think: The original instruction mentions Excel built-in functions CONCATENATE, which is undesired.
Paraphrase: Concatenate the data from columns A through Z for all rows and write the results in a new column named "Combined data".

3.
Think: The original instruction mentions an Excel built-in operation (i.e., advanced filters), which is undesired.
Paraphrase: Display only the rows with weeks from Week 1 to Week 4.

4.
Think: The original instruction refers to columns by indices in brackets, which is undesired.
Paraphrase: Create a line chart on a new sheet with the Week column on the x-axis and the Sales column as the y-axis.

5.
Think: The original instruction mentions Excel built-in functions (i.e., conditional formatting) and refers to columns by indices, which is undesired.
Paraphrase: If the cell value in Tax Expense Column is not 0 and that in Operating Profit Column > 0, display the cell text in the Net Profit column in red.

6.
Think: The original instruction mentions Excel built-in functions (i.e., VLOOKUP) and refers to columns by indices in brackets, which is undesired.
Paraphrase: Match cells in the Year column and return the corresponding values in the Net Sales Column. Put the results in a new sheet.

7.
Think: The original instruction refers to columns by indices in brackets, which is undesired.
Paraphrase: Calculate the tax by multiplying the "Subtotal" column by 0.1 in column G named "Tax Calculation".

Now it's your turn. Please follow the requirements to paraphrase the original instructions according to the given workbook descriptions.