
import tiktoken, openai, requests
encoding = tiktoken.encoding_for_model('gpt-3.5-turbo')

def num_tokens_from_string(string: str) -> int:
    """Returns the number of tokens in a text string."""
    num_tokens = len(encoding.encode(string))
    return num_tokens

def generate_column_details(ws):
    column_description = []
    for col_id in range(ws.max_column):
        col_name = chr(ord('A') + col_id)

        if ws[f'{col_name}2'].data_type in ['s', 'b', 'd']: # list value options with type string, bool, date/time
            cell_options = set()
            for row_id in range(2, ws.max_row + 1):
                v = ws[f'{col_name}{row_id}'].value
                if v is not None:
                    cell_options.add(str(v))
            
            if len(cell_options) > 0:
                column_description.append('The cells in the "{}" column can be {}.'.format(
                    ws[f'{col_name}1'].value,
                    ', '.join('"{}"'.format(cell_options.pop()) for _ in range(min(20, len(cell_options))))
                ))
        elif ws[f'{col_name}2'].data_type in ['n']:
            cell_options = []
            for row_id in range(2, ws.max_row + 1):
                v = ws[f'{col_name}{row_id}'].value
                if v is not None:
                    cell_options.append(float(v))

            if len(cell_options) > 0:
                column_description.append('The cells in the "{}" column range from {:.2f} to {:.2f}.'.format(
                    ws[f'{col_name}1'].value,
                    min(cell_options),
                    max(cell_options)
                ))
            

    return ' '.join(column_description)

def generate_state(wb, use_col_detail=True):
    state = []
    for ws_name in wb.sheetnames:
        ws = wb.get_sheet_by_name(ws_name)
        column_id = [chr(ord('A') + i) for i in range(ws.max_column)]

        headers = []
        for col_id in column_id:
            header = ws[f'{col_id}1'].value
            if header is not None:
                headers.append(header)

        header_des = ', '.join('{}: "{}"'.format(column_id[i], headers[i]) for i in range(len(headers)))

        column_info = 'Sheet "{}" has {} columns'.format(ws_name, len(headers))
        header_info = ' (Headers are {})'.format(header_des) if len(headers) > 0 else ''
        row_info = ' and {} rows (including the header row).'.format(ws.max_row) if len(headers) > 0 else ' and 0 rows.'

        ws_info = column_info + header_info + row_info
        if use_col_detail:
            col_detail = generate_column_details(ws)
            ws_info += ' ' + col_detail
        state.append(ws_info)
    
    return ' '.join(state)


def ask(input_text, gpt_mode, bot=None):
    if gpt_mode == 'wrapper':
        bot.new_conversation()
        response_text = bot.ask(input_text)
    elif gpt_mode == 'api':
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                    {"role": "system", "content": "You are an Excel expert."},
                    {"role": "user", "content": input_text}
                ],
            n=1,
            temperature = 0.0
            )
        response_text = response['choices'][0]['message']['content']
    elif gpt_mode == 'proxy':
        keys = [
            {'url': 'https://api.openai-sb.com/v1/chat/completions',
            'Authorization':'Bearer sb-139b0e3d71f238a0fbacc73adba5f09f0151c4848a93c80b'},
            {'url': 'https://o-api-mirror01.gistmate.hash070.com/v1/chat/completions',
            'Authorization':'Bearer sk-XWUkrLrPeUjqQ1k9ipAaHig4DgDyd2jhh09eVIVRwvhRTi5g'},
            {'url': 'https://api.openai.com/v1/chat/completions',
             'Authorization': 'Bearer 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'}
        ]

        sucess = False
        while not sucess:
            for prop in keys:
                try:
                    url = prop['url']
                    headers = {'Content-Type':'application/json', 'Authorization':prop['Authorization']}
                    data = {
                        
                        'model':'gpt-3.5-turbo',
                        "messages": [{"role": "system", "content": "You are an Excel expert."},
                                {"role": "user", "content": input_text}
                                ]
                    }
                    response = requests.post(url, headers=headers, json=data, timeout=120).json()

                    sucess = True
                    break
                except Exception as e:
                    print("Time out. Change proxy...")
        
        response_text = response['choices'][0]['message']['content']
    
    return response_text
