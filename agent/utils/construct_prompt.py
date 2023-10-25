import re

def get_prompt(context, instruction, template):
    return template.format(context, instruction)

def get_sheetcopilot_api_doc(raw_doc):
    api_list = []
    api_usage = []
    api_detail_doc = {}
    for k, v in raw_doc.items():
        if v.get('display') is not None:
            api_usage.append(f"{v['display']} # Args: {v['args']} Usage: {v['usage']}")
            api_list.append(v['display'])
            new_example = v['example'].replace(k+'(', v['display']+'(') if v['example'] is not None else None
            api_detail_doc[v['display']] = f'{v["display"]}{v["args"]}\nArgs explanation:\n{v["args explanation"]}\nUsage example:\n{new_example}'
            # api_detail_doc[v['display']] = f'{v["display"]}{v["args"]}\nArgs explanation:\n{v["args explanation"]}\n'
        else:
            api_usage.append(f"{k} # Args: {v['args']} Usage: {v['usage']}")
            api_list.append(k)
            api_detail_doc[k] = f'{k}{v["args"]}\nArgs explanation:\n{v["args explanation"]}\nUsage example:\n{v["example"]}'

    api_usage = '\n'.join(api_usage)
    
    return api_list, api_usage, api_detail_doc

def get_toolllama_api_doc(raw_doc):
    api_list = []
    api_usage = [] # used as the doc. in the sys. prompt
    api_detail_doc = {}
    for k, v in raw_doc.items():
        examples = v["example"]
        
        if examples is None:
            formatted_examples = "" #
        else:
            formatted_examples = []
            for line in examples.split('\n'):
                if line.strip() == "":
                    continue
                elif "#" in line or "..." in line: 
                    formatted_examples.append(line + "\n")
                else:                
                    api_call_and_comment = line.split('#')
                    if len(api_call_and_comment) == 1:
                        comment = ""
                        api_call = api_call_and_comment[0]
                    else:
                        api_call, comment = api_call_and_comment
                    
                    arg_dict = {}
                    
                    Lbr = api_call.find('(')
                    Rbr = api_call.rfind(')')
                    api_name, args = api_call[:Lbr], api_call[Lbr+1:Rbr].split(', ')
                    
                    arg_names_in_doc = re.findall(r'(\w+)(?=\s*:\s*\w+)', raw_doc[api_name]['args'])
                    
                    if len(arg_names_in_doc) > 0: # Some APIs have no arguments, e.g. DeleteFilter
                        # Get the argument list of the current API
                        for arg_position, arg in enumerate(args):
                            if "=" in arg:
                                equal_id = arg.find("=") # Values may also contain "=" so we find the first one
                                arg_name, arg_value = arg[:equal_id], arg[equal_id+1:]
                                arg_dict[arg_name] = arg_value
                            else:
                                arg_dict[arg_names_in_doc[arg_position]] = arg
                    
                    reformatted_api_call = "Thought: {}\nAction: {}\nAction Input:".format(comment, api_name) + " {\n" + ',\n'.join(['"{}": "{}"'.format(k.strip(), v.strip('\'" ')) for k, v in arg_dict.items()]) + "\n}\n"
                    
                    formatted_examples.append(reformatted_api_call)
            
            formatted_examples = "".join(formatted_examples)
            
        if v.get('display') is not None:
            api_usage.append(f"{v['display']} # Args: {v['args']} Usage: {v['usage']}")
            api_list.append(v['display'])
            new_example = v['example'].replace(k+'(', v['display']+'(') if v['example'] is not None else None
            api_detail_doc[v['display']] = f'{v["display"]}{v["args"]}\nArgs explanation:\n{v["args explanation"]}\nUsage example:\n{new_example}'
            # api_detail_doc[v['display']] = f'{v["display"]}{v["args"]}\nArgs explanation:\n{v["args explanation"]}\n'
        else:
            api_usage.append(f"{k} # Args: {v['args']} Usage: {v['usage']}")
            api_list.append(k)
            api_detail_doc[k] = f'{k}{v["args"]}\nArgs explanation:\n{v["args explanation"]}\nUsage example:\n{formatted_examples}'

    api_usage = '\n'.join(api_usage)
    
    return api_list, api_usage, api_detail_doc

def get_api_doc(prompt_format, raw_doc):
    if "toolllama" in prompt_format.lower():
        return get_toolllama_api_doc(raw_doc)
    else:
        return get_sheetcopilot_api_doc(raw_doc)
