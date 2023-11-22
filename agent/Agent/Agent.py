from .xwAPI import xwBackend
from typing import Tuple, List, Optional
import time
import requests
import pandas
import re, os
from colorama import Fore, Style
from utils import ChatGPT, StateMachine
import yaml, json
import asyncio
import copy

import os

from utils.construct_prompt import get_api_doc

class Agent:
    """
    The concrete Agent to execute Actions.
    """

    def __init__(self, config) -> None:
        """
        Constructor of the Agent class. Initializes the _backend attribute.

        Parameters:
            config (dict): the configuration dictionary.

        Returns:
            None.
        """
        self.agent_config = config['Agent']
        self.model_name = self.agent_config['ChatGPT_1']['model_name']
        self.prompt_format = self.agent_config['ChatGPT_1']['prompt_format']
        self.use_doc_in_syspromt = self.agent_config.get('use_doc_in_syspromt', True)
        self.use_ext_doc = self.agent_config.get('use_ext_doc', True)
        self.use_same_LLM = self.agent_config.get('use_same_LLM', True)
        self.add_example_data2feedback = self.agent_config.get('add_example_data2feedback', False)
        
        self.interaction_mode = config.get('interaction_mode', False)
        self.max_cycle_times = self.agent_config['max_cycle_times']
        self.max_error_count = self.agent_config['max_error_count']
        
        # Whether to use the ground truth API docs to restrict the action space
        self.use_oracle_API_doc = self.agent_config.get('use_oracle_API_doc', False)
        if self.use_oracle_API_doc:
            from utils import action2API
            dataset_df = pandas.read_excel(config['path']['task_path'])
            
            self.oracle_API_dict = {}
            for task_id, (sheet_name, acts) in enumerate(zip(dataset_df['Sheet Name'], dataset_df['Atomic actions']), start=1):
                oracle_acion_names = [x[:x.find('(')-1] if '(' in x else x for x in acts.split(', ') if "function" not in x]
                self.oracle_API_dict[f'{task_id}_{sheet_name}'] = [action2API[act] for act in oracle_acion_names]
                
        with open(self.agent_config['prompt_path'], 'r') as f:
            self.prompt = yaml.load(f, Loader=yaml.Loader)
        
        with open(self.agent_config['api_doc_path']) as f:
            self.api_doc = yaml.load(f, Loader=yaml.FullLoader)

        # Use fine-grained APIs
        self.use_finegrained_APIs = self.agent_config.get('use_finegrained_API', False)
        if self.use_finegrained_APIs:
            # Remove the SetFormat from the original API doc
            self.api_doc.pop('SetFormat')

            # Load the fine-grained formatting APIs
            with open(self.agent_config['api_doc_path'].replace('API_document.yaml', 'API_document_finegrained.yaml')) as f:
                self.api_doc.update(yaml.load(f, Loader=yaml.FullLoader))

        self.api_list, self.api_usage, self.api_detail_doc = get_api_doc(self.prompt_format, self.api_doc)
        self.statemachine = self.InitStateMachines()

        if self.agent_config['API_backend'] == 'xw':
            self._backend = xwBackend(self.agent_config['APP_backend'], self.api_doc)
        else:
            raise NotImplementedError('Backend {} is not implemented.'.format(self.agent_config['backend']))
        
        print("Initializing SheetCopilot...")
        print(f"-> Use the external doc: {self.use_ext_doc}\n> Use fine-grained APIs: {self.use_finegrained_APIs}\n-> Use oracle API doc: {self.use_oracle_API_doc}\n-> Use the same LLM for planning and parsing: {self.use_same_LLM}\n-> Add example data to the sheet state: {self.add_example_data2feedback}")
        
        self.step = 1
        self.error_count = 0

    def InitStateMachines(self) -> StateMachine:

        log = {}
        async def state1(chatbot: ChatGPT, prompt, context_index, new_step=False):
            if new_step:
                print(Fore.YELLOW + f"Step {self.step} ... ")
                self.step += 1
            
            print(Fore.CYAN + "Processing state1 - Coarse-grained planning"); print(Style.RESET_ALL)
            
            nonlocal log, cycles_times
            # Preventing dead loops
            cycles_times += 1
            if cycles_times > self.max_cycle_times:
                return 'fail', (chatbot, f'Too many cycles (> {self.max_cycle_times})')
            
            try:
                response = await chatbot(prompt)
            except Exception as e:
                return 'fail', (chatbot, '\n'.join(str(x) for x in e.args))
            log['raw response'].append(response)
            log['context_log'].append(copy.deepcopy(chatbot.context))
            
            # check if finished
            if 'Done' in response:
                return 'end', (chatbot,)
            # extract the function name
            coarse_function_names = re.findall(r'(?<=@)([A-Z].*?)\(.*?\)(?=@|\n|$)', response)
            
            print("Extracted API at coarse stage:", coarse_function_names)
            
            # check if there is any api detected
            if not coarse_function_names:
                return 'no_api_detected', (chatbot, context_index, 'state1', False)
            # check if the function name is in the api list
            invalid_api = []
            for i in range(len(coarse_function_names)):
                for api_candidate in self.api_list:
                    # LLms may confuse letter cases so we handle this here to avoid unnecessary runtime excpetions.
                    if coarse_function_names[i].lower() == api_candidate.lower():
                        coarse_function_names[i] = api_candidate
                        break
                else:
                    invalid_api.append(coarse_function_names[i])
            
            coarse_function_names = set(coarse_function_names) # Remove dulicate coarse function names to prepare for the next stage (querying the external doc)
            print("Legal APIs: ", coarse_function_names)
            
            if invalid_api:
                return 'invalid_api', (chatbot, invalid_api, context_index, 'state1', False)
            
            # clear the error count
            self.error_count = 0
            # go to fine-grained state
            return 'state2', (chatbot, response, prompt, coarse_function_names, context_index)

        cycles_times = 0
        async def state2(chatbot: ChatGPT, response, prompt, coarse_function_names, context_index):
            nonlocal cycles_times
            # Preventing dead loops
            cycles_times += 1
            if cycles_times > self.max_cycle_times:
                return 'fail', (chatbot, 'Too many cycles')
            
            if self.use_ext_doc:
                print(Fore.CYAN + "Processing state2: Referring to external documents"); print(Style.RESET_ALL)
                prompt_for_fine = chatbot.context[context_index]['content']
                chatbot.context = chatbot.context[:context_index]
                # extract the function detailed doc
                api_detail_doc = '\n'.join([self.api_detail_doc[name] for name in coarse_function_names])
                # generate the prompt
                prompt = prompt_for_fine + self.prompt.get('fetch exterlnal doc', f'\nHere is the supplementary doc you can reference:\n{api_detail_doc}\nPlease use the above documents to generate the next step.\n')

                # clear error count
                self.error_count = 0
                # go to state3
                return 'state3', (chatbot, prompt, prompt_for_fine, coarse_function_names, context_index)
            else:
                print(Fore.CYAN + "Skipping state2: Inserting external documents"); print(Style.RESET_ALL)

                # clear error count
                self.error_count = 0
                return 'state4', (chatbot, response, prompt, prompt_for_fine, coarse_function_names, context_index)
            
        async def state3(chatbot: ChatGPT, prompt, base_prompt, coarse_function_names, context_index):
            print(Fore.CYAN + "Processing state3 - Finer-grained planning"); print(Style.RESET_ALL)
            nonlocal log, cycles_times
            # Preventing dead loops
            cycles_times += 1
            if cycles_times > self.max_cycle_times:
                return 'fail', (chatbot, f'Too many cycles ({self.max_cycle_times})')
            try:
                response = await chatbot(prompt)
            except Exception as e:
                return 'fail', (chatbot, ' '.join([str(x) for x in e.args]))
            log['intermediate response'].append(response)
            log['context_log'].append(copy.deepcopy(chatbot.context))
            if 'Done' in response:
                return 'end', (chatbot,)
            
            # extract the function name
            fine_function_names = re.findall(r'(?<=@)([A-Z].*?)\(.*?\)(?=@|\n|$)', response)

            print("Extracted API at fine stage:", fine_function_names)
            
            # check if there is any api detected
            if not fine_function_names:
                return 'no_api_detected', (chatbot, context_index, 'state3', base_prompt, coarse_function_names)
            
            # check if the function name is in the api list
            try:
                invalid_api = []
                for i in range(len(fine_function_names)):
                    for api_candidate in self.api_list:
                        if fine_function_names[i].lower() == api_candidate.lower():
                            response = response.replace("@"+fine_function_names[i], "@"+api_candidate)
                            fine_function_names[i] = api_candidate
                            break
                    else:
                        invalid_api.append(fine_function_names[i])

                fine_function_names = set(fine_function_names)
                print("Legal APIs: ", fine_function_names)
            except:
                print("Exception during checking APIs")
            if invalid_api:
                return 'invalid_api', (chatbot, invalid_api, context_index, 'state3', base_prompt, coarse_function_names)
            # check if all the fine-grained apis are in the coarse-grained apis
            if not fine_function_names.issubset(coarse_function_names):
                chatbot.context = chatbot.context[:context_index+1]
                return 'state2', (chatbot, response, base_prompt, fine_function_names, context_index)
            
            # clear error count
            self.error_count = 0
            # go to final state
            return 'state4', (chatbot, response, prompt, base_prompt, coarse_function_names, context_index)
        
        async def state4(chatbot: ChatGPT, response, prompt, base_prompt, coarse_function_names, context_index):
            print(Fore.CYAN + "Processing state4 - Executing"); print(Style.RESET_ALL)
            nonlocal log, cycles_times
            # extract the full function
            try:
                functions = re.findall(r'(?<=@)([A-Z].*?\))(?=@|\n|$)', response)
            except Exception as e:
                print(f"[State 4] Invalid syntax in the reponse: {response}")
                print(e)
                return 'syntax_error', (chatbot, context_index, 'state4', base_prompt, coarse_function_names)
                
            log['refined response'].append(functions)
            success, msg = self.Process(functions)
            if not success:
                # go to failing process state
                return 'execute_error', (chatbot, msg, prompt, base_prompt, coarse_function_names, context_index)
            
            # Clear the cycles times
            cycles_times = 0

            chatbot.context[context_index]['content'] = base_prompt
            chatbot.context[context_index+1]['content'] = response
            context_index += 2
            chatbot.context = chatbot.context[:context_index]
            
            # go to state 1
            next_step_prompt = self.prompt.get('next step planning', None)
            if next_step_prompt is None:
                next_step_prompt = """If task is not finished, please provide the next step (add @ both before and after each API call, like "Action API: @Write(range=..., value=...)@"); otherwise, please type "Done!". Do select an API from the API document and provide concise explanation of your choice."""
            prompt = self.GetSheetState() + '\n' + next_step_prompt
            
            # clear error count
            self.error_count = 0
            # prompt = 'If task is not finished, please provide the next step, otherwise, please type "Done!".'
            return 'state1', (chatbot, prompt, context_index, True)
        
        async def execute_error(chatbot: ChatGPT, msg, prompt, base_prompt, coarse_function_names, context_index):
            nonlocal cycles_times
            # Preventing dead loops
            cycles_times += 1
            if cycles_times > self.max_cycle_times:
                return 'fail', (chatbot, f'Too many cycles (>{self.max_cycle_times})')
            print("processing execute_error")
            prompt = f'Execution error: {msg}\nPlease regenerate this step.'
            return 'state3', (chatbot, prompt, base_prompt, coarse_function_names, context_index)

        async def syntax_error(chatbot: ChatGPT, context_index, prev_state, base_prompt = None, coarse_function_names = None):
            print("processing output syntax_error")
            # prompt = 'Please return the API in one line. Please add @ both before and after the function call to indicate that the content between the two @ characters is a function call, like @Function1()@, Function2()@.'
            
            prompt = self.prompt.get('syntax error', None)
            if prompt is None:
                prompt = """Your answer does not conform with the output format specified in the requirements. Please generate this step again."""
            
            if prev_state == 'state1':
                return 'state1', (chatbot, prompt, context_index)
            elif prev_state in ['state3', 'state4']:
                return 'state3', (chatbot, prompt, base_prompt, coarse_function_names, context_index)
            
        async def no_api_detected(chatbot: ChatGPT, context_index, prev_state, base_prompt = None, coarse_function_names = None):
            print("processing no_api_detected")
            self.error_count += 1
            if self.error_count >= self.max_error_count:
                return 'fail', (chatbot, f'Too many error counts (>= {self.max_error_count})')
            # prompt = 'Please return the API in one line. Please add @ both before and after the function call to indicate that the content between the two @ characters is a function call, like @Function1()@, Function2()@.'
            
            prompt = self.prompt.get('no api correction', None)
            if prompt is None:
                prompt = """Your response contains no APIs! Please regenerate this step in one line. Please add @ both before and after the atomic action to indicate that the content between the two @ characters is an API call, e.g. Action API: @CopyPaste(range=..., value=...)@."""
            
            if prev_state == 'state1':
                return 'state1', (chatbot, prompt, context_index)
            elif prev_state == 'state3':
                return 'state3', (chatbot, prompt, base_prompt, coarse_function_names, context_index)
        
        async def invalid_api(chatbot: ChatGPT, invalid_api, context_index, prev_state, base_prompt = None, coarse_function_names = None):
            print("processing invalid_api")
            self.error_count += 1
            if self.error_count >= self.max_error_count:
                return 'fail', (chatbot, f'Too many error counts (>= {self.max_error_count})')
            
            prompt = f'There is no API: {" ".join(invalid_api)}\n. You can only choose from the following APIs:\n{" ".join(self.api_list)}\nPlease regenerate this step.'
            if prev_state == 'state1':
                return 'state1', (chatbot, prompt, context_index, False)
            elif prev_state == 'state3':
                return 'state3', (chatbot, prompt, base_prompt, coarse_function_names, context_index)

        async def fail(chatbot: ChatGPT, msg):
            self.step = 1
            print("processing fail", msg)
            nonlocal log, cycles_times
            # Clear the cycles times
            cycles_times = 0
            log['msg'] = msg
            return False, log
        
        async def end(chatbot: ChatGPT):
            self.step = 1
            print(Fore.CYAN + "processing end"); print(Style.RESET_ALL)
            nonlocal log
            return True, log
        
        async def start(chatbot: ChatGPT, prompt, context_index):
            nonlocal log
            log = {
                'raw response': [],
                'intermediate response': [],
                'refined response': [],
                'context_log': [],
                'msg': []
            }
            return 'state1', (chatbot, prompt, context_index, True)
        
        statemachine = StateMachine()
        statemachine.add_state('start', start)
        statemachine.add_state('state1', state1)
        statemachine.add_state('execute_error', execute_error)
        statemachine.add_state('syntax_error', syntax_error)
        statemachine.add_state('no_api_detected', no_api_detected)
        statemachine.add_state('invalid_api', invalid_api)
        statemachine.add_state('state2', state2)
        statemachine.add_state('state3', state3)
        statemachine.add_state('state4', state4)
        statemachine.add_state('fail', fail, end_state=True)
        statemachine.add_state('end', end, end_state=True)
        statemachine.set_start('start')
        return statemachine


    def Process(self, actions: list) -> bool:
        """
        Executes a list of actions, where each action is a tuple of the function name and its arguments.

        Parameters:
            actions (list): a list of tuples, where each tuple contains the name of the function to be executed and its arguments.

        Returns:
            None.
        """
        for func in actions:
            print('Trying to execute {} on {}'.format(func, self._backend.activeWB.Name), '\n')
            try:
                func = func.replace('\\', '')
                eval('self._backend.{}'.format(func))
            except Exception as e:
                print('Error: {}'.format(e))
                return False, 'Failed to execute {}.\nError: {}\n'.format(func, e)
            
        return True, None

    def ProcessSingleAction(self, funcName: str, funcParam: tuple) -> None:
        """
        Executes a single action, which is a function with its arguments.

        Parameters:
            funcName (str): the name of the function to be executed.
            funcParam (tuple): the arguments to be passed to the function.

        Returns:
            None.
        """
        func = eval('self._backend.{}'.format(funcName))
        func(*funcParam)
    
    def ProcessMT(self, actions: list) -> None:
        """
        Executes a list of actions using multithreading.

        Parameters:
            actions (list): a list of tuples, where each tuple contains the name of the function to be executed and its arguments.

        Returns:
            None.
        """
        import pythoncom
        pythoncom.CoInitialize()
        self._backend = xwBackend()
        self.Process(actions)
        pythoncom.CoUninitialize()

    def ProcessProxy(self, actions: list, url: str = 'http://10.211.55.3:8888/') -> None:
        """
        Sends a list of actions to a server specified by url.

        Parameters:
            actions (list): a list of tuples, where each tuple contains the name of the function to be executed and its arguments.
            url (str): the URL of the server.

        Returns:
            None.
        """
        actions = [[elem[0], list(elem[1]) if isinstance(elem[1], tuple) else [elem[1]]] for elem in actions]
        self.SendToServer({'actions': actions}, url)

    def StartServer(self, port: int = 8888) -> None:
        """
        Starts a server at a specified port.

        Parameters:
            port (int): the port number for the server.

        Returns:
            None.
        """
        from fastapi import FastAPI
        import uvicorn
        from pydantic import BaseModel

        class Item(BaseModel):
            actions: list

        app = FastAPI()

        @app.get("/")
        def read_root():
            return {"Hello": "World"}

        @app.post('/')
        def HandlePOST(items: Item):
            actions = [(elem[0], tuple(elem[1])) for elem in items.actions]
            self.ProcessMT(actions)
            return None
        
        uvicorn.run(app, host='0.0.0.0', port=port)

    def SendToServer(self, payload: dict, url: str) -> None:
        """
        Send a payload to a server via HTTP POST request.

        Args:
            payload (dict): The payload to be sent to the server.
            url (str): The URL of the server.

        Returns:
            None: The response JSON object from the server.
        """
        response = requests.post(url=url, json=payload)
        return response.json()
    
    def GetSheetState(self) -> str:
        """
        Gets the current state of the sheet.

        Returns:
            str: the current state of the sheet.
        """
        return self._backend.GetSheetsState(add_example_data2feedback=self.add_example_data2feedback)
    
    async def ExtractActions(self, document: str) -> str:
        prompt = self.prompt['extract actions'].copy()
        chatbot = ChatGPT(self.agent_config['ChatGPT_1' if self.use_same_LLM else 'ChatGPT_2'], prompt, interaction_mode=self.interaction_mode)
        prompt = 'Document:\n' + document
        try:
            res = await chatbot(prompt)
            functions = re.findall(r'- (.*)', res)
        except Exception as e:
            print(f"error occurs when parsing response: {e}")
        else:
            return functions, res
        
    # async def Instruction(self, context: str, instruction: str, file: str = None, savepath: str = None) -> Tuple[bool, List[str]]:
    #     """
    #     Executes an instruction on the sheet.

    #     Parameters:
    #         instruction (str): the instruction to be executed.
    #         file (str): the path to the sheet.

    #     Returns:
    #         None.
    #     """
    #     if file is not None:
    #         self._backend.OpenWorkbook(file)
    #     prompt = self.prompt['task planning'].copy()
    #     prompt[-2] = prompt[-2].copy()
    #     prompt[-2]['content'] = prompt[-2]['content'].format(context=context, instruction=instruction)
    #     chatbot = ChatGPT(self.agent_config['ChatGPT_1'], prompt)
    #     log = {
    #         'raw response': [],
    #         'intermediate response': [],
    #         'refined response': []
    #     }
    #     while True:
    #         sheetstate = self.GetSheetState()
    #         response = await chatbot(sheetstate)
    #         log['raw response'].append(response)
    #         if 'Done' in response:
    #             break
    #         refined_res, intermediate_res = await self.ExtractActions(response)
    #         if not refined_res:
    #             if savepath is not None:
    #                 print('closing workbook {}'.format(self._backend.activeWB.Name))
    #                 self._backend.activeWB.Close(SaveChanges=False)
    #             return False, log
    #         log['intermediate response'].append(intermediate_res)
    #         log['refined response'].append(refined_res)
    #         success, msg = self.Process(refined_res)
    #         if not success:
    #             log['error'] = msg
    #             if savepath is not None:
    #                 print('closing workbook {}'.format(self._backend.activeWB.Name))
    #                 self._backend.activeWB.Close(SaveChanges=False)
    #             return False, log
    #         time.sleep(0.1)
    #     if savepath is not None:
    #         self._backend.SaveWorkbook(savepath)
    #         self._backend.activeWB.Close()
        
    #     return True, log
    
    async def Instruction2(self, context: str, instruction: str, file: str = None, savepath: str = None) -> Tuple[bool, List[str]]:
        """
        Executes an instruction on the sheet.

        Parameters:
            instruction (str): the instruction to be executed.
            file (str): the path to the sheet.

        Returns:
            None.
        """
        if file is not None:
            time.sleep(0.5)
            self._backend.OpenWorkbook(file)
        base_prompt = self.prompt['task planning'].copy()
        base_prompt[0] = base_prompt[0].copy()
        
        if self.use_oracle_API_doc:
            # Load an answer
            sheet_id_name = os.path.basename(file)[:-12] # len("_source.xlsx") = 12
            api_doc = []
            oracle_APIs = self.oracle_API_dict[sheet_id_name][:]
            
            for line in self.api_usage.split('\n'):
                for API in oracle_APIs:
                    if API in line:
                        api_doc.append(line)
                        oracle_APIs.remove(API)
                        break
            
            api_doc = '\n'.join(api_doc)
        else:
            api_doc = self.api_usage
        
        if self.use_doc_in_syspromt:
            base_prompt[0]['content'] = base_prompt[0]['content'].format(API_Doc=f"\nHere is the API document:\n{api_doc}\n")
        else:
            base_prompt[0]['content'] = base_prompt[0]['content'].format(API_Doc="")
        
        prompt = base_prompt.pop()['content']
        
        sheet_state = self.GetSheetState()
        print(50*'-' + '\n' + sheet_state + '\n' + 50*'-')
        prompt = prompt.format(context=context, instruction=instruction, sheet_state=sheet_state)
        chatbot = ChatGPT(self.agent_config['ChatGPT_1'], base_prompt, interaction_mode=self.interaction_mode)
        context_index = len(chatbot.context)
        success, log = await self.statemachine.run((chatbot, prompt, context_index))
        if savepath is not None:
            self._backend.SaveWorkbook(savepath)
            self._backend.activeWB.Close()
        else:
            # self._backend.activeWB.Close(SaveChanges=False)
            pass

        return success, log
        

async def chat_without_save_context(chatbot, prompt):
    response = await chatbot(prompt)
    chatbot.context.pop()
    chatbot.context.pop()
    return response

def find_APIs(response, api_list):
    function_names = re.findall(r'Action API: (.*?)\(', response)
    function_names = set(function_names)
    # check if the function is valid
    invalid_functions = []
    for name in function_names:
        if name not in api_list:
            invalid_functions.append(name)
    if invalid_functions:
        print(f'function {invalid_functions} is not valid')
        function_names = function_names - set(invalid_functions)


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description='Process config.')
    parser.add_argument('--config', '-c', type=str, help='path to config file')
    args = parser.parse_args()

    with open(args.config, 'r') as f:
        config = yaml.load(f, Loader=yaml.Loader)
        
    agent = Agent(config['Agent'])
    while True:
        instruction = input('Enter your instruction: \n')
        agent.Instruction2('', instruction)
        