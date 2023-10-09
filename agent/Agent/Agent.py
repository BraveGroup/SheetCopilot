from .xwAPI import xwBackend
from typing import Tuple, List, Optional
import time
import requests
import openai
import re
import sys
from utils import ChatGPT, StateMachine
import yaml, json
import asyncio
import copy

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
        self.config = config
        self.model_name = config['ChatGPT_1']['model_name']
        self.prompt_format = config['ChatGPT_1']['prompt_format']
        self.use_ext_doc = config['use_ext_doc']
        self.use_same_LLM = config['use_same_LLM']
        with open(self.config['prompt_path'], 'r') as f:
            self.prompt = yaml.load(f, Loader=yaml.Loader)
        
        with open(self.config['api_doc_path']) as f:
            self.api_doc = yaml.load(f, Loader=yaml.FullLoader)

        self.api_list, self.api_usage, self.api_detail_doc = get_api_doc(self.prompt_format, self.api_doc)
        self.statemachine = self.InitStateMachines()

        if config['API_backend'] == 'xw':
            self._backend = xwBackend(config['APP_backend'], self.api_doc)
        else:
            raise NotImplementedError('Backend {} is not implemented.'.format(config['Agent']['backend']))

    def InitStateMachines(self) -> StateMachine:

        log = {}
        async def state1(chatbot: ChatGPT, prompt, context_index):
            print("processing state1")
            nonlocal log, cycles_times
            # Preventing dead loops
            cycles_times += 1
            if cycles_times > 8:
                return 'fail', (chatbot, 'Too many cycles')
            
            if context_index > 30: # Only (limit - 10) // 2 steps are allowed
                return 'fail', (chatbot, 'Too many queries')
            
            try:
                response = await chatbot(prompt)
            except Exception as e:
                return 'fail', (chatbot, str(e.args[0]) + ' ' + str(e.args[1]))
            log['raw response'].append(response)
            log['context_log'].append(copy.deepcopy(chatbot.context))

            # check if finished
            if 'Done' in response: # 'Finish' is checked for ToolLLM
                return 'end', (chatbot,)
            
            # extract the function name
            coarse_function_names = re.findall(r'(?<=@)([A-Z].*?)\(.*?\)(?=@|\n|$)', response)
            print("Extracted API at coarse stage:", coarse_function_names)
            
            if 'Finish' in coarse_function_names: # 'Finish' is checked for ToolLLM
                return 'end', (chatbot,)
            
            # check if there is any api detected
            if not coarse_function_names:
                return 'no_api_detected', (chatbot, context_index, 'state1')
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
            print("Extracted coarse APIs correspond to the legal APIs: ", coarse_function_names)
            
            if invalid_api:
                return 'invalid_api', (chatbot, invalid_api, context_index, 'state1')
            
            # go to fine-grained state
            return 'state2', (chatbot, response, prompt, coarse_function_names, context_index)

        cycles_times = 0
        async def state2(chatbot: ChatGPT, response, prompt, coarse_function_names, context_index):
            nonlocal cycles_times
            # Preventing dead loops
            cycles_times += 1
            if cycles_times > 8:
                return 'fail', (chatbot, 'Too many cycles')
            
            if self.use_ext_doc:
                print("processing state2: Inserting external documents")
                prompt_for_fine = chatbot.context[context_index]['content']
                chatbot.context = chatbot.context[:context_index]
                # extract the function detailed doc
                api_detail_doc = '\n'.join([self.api_detail_doc[name] for name in coarse_function_names])
                # generate the prompt
                prompt = prompt_for_fine + f'\nHere is supplementary doc you can reference:\n{api_detail_doc}'
                # go to state3
                return 'state3', (chatbot, prompt, prompt_for_fine, coarse_function_names, context_index)
            else:
                print("Skipping state2: Inserting external documents")
                return 'state4', (chatbot, response, prompt, prompt_for_fine, coarse_function_names, context_index)
            
        async def state3(chatbot: ChatGPT, prompt, base_prompt, coarse_function_names, context_index):
            print("processing state3")
            nonlocal log, cycles_times
            # Preventing dead loops
            cycles_times += 1
            if cycles_times > 8:
                return 'fail', (chatbot, 'Too many cycles')
            try:
                response = await chatbot(prompt)
            except Exception as e:
                return 'fail', (chatbot, str(e.args[0]) + ' ' + str(e.args[1]))
            log['intermediate response'].append(response)
            log['context_log'].append(copy.deepcopy(chatbot.context))
            if 'Done' in response:
                return 'end', (chatbot,)
            
            # extract the function name
            fine_function_names = re.findall(r'(?<=@)([A-Z].*?)\(.*?\)(?=@|\n|$)', response)

            print("Extracted API at fine stage:", fine_function_names)
            
            if 'Finish' in fine_function_names: # 'Finish' is checked for ToolLLM
                return 'end', (chatbot,)
            
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
                print("Extracted fine APIs correspond to the legal APIs: ", fine_function_names)
            except:
                print("Exception during checking APIs")
            if invalid_api:
                return 'invalid_api', (chatbot, invalid_api, context_index, 'state3', base_prompt, coarse_function_names)
            # check if all the fine-grained apis are in the coarse-grained apis
            if not fine_function_names.issubset(coarse_function_names):
                chatbot.context = chatbot.context[:context_index+1]
                return 'state2', (chatbot, response, base_prompt, fine_function_names, context_index)
            
            # go to final state
            return 'state4', (chatbot, response, prompt, base_prompt, coarse_function_names, context_index)
        
        async def state4(chatbot: ChatGPT, response, prompt, base_prompt, coarse_function_names, context_index):
            print("processing state4")
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
                next_step_prompt = """If task is not finished, please provide the next step (add @ both before and after each API call, like "Action API: @Write(range=..., value=...)@); otherwise, please type "Done!". Do select an API from the API document. Keep concise and do not present explanations."""
            prompt = self.GetSheetState() + '\n' + next_step_prompt
            
            # prompt = 'If task is not finished, please provide the next step, otherwise, please type "Done!".'
            return 'state1', (chatbot, prompt, context_index)
        
        async def execute_error(chatbot: ChatGPT, msg, prompt, base_prompt, coarse_function_names, context_index):
            nonlocal cycles_times
            # Preventing dead loops
            cycles_times += 1
            if cycles_times > 8:
                return 'fail', (chatbot, 'Too many cycles')
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
            elif prev_state == 'state3' or prev_state == 'state4':
                return 'state3', (chatbot, prompt, base_prompt, coarse_function_names, context_index)
            
        async def no_api_detected(chatbot: ChatGPT, context_index, prev_state, base_prompt = None, coarse_function_names = None):
            print("processing no_api_detected")
            # prompt = 'Please return the API in one line. Please add @ both before and after the function call to indicate that the content between the two @ characters is a function call, like @Function1()@, Function2()@.'
            
            prompt = self.prompt.get('no api correction', None)
            if prompt is None:
                prompt = """Please return the API in one line. Please add @ both before and after the atomic action to indicate that the content between the two @ characters is an API call, like "Action API: @CopyPaste(range=..., value=...)@."""
            
            if prev_state == 'state1':
                return 'state1', (chatbot, prompt, context_index)
            elif prev_state == 'state3':
                return 'state3', (chatbot, prompt, base_prompt, coarse_function_names, context_index)
        
        async def invalid_api(chatbot: ChatGPT, invalid_api, context_index, prev_state, base_prompt = None, coarse_function_names = None):
            print("processing invalid_api")
            prompt = f'There is no API: {" ".join(invalid_api)}\n. You can only choose from the following APIs:\n{" ".join(self.api_list)}\nPlease regenerate this step.'
            if prev_state == 'state1':
                return 'state1', (chatbot, prompt, context_index)
            elif prev_state == 'state3':
                return 'state3', (chatbot, prompt, base_prompt, coarse_function_names, context_index)

        async def fail(chatbot: ChatGPT, msg):
            print("processing fail", msg)
            nonlocal log, cycles_times
            # Clear the cycles times
            cycles_times = 0
            log['msg'] = msg
            return False, log
        
        async def end(chatbot: ChatGPT):
            print("processing end")
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
            return 'state1', (chatbot, prompt, context_index)
        
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
                print('self._backend.{}'.format(func), '\n')
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

    def GetSheetState(self) -> str:
        """
        Gets the current state of the sheet.

        Returns:
            str: the current state of the sheet.
        """
        return self._backend.GetSheetsState()
    
    async def ExtractActions(self, document: str) -> str:
        prompt = self.prompt['extract actions'].copy()
        chatbot = ChatGPT(self.config['ChatGPT_1' if self.use_same_LLM else 'ChatGPT_2'], prompt)
        prompt = 'Document:\n' + document
        try:
            res = await chatbot(prompt)
            functions = re.findall(r'- (.*)', res)
        except Exception as e:
            print(f"error occurs when parsing response: {e}")
        else:
            return functions, res
        
    async def Instruction(self, context: str, instruction: str, file: str = None, savepath: str = None) -> Tuple[bool, List[str]]:
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
        base_prompt[0]['content'] = base_prompt[0]['content'].format(API_Doc=self.api_usage)
        prompt = base_prompt.pop()['content']
        prompt = prompt.format(context=context, instruction=instruction, sheet_state=self.GetSheetState())
        chatbot = ChatGPT(self.config['ChatGPT_1'], base_prompt)
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
    agent = Agent(xwBackend())
    while True:
        instruction = input('Enter your instruction: \n')
        agent.Instruction(instruction)
        