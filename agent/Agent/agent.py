from .xwAPI import xwBackend
from typing import Tuple, List
import time
import re, os
from colorama import Fore, Style
from utils import ChatGPT, StateMachine
import yaml
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
        self.use_ext_doc = self.agent_config.get('use_ext_doc', True)
        self.use_same_LLM = self.agent_config.get('use_same_LLM', True)
        self.interaction_mode = config.get('interaction_mode', False)
        self.max_cycle_times = self.agent_config['max_cycle_times']
        self.max_error_count = self.agent_config['max_error_count']

        with open(self.agent_config['prompt_path'], 'r') as f:
            self.prompt = yaml.load(f, Loader=yaml.Loader)
        
        with open(self.agent_config['api_doc_path']) as f:
            self.api_doc = yaml.load(f, Loader=yaml.FullLoader)

        self.api_list, self.api_usage, self.api_detail_doc = get_api_doc(self.prompt_format, self.api_doc)
        self.statemachine = self.InitStateMachines()

        if self.agent_config['API_backend'] == 'xw':
            self._backend = xwBackend(self.agent_config['APP_backend'], self.api_doc)
        else:
            raise NotImplementedError('Backend {} is not implemented.'.format(self.agent_config['backend']))
        
        print("Initializing SheetCopilot...")
        print(f"-> Use the external doc: {self.use_ext_doc}\n-> Use the same LLM for planning and parsing: {self.use_same_LLM}")

        self.step = 1
        self.error_count = 0

    def InitStateMachines(self) -> StateMachine:

        log = {}
        async def state1(chatbot: ChatGPT, prompt, context_index, new_step=False):
            if new_step:
                print(Fore.YELLOW + f"Step {self.step} ... ", end='')
                self.step += 1
            
            print(Fore.CYAN + "\nProcessing state1 - Coarse-grained planning" + Style.RESET_ALL)
            
            nonlocal log, cycles_times
            # Preventing dead loops
            cycles_times += 1
            if cycles_times > self.max_cycle_times:
                return 'fail', (chatbot, 'State 1', f'Too many cycles (> {self.max_cycle_times})')
            try:
                response = await chatbot(prompt)
            except Exception as e:
                return 'fail', (chatbot, 'State 1', '\n'.join(str(x) for x in e.args))
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
                return 'no_api_detected', (chatbot, context_index)
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
                return 'invalid_api', (chatbot, invalid_api, response)
            
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
                return 'fail', (chatbot, 'State 2', 'Too many cycles')
            
            if self.use_ext_doc:
                print(Fore.CYAN + "\nProcessing state2: Referring to external documents" + Style.RESET_ALL)
                prompt_for_fine = prompt # chatbot.context[context_index]['content']
                chatbot.context = chatbot.context[:context_index]
                # extract the function detailed doc
                api_detail_doc = '\n'.join([self.api_detail_doc[name] for name in coarse_function_names])
                # generate the prompt
                prompt = prompt_for_fine + self.prompt.get('fetch exterlnal doc', '\nHere is the supplementary doc you can reference:\n{api_detail_doc}\nPlease use the above documents to generate the next step.\n').replace("{api_detail_doc}", api_detail_doc).replace("{chosen_apis}", ', '.join(coarse_function_names))

                # clear error count
                self.error_count = 0
                # go to state3
                return 'state3', (chatbot, prompt, prompt_for_fine, coarse_function_names, context_index)
            else:
                print(Fore.CYAN + "Skipping state2: Inserting external documents" + Style.RESET_ALL)

                # clear error count
                self.error_count = 0
                return 'state4', (chatbot, response, prompt, prompt_for_fine, coarse_function_names, context_index)
            
        async def state3(chatbot: ChatGPT, prompt, base_prompt, coarse_function_names, context_index):
            print(Fore.CYAN + "\nProcessing state3 - Finer-grained planning" + Style.RESET_ALL)
            nonlocal log, cycles_times
            # Preventing dead loops
            cycles_times += 1
            if cycles_times > self.max_cycle_times:
                return 'fail', (chatbot, 'State 3', f'Too many cycles ({self.max_cycle_times})')

            try:
                response = await chatbot(prompt)
            except Exception as e:
                return 'fail', (chatbot, 'State 3', ' '.join([str(x) for x in e.args]))
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
                return 'no_api_detected', (chatbot, response)
            
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
                return 'invalid_api', (chatbot, invalid_api, response)
            # check if all the fine-grained apis are in the coarse-grained apis
            if not fine_function_names.issubset(coarse_function_names):
                chatbot.context = chatbot.context[:context_index+1]
                return 'state2', (chatbot, response, base_prompt, fine_function_names, context_index)
            
            # clear error count
            self.error_count = 0
            # go to final state
            return 'state4', (chatbot, response, prompt, base_prompt, coarse_function_names, context_index)
        
        async def state4(chatbot: ChatGPT, response, prompt, base_prompt, coarse_function_names, context_index):
            print(Fore.CYAN + "\nProcessing state4 - Executing" + Style.RESET_ALL)
            nonlocal log, cycles_times
            # extract the full function
            try:
                functions = re.findall(r'(?<=@)([A-Z].*?\))(?=@|\n|$)', response)
            except Exception as e:
                print(f"[State 4] Invalid syntax in the reponse: {response}")
                print(e)
                return 'syntax_error', (chatbot, response)
                
            log['refined response'].append(functions)
            success, msg = self.Process(functions)
            if not success:
                # go to failing process state
                return 'execute_error', (chatbot, msg)
            
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
        
        async def execute_error(chatbot: ChatGPT, msg: str):
            print("\nProcessing execute_error")
            return 'fail', (chatbot, 'execute_error', f'Execution error: {msg}')

        async def syntax_error(chatbot: ChatGPT, response: str):
            print("\nProcessing output syntax_error")
            return 'fail', (chatbot, 'syntax_error', f'Syntax errors found in the response: {response}')
            
        async def no_api_detected(chatbot: ChatGPT, response: str):
            print("\nProcessing no_api_detected")
            return 'fail', (chatbot, 'no_api_detected', f'No API detected in the response: {response}')
        
        async def invalid_api(chatbot: ChatGPT, invalid_api, response: str):
            print("\nProcessing invalid_api")
            return 'fail', (chatbot, 'Invalid API handling', f'Invalid APIs: {invalid_api} were found in the response: {response}')

        async def fail(chatbot: ChatGPT, prev_state, msg):
            self.step = 1
            print(Fore.RED + f"{prev_state} failed. Cause: {msg}" + Style.RESET_ALL)
            nonlocal log, cycles_times
            # Clear the cycles times
            cycles_times = 0
            log['msg'] = msg
            return False, log
        
        async def end(chatbot: ChatGPT):
            self.step = 1
            chatbot.reset_query_count()

            print(Fore.CYAN + "\nProcessing end" + Style.RESET_ALL)
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

    def GetSheetState(self) -> str:
        """
        Gets the current state of the sheet.

        Returns:
            str: the current state of the sheet.
        """
        return self._backend.GetSheetsState()
    
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
        api_doc = self.api_usage
        
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
        agent.Instruction('', instruction)
        