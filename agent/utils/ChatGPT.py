import openai
import datetime
import time
import itertools
import asyncio
import yaml, json
import requests
import re

default_config = {
    "api_base": "https://api.openai.com/v1",
    "model_name": "gpt-3.5-turbo",
    "api_keys": [
        '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',
    ],
    "max_retries": 3,
    "timeout": 20,
    "temperature": 0.0,
}

headers = {
            'Content-Type': 'application/json'
            }

class ChatGPT:
    
    def __init__(self, config = default_config, context = [], interaction_mode=False):
        self.max_retries = config['max_retries']
        self.timeout = config['timeout']
        self.model_name = config['model_name']
        self.prompt_format = config['prompt_format']
        self.api_keys = itertools.cycle(config['api_keys'])
        self.max_total_tokens = config['max_total_tokens']
        self.context = context
        self.interaction_mode = interaction_mode
        openai.api_base = config['api_base']

        if 'gpt' not in self.model_name.lower() and "toolbench" not in self.model_name.lower():
            payload = json.dumps({'temp': config['temperature'], 'topk': config['topk'], 'max_new_tokens': config['max_new_tokens']})
            feedback = requests.request("POST", openai.api_base.replace("generate", "set_generation_config"), headers=headers, data=payload)
            print("Setting generation config result: ", feedback.text.strip('"'))
        
        self.temperature = config['temperature']
        
        print("=============== Initializing the agent ===============")
        print("=   Model: {}".format(self.model_name))
        print("=   API Keys: {}".format(config['api_keys']))
        print("=   Max Retries: {}".format(self.max_retries))
        print("=   Timeout: {}".format(self.timeout))
        print("=   Max Total Tokens: {}".format(self.max_total_tokens))
        print("=   Temperature: {}".format(self.temperature))
        print("======================================================")
        # openai.api_key = next(self.api_keys)

    async def __call__(self, prompt) -> str:
        self.context.append(
                {
                    "role": "user",
                    "content": prompt
                }
            )
        return await self.__get_response__()

    async def __get_response__(self) -> str:
        for i in range(self.max_retries):
            try:
                result = await asyncio.wait_for(self.__request__(), self.timeout)
            except asyncio.TimeoutError:
                print("API call timed out after {} seconds. Retring {}/{}...".format(self.timeout, i+1, self.max_retries))
            except openai.error.RateLimitError as e:
                print("API call rate limited. Retring {}/{}...\n{}".format(i+1, self.max_retries, e))
            except openai.error.APIError:
                print("API call failed. Retring {}/{}...".format(i+1, self.max_retries))
                # time.sleep(20)
            except Exception as e:
                raise e
            else:
                # self.context.pop()
                self.context.append(result)
                return result['content']

            await asyncio.sleep(30)

        raise Exception("API call failed after {} retries".format(self.max_retries))
    
    async def __request__(self) -> str:
        
        # Querying...
        if "gpt" in self.model_name.lower():
            # Use acreate for interaction will raise OpanAI communication errors
            if self.interaction_mode:
                response = openai.ChatCompletion.create(
                    model = self.model_name,
                    messages = self.context,
                    temperature = self.temperature,
                    api_key = next(self.api_keys)
                )
            else:
                response = await openai.ChatCompletion.acreate(
                    model = self.model_name,
                    messages = self.context,
                    temperature = self.temperature,
                    api_key = next(self.api_keys),
                    stream=True
                )
            if response.usage.total_tokens > self.max_total_tokens:
                raise Exception(f"Token usage exceeded max_total_tokens ({self.max_total_tokens}), used {response.usage.total_tokens}")

        elif "toolbench" in self.model_name.lower():
            prompt = '[' + ','.join(['{' + '"role": "{}", "content": """{}"""'.format(x["role"], x["content"]) + '}' for x in self.context]) + ']'
            
            payload = json.dumps({
            "text": prompt,
            "top_k": 10,
            "method": "CoT@1"
            })
            
            while True:
                try:
                    print("Querying the LLM... at ", openai.api_base)
                    response = requests.request("POST", openai.api_base, headers=headers, data=payload, timeout=60)
                    if response.text == "":
                        raise Exception("empty response!")
                except Exception as e:
                    print(e.args[0]); continue
                
                break
            
        else:
            prompt = '[' + ','.join(['{' + '"role": "{}", "content": """{}"""'.format(x["role"], x["content"]) + '}' for x in self.context]) + ']'
            
            payload = json.dumps({
            "prompt": prompt
            })
            
            while True:
                try:
                    print("Querying the LLM...")
                    response = requests.request("POST", openai.api_base, headers=headers, data=payload)
                    if response.text == "":
                        raise Exception("empty response!")
                except Exception as e:
                    print(e.args[0]); continue
                
                break
            
        # Process responses
        if "gpt" in self.prompt_format.lower():
            processed = response.choices[0].message.to_dict()
        elif "toolllama" in self.prompt_format.lower():
            # split_Lid = response.text.rfind('Thought')
            # if split_Lid == -1:
            #     split_Lid = response.text.rfind('thought:')
            
            # split_Rid = response.text.rfind("}")
            # try:
            #     split = '{' + response.text[split_Lid:split_Rid]
            # except Exception as e:
            #     print(e)
            # answer_dict = eval(split.strip('" ')) #.encode().decode('unicode_escape')
            
            # thought = answer_dict['thought']
            # API_name = answer_dict['name']
            # raw_args = eval(answer_dict['content'])
            
            # print(f"Querying finished. Returned raw texts:\n{split}")
            
            # Reformat the response according to the ToolLlama format
            # Example:
            # Thought: Step 1. Create a new column D.
            # Action: Write
            # Action Input: {
            # "range": "Sheet1!D1",
            # "value": "Revenue"
            # }
            # pattern = r'(?P<key>[^:]+): (?P<value>[^{}]*(?:{[^{}]*}|\{(?<DEPTH>)|\})*?)'
            # matches = re.findall(pattern, response.text)
            
            # args = eval(matches[2][1])
            # reformated_response = "Thought: {}\nAction: {}\nAction Input:".format(matches[0][1], matches[0][1]) + " {\n" + ',\n'.join(['"{}": "{}"'.format(k, v) for k, v in args.items()]) + "\n}"
            
            response_text = response.text[:response.text.rfind('}') + 1].strip('" ').encode().decode('unicode_escape')
            processed = {"role": "assistant", "content": response_text}
        else:
            cleaned_text = response.text.strip('" ').encode().decode('unicode_escape')
            
            print(f"Querying finished. Returned raw texts:\n{response.text}.\nAfter cleaning:\n{cleaned_text}")
            
            processed = {"role": "assistant", "content": cleaned_text}
        
        return processed

async def test():
    chatbot = ChatGPT()
    while True:
        prompt = input("You: ")
        response = await chatbot(prompt)
        print("Bot:", response)

if __name__ == "__main__":
    asyncio.run(test())