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
        '',
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
        
        self.temperature = config['temperature']
        
        print("=============== Initializing the LLM ===============")
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
        # Use acreate for interaction will raise OpanAI communication errors
        print("Querying ChatGPT ...")
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
            
        # Process responses
        processed = response.choices[0].message.to_dict()

        print(f"Querying finished. Response:\n{processed['content']}")
        
        return processed

async def test():
    chatbot = ChatGPT()
    while True:
        prompt = input("You: ")
        response = await chatbot(prompt)
        print("Bot:", response)

if __name__ == "__main__":
    asyncio.run(test())