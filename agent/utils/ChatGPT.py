import openai
import itertools
import asyncio

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
        self._query_count = 1
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

    def reset_query_count(self):
        self._query_count = 1

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
                
                if 'exceed the limit' in result['content']:
                    raise Exception(f"Generated tokens exceed the limit: {self.max_total_tokens}!")
            except asyncio.TimeoutError:
                print("API call timed out after {} seconds. Retring {}/{}...".format(self.timeout, i+1, self.max_retries))
            except openai.RateLimitError  as e:
                print("API call rate limited. Retring {}/{}...\n{}".format(i+1, self.max_retries, e))
            except openai.APIError:
                print("API call failed. Retring {}/{}...".format(i+1, self.max_retries))
                # time.sleep(20)
            except Exception as e:
                raise e
            else:
                self._query_count += 1

                self.context.append(result)
                return result['content']

            await asyncio.sleep(30)

        raise Exception("API call failed after {} retries".format(self.max_retries))
    
    async def __request__(self) -> str:
        
        # Querying...
        # Use acreate for interaction will raise OpanAI communication errors
        if self.interaction_mode:
            response = response = await openai.AsyncOpenAI(api_key = next(self.api_keys)).chat.completions.create(
                model = self.model_name,
                messages = self.context,
                temperature = self.temperature,
            )
        else:
            response = await openai.AsyncOpenAI(api_key = next(self.api_keys)).chat.completions.create(
                model = self.model_name,
                messages = self.context,
                temperature = self.temperature,
            )
        if response.usage.total_tokens > self.max_total_tokens:
            raise Exception(f"Generated tokens ({response.usage.total_tokens}) exceed the limit: {self.max_total_tokens}!")

            
        # Process responses
        processed = {
        'role': response.choices[0].message.role,
        'content': response.choices[0].message.content,
        }
        
        return processed

async def test():
    chatbot = ChatGPT()
    while True:
        prompt = input("You: ")
        response = await chatbot(prompt)
        print("Bot:", response)

if __name__ == "__main__":
    asyncio.run(test())