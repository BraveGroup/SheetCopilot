class StateMachine:
    def __init__(self):
        self.handlers = {}
        self.start_state = None
        self.end_states = []

    def add_state(self, name, handler, end_state=False):
        name = name.upper()
        self.handlers[name] = handler
        if end_state:
            self.end_states.append(name)

    def set_start(self, name):
        self.start_state = name.upper()

    async def run(self, input):
        try:
            handler = self.handlers[self.start_state]
        except:
            raise Exception("must call .set_start() before .run()")
        if not self.end_states:
            raise Exception("at least one state must be an end_state")

        step_cnt = 0
        while True:
            (new_state, input) = await handler(*input)
            handler = self.handlers[new_state.upper()]
            if new_state.upper() in self.end_states:
                break

            step_cnt += 1

            if step_cnt >= 60:
                print("\033[0;33;40m**********************\n[Warning] Too many queries...\n**********************\033[0m\n")
            
        return await handler(*input)

# example usage
def state1(inputData):
    print("processing state1")
    newState = "state2"
    return (newState, inputData)

def state2(inputData):
    print("processing state2")
    newState = "state3"
    return (newState, inputData)

def state3(inputData):
    print("processing state3")
    return inputData

if __name__ == "__main__":
    m = StateMachine()
    m.add_state("state1", state1)
    m.add_state("state2", state2)
    m.add_state("state3", state3, end_state=True)
    m.set_start("state1")
    m.run("input")