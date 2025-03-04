# Load the file

import win32com.client
import numpy as np
import random

# Connect to Vissim
Vissim = win32com.client.Dispatch("Vissim.Vissim")

# Load your Vissim network
Vissim.LoadNet(r"C:\Users\abdul\OneDrive - BUET\Thesis\Final Final Final\New Market\Present Condition\New Market at present.inpx")
Vissim.LoadLayout(r"C:\Users\abdul\OneDrive - BUET\Thesis\Final Final Final\New Market\Present Condition\New Market at present.layx")
# Load the Scenario
Vissim.ScenarioManagement.LoadScenario(5)

# Define the action space (phase and green time combinations)
phases = [0, 1, 2, 3]  # 0: North, 1: East, 2: South, 3: West
green_times = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120]  # Green times in seconds
actions = [(p, g) for p in phases for g in green_times]
num_actions = len(actions)

# Define the state space (queue lengths for each approach)
state_size = 4  # Queue lengths for each leg

# Q-table (assuming 100 states and the size of the action space)
q_table = np.zeros((100, num_actions))  # 100 states and num_actions columns

alpha = 0.1     # Learning rate
gamma = 0.9     # Discount factor
epsilon = .1   # Exploration rate
#epsilon_min = 0.01
#epsilon_decay = 0.995

# Get queue lengths for each leg
def get_queue_lengths():
    queue_lengths = []
    for i in range(1, 5):  # There are 4 legs, each with its own queue counter
        queue_length = Vissim.Net.QueueCounters.ItemByKey(i).AttValue('QLen(Current,1)')
        queue_lengths.append(queue_length)
    return queue_lengths

# Set the signal phase and green time in VISSIM
def set_signal_phase(phase, green_time):
    SignalController = Vissim.Net.SignalControllers.ItemByKey(1)  # Intersection Signal Controller ID

    # Set all phases to red initially
    for i in range(1, 5):
        SignalController.SGs.ItemByKey(i).SetAttValue("SigState", "RED")

    # Set the selected phase to green
    SignalController.SGs.ItemByKey(phase + 1).SetAttValue("SigState", "GREEN")

    # Keep the signal green for the selected green time
    start_time = Vissim.Simulation.AttValue('SimSec')
    end_time = start_time + green_time
    while Vissim.Simulation.AttValue('SimSec') < end_time:
        Vissim.Simulation.RunSingleStep()

    # Set the selected phase to amber
    SignalController.SGs.ItemByKey(phase + 1).SetAttValue("SigState", "AMBER")

    # Keep the signal amber for 3 seconds
    start_time = Vissim.Simulation.AttValue('SimSec')
    end_time = start_time + 3
    while Vissim.Simulation.AttValue('SimSec') < end_time:
        Vissim.Simulation.RunSingleStep()

# Choose the best action (phase and green time) based on the current state
def choose_action(state):
    if np.random.rand() <= epsilon:
        return random.choice(actions)  # Explore
    else:
        return actions[np.argmax(q_table[state])]  # Exploit

# Update the Q-table
def update_q_table(state, action_idx, reward, next_state):
    best_next_action = np.argmax(q_table[next_state])
    q_table[state, action_idx] += alpha * (reward + gamma * q_table[next_state, best_next_action] - q_table[state, action_idx])

# Discretize the state based on queue lengths
def get_state(queue_lengths):
    return int(sum(queue_lengths) // 10)  # Example: Discretize by summing and dividing

# Calculate the reward based on queue length reduction
def get_reward(previous_queues, current_queues):
    return sum(previous_queues) - sum(current_queues)


# Q-learning loop
for episode in range(30):  # Number of episodes
    print(f'Episode: {episode}')

    Vissim.Simulation.RunSingleStep()

    state = get_state(get_queue_lengths())
    action = choose_action(state)
    action_idx = actions.index(action)

    previous_queues = get_queue_lengths()

    # Set the signal phase and green time in VISSIM
    phase, green_time = action
    set_signal_phase(phase, green_time)

    current_queues = get_queue_lengths()
    reward = get_reward(previous_queues, current_queues)

    next_state = get_state(current_queues)
    update_q_table(state, action_idx, reward, next_state)

    state = next_state  # Move to the next state

    #if epsilon > epsilon_min:
       # epsilon *= epsilon_decay

print(q_table)
