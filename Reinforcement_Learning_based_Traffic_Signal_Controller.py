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