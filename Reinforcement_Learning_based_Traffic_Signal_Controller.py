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

