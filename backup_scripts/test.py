import numpy as np

import random

# Define the length of the array
array_length = 10000

# Create an empty array
random_array = []

# Generate random numbers and append them to the array
for _ in range(array_length):
    random_number = random.random()  # Generates a random float between 0 and 1
    random_array.append(random_number)

# Print the resulting array
alphas=list(map(lambda x: (1-x)/x, random_array))

all(e >0 for e in alphas)
