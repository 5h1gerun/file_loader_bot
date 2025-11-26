
import pandas as pd
import numpy as np

# Create a wide dataframe
df = pd.DataFrame(np.random.randint(0, 100, size=(50, 50)), columns=[f"Col_{i}" for i in range(50)])

# Save to Excel
df.to_excel("wide_test.xlsx", index=False)
