import pandas as pd
import numpy as np

def generate_truncated_normal(lower, upper, size, loc=None, scale=None):
    """
    Generate `size` samples from a normal distribution truncated to [lower, upper].
    """
    loc = loc if loc is not None else (lower + upper) / 2
    scale = scale if scale is not None else (upper - lower) / 6
    samples = []
    while len(samples) < size:
        needed = size - len(samples)
        s = np.random.normal(loc, scale, size=needed * 2)
        samples.extend([x for x in s if lower <= x <= upper])
    return np.array(samples[:size])

def calculate_cpk(mean, std_dev, lsl, usl):
    """
    Calculate the Process Capability Index (Cpk).
    """
    cpu = (usl - mean) / (3 * std_dev)
    cpl = (mean - lsl) / (3 * std_dev)
    return min(cpu, cpl)

# 1. Read the CSV file
df = pd.read_csv('data/尺寸公差/KAP-7461中板-A76-50.csv', encoding='utf-8')

# 2. Rename columns
df = df.rename(columns={
    '项目': 'Description',
    '描述': 'StandardType',
    '标准值': 'Nominal',
    '上公差': 'TOL+',
    '下公差': 'TOL-'
})

# 3. Calculate USL and LSL
df['USL'] = df['Nominal'] + df['TOL+']
df['LSL'] = df['Nominal'] - df['TOL-']

# 4. Calculate the required standard deviation for Cpk ≥ 1.33
df['StdDev'] = (df['USL'] - df['Nominal']) / (3 * 1.33)

# 5. Generate 32 measurement values per row
for i, row in df.iterrows():
    lower, upper = row['LSL'], row['USL']
    nominal = row['Nominal']
    sigma = row['StdDev']
    vals = generate_truncated_normal(lower, upper, 32, loc=nominal, scale=sigma)
    df.loc[i, [f'Val_{j+1}' for j in range(32)]] = np.round(vals, 3)

# 6. Calculate mean and standard deviation for each row
df['Mean'] = df[[f'Val_{j+1}' for j in range(32)]].mean(axis=1)
df['StdDev'] = df[[f'Val_{j+1}' for j in range(32)]].std(axis=1)

# 7. Calculate Cpk for each row
df['Cpk'] = df.apply(lambda row: calculate_cpk(row['Mean'], row['StdDev'], row['LSL'], row['USL']), axis=1)

# 8. Save the result to a new CSV file
df.to_csv('cpk_output.csv', index=False, encoding='utf-8')
print("Saved 'cpk_output.csv' with 32 random values and Cpk ≥ 1.33 per row.")
