import pandas as pd

df = pd.read_excel('test.xlsx')

result = df.groupby(['Job Title', 'Question', 'Correct Answer']).size().unstack(fill_value=0)
result['% Correct'] = result['Correct Answer'] / (result['Correct Answer'] + result['Wrong Answer']) * 100
result['% Wrong'] = 100 - result['% Correct']

final_result = result[["% Correct", "% Wrong"]].reset_index()

final_result.to_excel('output.xlsx', index=False)

print(final_result)