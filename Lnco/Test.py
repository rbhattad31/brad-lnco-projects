import pandas as pd
data = {'name': ['nick', 'david', 'joe', 'ross'],
        'age': ['5', '10', '7', '6']}
new = pd.DataFrame.from_dict(data)

print(new)