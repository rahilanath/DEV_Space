import pandas as pd

test1 = [[1,2,3,4],[2,3,4,5],[3,4,5,6],[11,11,12,13]]
test2 = [[9,8,7,6],[1,2,3,4],[10,11,12,13]]

df1 = pd.DataFrame(test1, columns=['first','second','third','fourth'])
df2 = pd.DataFrame(test2, columns=['first','second','third','fourth'])

# print(all(sub in df1.values for sub in df2.values))
# check = [sub for sub in df1.values if sub in df2.values]
# print(check)
# print([sub for sub in df1.values if sub in df2.values[:None]])

print(df1[::1].values)
print(df1[::2].values)
# for sub in df1[::1].values:
#     print(sub)
#     print(sub in df2.values)

# merged = df1.merge(df2, how = 'left', indicator = True)
# filtered = merged[merged['_merge'] == 'left_only']
# filtered.drop('_merge', axis='columns', inplace=True)