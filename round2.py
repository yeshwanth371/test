import pandas as pd

# Read in the users.xlsx file
users = pd.read_excel('users.xlsx')

#read statements.xlsx file
statements = pd.read_excel('statements.xlsx')

#merge based on uid
merged = pd.merge(users, statements, on='uid')

#show total statements and reasons for each team summing up the total statements and reasons by add ing all counts of statements and reasons of each person in the team
merged.groupby('Team Name').agg({'total_statements': 'sum', 'total_reasons': 'sum'})

#calculate the average statements and average reasons for each team
merged.groupby('Team Name').agg({'total_statements': 'mean', 'total_reasons': 'mean'})

#save it in team_stat.xlsx in descending order of total statements and total reasons with rank for each team

merged.groupby('Team Name').agg({'total_statements': 'mean', 'total_reasons': 'mean'}).sort_values(by=['total_statements', 'total_reasons'], ascending=False).to_excel('team_stat.xlsx')


#show total statements and reasons for each individual
merged.groupby('Name').agg({'total_statements': 'sum', 'total_reasons': 'sum'})

#save it in individual_stat.xlsx in descending order of total statements and total reasons with rank
merged.groupby('Name').agg({'total_statements': 'sum', 'total_reasons': 'sum'}).sort_values(by=['total_statements', 'total_reasons'], ascending=False).reset_index().rename_axis(None, axis=1).reset_index().rename(columns={'index': 'Rank'}).to_excel('individual_stat.xlsx')
