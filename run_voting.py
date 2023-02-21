import voting as vT

'''
This file may be used to run 'voting.py' 

This file loads a workbook using openpyxl library, which is already imported in 'voting.py'
The workbook shall contain evaluations provided by agents to different available alternatives
Evaluations shall be numerical values, the rows shall correspond to agents and the columns shall correspond to the alternatives

Various voting rules may be used to chose a winning alternative.
To use a voting rule, the corresponding command may be uncommented

'''

#Load the 'xlsx' file as a workbook
valuationsBook = vT.op.load_workbook('votingTest.xlsx')

# open the worksheet
valuationsSheet = valuationsBook.active

# Call the function to generate a preference profile
preferenceProfile = vT.generatePreferences(valuationsSheet)

#Print the preference profile
print(f'The preference profile generated based on the votes is: {preferenceProfile}')


# Scoring Rule:
scoreVector = [1, 2, 4, 3, 2]
winnerAlternative = vT.scoringRule(preferenceProfile, scoreVector, 1)
print(f"winnerAlternative by Scoring rule is {winnerAlternative}")


#Plurality
print("Winner by plurality", vT.plurality(preferenceProfile, 6))

#Veto
print("Winner by Veto", vT.veto(preferenceProfile, 6))

#Borda
print("Winner by Borda", vT.borda(preferenceProfile, 'min'))

#Harmonic
print("Winner by Harmonic is", vT.harmonic(preferenceProfile, 'max'))

#STV
print("Winner by STV", vT.STV(preferenceProfile, 6))

#Range Voting
print("Winner by Range Voting", vT.rangeVoting(valuationsSheet, 1))
