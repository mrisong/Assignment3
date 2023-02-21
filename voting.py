
import  openpyxl as op
import copy




def generatePreferences(valuationSheet) -> dict:

    '''
    Input: a worksheet containing a set of numerical values assigned by the agents for the different alternatives

    Output: the preference profile according to the valuation in the worksheet

    The steps followed are:
    Associate an alternative number for each column starting from 1 and adding 1 for subsequent columns
    Associate an agent number for each row starting from 1 and adding 1 for subsequent rows 
    '''
    
    preferenceProfile = {}
    agent = 1

    for row in valuationSheet.iter_rows(min_row = 1, values_only=True):
        
        # enumerate each row of the input worksheet to associate a number with each element, corresponding to the alternative number
        alternatives = list(enumerate(row, start=1))

        # reverse the order of the alternatives; It is done because of the requirement that in case of similar valuation for different 
        # alternatives, the alternative with larger indices are considered to be more preferred by the agent.
        # reversing the list results in the list with descending order of alternative numbers
        alternatives.reverse()
        
        # sort the list in decreasing order of valuations for each agent
        # In case of equal valuations, the higher indices value appears first
        # i refers to each enumerated alternative
        # i[0] referes to the index value of each alternative; the index value is assigned in the enumerate function above
        preference = [i[0] for i in sorted(alternatives, reverse=True, key=lambda x:x[1])]
        
        # add new entry in the dictionary, where the keys correspond to the agent number and the values correspond to the ordered list 
        # of its preferences of alternatives
        preferenceProfile[agent] = preference

        # before next loop, update the agent number
        agent += 1
    
    #return the dictionary of preference profile
    return preferenceProfile




def dictatorship(preferenceProfile, agent) -> int:

    '''
    Input: A dictionary containing preference profile; and, a selected agent number

    Output: the most prefered alternative of the selected agent

    Error handling:
        Error shall be raised if input is an integer but does not correspond to an agent.
        Error handling is not included for the cases when the input is not an integer.
    '''

    if isinstance(agent, int):

        #Implement Error handling only if the input for agent number is an integer
        try:

            return preferenceProfile[agent][0]

        except KeyError:
            
            print("Please enter a valid agent number!!")
            return "Invalid agent number"

    else:

        #If the input is not an integer, the return will terminate the program with an error
        return preferenceProfile[agent][0]




def tieBreak(option, bestAlternatives, preferenceProfile = None) -> int:

    '''
    Input: the preferred option for tie break;
            a list of the best alternatives among which a winner shall be choosen according to the tie break option;
            an optional input of the dictionary containg preference profile

    Output: the winner alternative among the list of the best alternatives

    Error handling:
        Error shall be raised if input is an integer but does not correspond to an agent.
        Error handling is not included for the cases when the input is not an integer.
    '''
    
    if preferenceProfile != None: 
        # Enter this section only if the input is a preference profile dictionary

        # Implement Error handling for the cases when the input integer does not correspond to any agent
        try:

            selectedAlternativePreference = preferenceProfile[option] # This statement will cause a 'KeyError' if the input is not an integer, resulting in execution of 'except' clause

            for alternative in selectedAlternativePreference:
                
                #Check the alternatives in the preference of the input agent number
                if alternative in bestAlternatives:
                    # The first alternative in the preference order which appear in the list of best alternative, shall be the winning alternative
                    return alternative
                
        except KeyError:

            print("Please enter a valid agent number!!")
            return False
        
    else:

        if option == 'max':
            return max(bestAlternatives)
        
        elif option == 'min':
            return min(bestAlternatives)
        
        else:
            return "Invalid Tie Break Option!! \nPlease choose 'max' OR 'min' OR an agent number as the tie break option... "




def scoringRule(preferenceProfile, scoreVector, tieBreakOption) -> int:

    '''
    Input: A dictionary containing preference profile;
            A scoring rule to be assigned to the alternatives according to the preference order;
            A tie breaking option

    Output: The alternative with the highest total score;
            In case of same highest score for multiple alternatives, choose the winner according to the tie braking option
    
    For each agent, the highest score in the scoring vector is assigned to the most preferred alternative, the second highest to the second most preferred alternative 
    and so on.
    Total score of each alternative is calculated by adding scores assigned for each agent's preferences.

    Error Handling is implemented for the cases when the length of score vector list is not equal to the number of alternatives
    '''
    
    #Find the number of alternatives
    preferences = list(preferenceProfile.values()) #This step makes the program more robust; If later it is decided to introduce name of agents instead of integers, the scoring rule shall still run without modification
    noOfAlternatives = len(preferences[0]) # Using the length of preference list of first agent. Any agent may be used since all vote for the same number of alternatives
    

    # Raise an error if the length of the input score vector is not equal to the number of alternatives
    try:

        if len(scoreVector) != noOfAlternatives:
            raise Exception("Number of elements in score vector different from number of alternatives")
        
    except:

        print("Incorrect input")
        return False

    # Sort the score vector in decreasing order, because the maximum score shall be assigned to the more prefered alternative and so on for each agent
    scoreVector.sort(reverse = True)

    # Initialise the total score vector
    totalScore = [0 for _ in range(noOfAlternatives)]

    # Add the scores assigned to an alternative and repeat for all the agents
    for preferenceList in preferences:

        i = 0
        for alternative in preferenceList:

            totalScore[alternative-1] += scoreVector[i] # 'alternative - 1' because the list index start with zero, while alternative numbers start with 1
            i += 1
    

    # Initialise the new list, which shall store the alternative number of all the winning alternatives
    winningAlternatives = []

    # Find the maximum score
    maxScore = max(totalScore)

    # Find all the alternatives whose score is equal to the maximum score
    for i in range(len(totalScore)):
        
        if maxScore == totalScore[i]:           
            winningAlternatives.append(i + 1)

    if len(winningAlternatives) == 1:
        #If there is only one alternative with maximum score, return the winner alternative
        return winningAlternatives[0]
    
    else:

        #If the tie break option is an integer, pass the preference profile dictionary to the tie break function
        if isinstance(tieBreakOption, int):
            return tieBreak(tieBreakOption, winningAlternatives, preferenceProfile)
        
        #Tie break option must be either 'min' OR 'max' to return a winner alternative, otherwise this will return a warning from the tie break function
        else:
            return tieBreak(tieBreakOption, winningAlternatives)




def plurality(preferenceProfile, tieBreakOption) -> int:

    '''

    Input: A dictionary containing preference profile;
            A tie breaking option

    Output: The alternative which appears most time in the first position of the agents' preference ordering
            In case more than one alternative is most preferred most frequently, choose the winner according to the tie braking option
    
    A frequency list is constructed; the elements correspond to the frequency of alternatives, the list indices correspond to 
    the alternative number minus 1 (because python list indices start from 0)
    '''
    
    # Find the number of alternatives
    preferences = list(preferenceProfile.values()) #This step makes the program more robust; If later it is decided to introduce name of agents instead of integers, the scoring rule shall still run without modification
    noOfAlternatives = len(preferences[0]) # Using the length of preference list of first agent. Any agent may be used since all vote for the same number of alternatives

    # Initialise a list to store frequency of the most preferred alternative
    alternativeFrequency = [0 for _ in range(noOfAlternatives)]
    
    # Find the frequency of most preffered alternative; loop through all the agents in the preference profile dictionary
    for preference in preferences:

        preferredAlternative = preference[0] # First alternative for each preference
        alternativeFrequency[preferredAlternative - 1] += 1 # '-1' in index is done to account for list indices starting with 0
    
    # Initialise a new list, which shall store the alternative number of all the winning alternatives
    winningAlternatives = []

    # Find the maximum frequency with which any alternative is most preffered
    maxFrequency = max(alternativeFrequency)

    # Find all the alternatives whose frequency is equal to the maximum frequency
    for i in range(len(alternativeFrequency)):
        
        if maxFrequency == alternativeFrequency[i]:
            
            winningAlternatives.append(i + 1)
 
    if len(winningAlternatives) == 1:
        #If there is only one alternative with maximum score, return the winner alternative
        return winningAlternatives[0] #index [0] is added to ensure that an integer type is returned and NOT a list type
    
    else:

        #If the tie break option is an integer, pass the preference profile dictionary to the tie break function
        if isinstance(tieBreakOption, int):
            return tieBreak(tieBreakOption, winningAlternatives, preferenceProfile)
        
        #Tie break option must be either 'min' OR 'max' to return a winner alternative, otherwise this will return a warning from the tie break function
        else:
            return tieBreak(tieBreakOption, winningAlternatives)




def veto(preferenceProfile, tieBreakOption) -> int:

    '''
    Input: A dictionary containing preference profile;
            A tie breaking option

    Output: The alternative with highest total score, where the scores are assigned according to the veto rule;
            In case of same highest score for multiple alternatives, choose the winner according to the tie braking option
    
    Scoring rule:
    For every agent-
    0 point is assigned to the least prefered alternative
    1 point to every other alternative

    '''

    # Find the number of alternatives
    preferences = list(preferenceProfile.values()) #This step makes the program more robust; If later it is decided to introduce name of agents instead of integers, the scoring rule shall still run without modification
    noOfAlternatives = len(preferences[0]) # Using the length of preference list of first agent. Any agent may be used since all vote for the same number of alternatives

    # Create a scoring rule; initialise a list with length equal to number of alternatives; The list contain all 1's except last element which is 0
    alternativePoints = [1 for _ in range(noOfAlternatives)]
    alternativePoints[-1] = 0
    
    # Call the scoring rule function and return its value; Pass the preference profile dictionary, scoring rule defined above and the tie break option
    return scoringRule(preferenceProfile, alternativePoints, tieBreakOption)




def borda(preferenceProfile, tieBreakOption) -> int:

    '''
 
    Input: A dictionary containing preference profile;
            A tie breaking option

    Output: The alternative with highest total score, where the scores are assigned according to the borda rule;
            In case of same highest score for multiple alternatives, choose the winner according to the tie braking option
    
    Scoring rule:
    For every agent, assign a score of -
        0 to the least preferred alternative
        1 to the second least preferred alternative
        2 to the third least preferred
        and so on.....

    '''

    # Find the number of alternatives
    preferences = list(preferenceProfile.values()) #This step makes the program more robust; If later it is decided to introduce name of agents instead of integers, the scoring rule shall still run without modification
    noOfAlternatives = len(preferences[0]) # Using the length of preference list of first agent. Any agent may be used since all vote for the same number of alternatives

    # Create a scoring rule; initialise a list with each element equal to its index
    # Order of the list is not important here because, for assigning score, the list will be ordered in the scoring rule function
    alternativePoints = [i for i in range(noOfAlternatives)]

    # Call the scoring rule function and return its value; Pass the preference profile dictionary, scoring rule defined above and the tie break option
    return scoringRule(preferenceProfile, alternativePoints, tieBreakOption)




def harmonic(preferenceProfile, tieBreakOption) -> int:

    '''
    Input: A dictionary containing preference profile;
            A tie breaking option

    Output: The alternative with highest total score, where the scores are assigned according to the harmonic rule;
            In case of same highest score for multiple alternatives, choose the winner according to the tie braking option
    
    Scoring rule:
    For every agent, assign a score of -
            1/m to the least preferred alternative
            1/(m-1) to the second least preferred alternative
            and so on.....
            such that:
            1/j is assigned to the jth favourable alternative
            
            So,
            1 should be assigned to the favourite alternative
            
    '''
    # Find the number of alternatives
    preferences = list(preferenceProfile.values()) #This step makes the program more robust; If later it is decided to introduce name of agents instead of integers, the scoring rule shall still run without modification
    noOfAlternatives = len(preferences[0]) # Using the length of preference list of first agent. Any agent may be used since all vote for the same number of alternatives

    # Create a scoring rule
    # Order of the list is not important here because, for assigning score, the list will be ordered in the scoring rule function
    alternativeScores = [1/(i+1) for i in range(noOfAlternatives)]
    
    # Call the scoring rule function and return its value; Pass the preference profile dictionary, scoring rule defined above and the tie break option
    return scoringRule(preferenceProfile, alternativeScores, tieBreakOption)




def STV(preferenceProfile, tieBreakOption) -> int:

    '''
    Input: A dictionary containing preference profile;
            A tie breaking option

    Output: 
            The alternative which satisfies the STV voting rule;
            In case of multiple such alternatives, choose the winner according to the tie braking option
    
    The winner is selected in rounds:
        In each round, the alternative appearing least frequently in first position of agents is removed
        A frequency list shall be made for first position of agents' alternatives
        The alternative corresponding to least frequency is identified 
        There may be more than one alternative for the least frequency
        All such alternative shall be identified and shall be deleted for subsequent operations
        
        These deleted alternatives shall be stored in another list
        The alternatives deleted last shall participate in tie breaker, which shall resolve a winner alternative
        If only one alternative is deleted last, it is the winner alternative

    '''

    # Copy the Preference Profile dictionary, the copied dictionary shall be modified by deleting the least frequent most preffered alternative
    #  because it need to be passed for tie break
    updatedProfile = copy.deepcopy(preferenceProfile) # All the methods to copy other than deepcopy, resulted in a shallow copy

    agents = list(updatedProfile.keys()) # Used later in the function to loop through all the agents

    # The following loop removes the least frequent most preffered alternative from the updated preference profile
    while updatedProfile[agents[0]]: # Run the loop untill the list of preferences does not empty
                                     # agent[0] is used to ensure robustness of the program- In case agents are identified by something
                                     # other than integers, this function shall still work
        
        # Initialise a dictionary which contains alternatives as keys and values as their frequency as most preffered option
        # The keys are all the alternatives currently available in the updated preference profile
        alternativeFrequency = {k:0 for k in list(updatedProfile.values())[0]}

        # Count frequency of all the most preferred alternatives
        for agent in agents:
            preferredAlternative = updatedProfile[agent][0] # index 0 referes to the most prefered alternative
            alternativeFrequency[preferredAlternative] += 1 # increase the frequency of the most prefered alternative by 1
        
        # Find the minimum frequency of all the available alternative as most preferred option by agents
        minimumFrequency = min(alternativeFrequency.values())

        # Initialise a list to store all the alternatives which appear with the minimum frequency
        leastFrequentAlternatives = []
        
        # Find all the alternatives which appear least frequently as the most preffered option
        for i in alternativeFrequency.keys():
            if alternativeFrequency[i] == minimumFrequency:
                leastFrequentAlternatives.append(i)
        
        # Update the preference profile by removing all the least frequent most preffered alternative
        for leastFrequent in leastFrequentAlternatives:

            #Loop through all the agents
            for agent in agents:
                updatedProfile[agent].remove(leastFrequent)
    

    if len(leastFrequentAlternatives) == 1:

        #If there is only one alternative with maximum score, return the winner alternative
        return leastFrequentAlternatives[0] #index [0] is added to ensure that an integer type is returned and NOT a list type
    
    else:

        #If the tie break option is an integer, pass the preference profile dictionary to the tie break function
        if isinstance(tieBreakOption, int):
            return tieBreak(tieBreakOption, leastFrequentAlternatives, preferenceProfile)
        
        #Tie break option must be either 'min' OR 'max' to return a winner alternative, otherwise this will return a warning from the tie break function
        else:
            return tieBreak(tieBreakOption, leastFrequentAlternatives)




def rangeVoting(valuationSheet, tieBreakOption) -> int:

    '''
    Input: A worksheet (xlsx file) containing valuation for the alternatives by the agents;
            A tie breaking option
    Output: 
            The alternative which satisfies the range voting rule;
            In case of multiple such alternatives, choose the winner according to the tie braking option
    
    The winner is the alternative with maximum sum of valuations
    '''

    # Initialise the list to store the sum of valuations for each alternative
    alternativeValueSum = []

    # Calculate the sum of valuations for each alternative assigned by each agent
    for valuation in valuationSheet.iter_cols(min_col=1, values_only=True):
        alternativeValueSum.append(sum(valuation))
    
    # Find the maximum sum
    maxPoints = max(alternativeValueSum)

    # Initialise a list to store all the alternatives with maximum total valuation
    winningAlternatives = []

    # Add the alternatives with maximum valuation in the new list
    for i in range(len(alternativeValueSum)):
        
        if maxPoints == alternativeValueSum[i]:           
            winningAlternatives.append(i + 1)
    

    if len(winningAlternatives) == 1:

        #If there is only one alternative with maximum score, return the winner alternative
        return winningAlternatives[0] #index [0] is added to ensure that an integer type is returned and NOT a list type
    else:

        #If the tie break option is an integer, pass the preference profile dictionary to the tie break function
        if isinstance(tieBreakOption, int):

            # A new dictionary is generated for preference profile because tie break function need dictionary as an input
            preferenceProfile = generatePreferences(valuationSheet)
            return tieBreak(tieBreakOption, winningAlternatives, preferenceProfile)
        
        else:
            #Tie break option must be either 'min' OR 'max' to return a winner alternative, otherwise this will return a warning from the tie break function
            return tieBreak(tieBreakOption, winningAlternatives)