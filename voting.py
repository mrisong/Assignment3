#Task1
import copy
import  openpyxl as op

def generatePreferences(valuationSheet) -> dict:

    '''
    Input: a worksheet containing a set of numerical values assigned by the agents for the different alternatives
    Output: the preference profile according to the valuation in the worksheet
    '''
    
    preferenceProfile = {}
    agent = 1

    for row in valuationSheet.iter_rows(min_row = 1, values_only=True):
        
        # enumerate each row of the input worksheet; enumeration will preserve the index of the alternatives
        alternatives = list(enumerate(row, start=1))

        # reverse the order of the alternatives; It is done because of the requirement that in case of similar valuation for different 
        # alternatives, the alternative with larger indices are considered to be more preferred by the agent.
        # reversing the list make descending order of indices as the defaut order
        alternatives.reverse()
        
        # sort the list of valuations for each agent, in case of equal valuations, the higher indices value is preferred
        # i refers to each alternative
        # i[0] referes to the index value of each alternative; the index value is assigned in the enumerate function above
        preference = [i[0] for i in sorted(alternatives, reverse=True, key=lambda x:x[1])]
        
        # add new entry in the dictionary, where the keys correspond to the agent number and the values correspond to the ordered list 
        # of preferences of alternatives
        preferenceProfile[agent] = preference

        # before next loop, update the agent number
        agent += 1
    
    #output the dictionary of preference profile
    return preferenceProfile

def dictatorship(preferenceProfile, agent) -> int:

    '''
    Input: the preference profile as per the provided valuations; and, a selected agent number
    Output: the most prefered alternative of the selected agent

    Error handling: the function has to project an error if the input number does not correspond to any agent number
        The function will check if the input number is among the available agent numbers
        If the input number is not in the available range of numbers, raise an error
        If no error is raised, return the required output

    '''

    try:
        return preferenceProfile[agent][0]
    except KeyError:
        print("Please enter a valid agent number!!")
        return "Invalid agent number"


def tieBreak(option, alternativesList, preferences = None) -> int:
    
    if preferences != None:
        
        try:
                
            selectedAlternativePreference = preferences[option]

            for alternative in selectedAlternativePreference:
                if alternative in alternativesList:
                    return alternative
        except KeyError:

            print("Please enter a valid agent number!!")
            return False


        
    else:
        if option == 'max':
            return max(alternativesList)
        elif option == 'min':
            return min(alternativesList)


def scoringRule(preferenceProfile, scoreVector, tieBreakOption) -> int:

    '''
    Input: the preference profile as per the provided valuations; a scoring rule to be assigned to according to the preference
            and, the prefered tie breaking strategy
    Output: the alternative with the highest total score; In case of same highest score, choose the alternate according to 
            the tie brake rule
    
    The highest score in the scoring vector is assigned to the most preferred alternative, the second highest to the second most 
    preferred alternative and so on.

    First step is to sort the score vector in decreasing order
    
    Various ways to achieve the next step -
        make a list which store the score total of the alternatives
        the index [0] saves total score for 1st alternative
        index[1] saves total score for 2nd alternative and so on

        for every agent assign the preferences a score and add that score to corresponding score in the total score list
    -----------------
        instead of list make a dictionary
        keys represent the alternatives
        the corresponding values represent the total score
    -----------------
    Identifying the number of elements with the maximum value in final scores of the alternatives

    '''
    preferences = list(preferenceProfile.values())

    noOfAlternatives = len(preferences[0])
    
    try:

        if len(scoreVector) != noOfAlternatives:
            raise Exception("Number of elements in score vector different from number of alternatives")
    except:
        print("Incorrect input")
        return False

    scoreVector.sort(reverse = True)

    totalScore = [0 for _ in range(noOfAlternatives)]

    for preferenceList in preferences:
        i = 0

        for alternative in preferenceList:
            totalScore[alternative-1] += scoreVector[i] # Because the list index start with aero, while alternative numbers start with 1
            i += 1
    
    # Finding the indices corresponding to the maximum score
    mValue = max(totalScore)
    #find all the alternatives
    winningAlternatives = []
    for i in range(len(totalScore)):
        
        if mValue == totalScore[i]:
            
            winningAlternatives.append(i + 1)
        
    if len(winningAlternatives) == 1:
        return winningAlternatives[0]
    else:
        if isinstance(tieBreakOption, int):
            return tieBreak(tieBreakOption, winningAlternatives, preferenceProfile)
        else:
            return tieBreak(tieBreakOption, winningAlternatives)



def plurality(preferenceProfile, tieBreakOption) -> int:
    '''
    Input: the preference profile as per the provided valuations
            and, the prefered tie breaking strategy
    Output: the alternative which appears most time in the first position of the agents' preference ordering
            In case of more than one alternative appears most time, the tie brake rule applies
    
    A frequency list is constructed; the elements correspond to the frequency of alternatives, the list indices correspond to 
    the alternative number (minus 1, because python list indices start from 0)
    '''
    
    preferences = list(preferenceProfile.values())

    noOfAlternatives = len(preferences[0])
    alternativeFrequency = [0 for _ in range(noOfAlternatives)]
    
    for i in list(preferenceProfile.keys()):
        preferredAlternative = preferenceProfile[i][0]
        alternativeFrequency[preferredAlternative - 1] += 1 # '-1' in index is done to account for list indices starting with 0
 
    maxFrequency = max(alternativeFrequency)
    #find all the alternatives
    winningAlternatives = []

    for i in range(len(alternativeFrequency)):
        
        if maxFrequency == alternativeFrequency[i]:
            
            winningAlternatives.append(i + 1)
 
    if len(winningAlternatives) == 1:
        return winningAlternatives[0] #index [0] is added to ensure that an integer type is returned and NOT a list type
    else:
        if isinstance(tieBreakOption, int):
            return tieBreak(tieBreakOption, winningAlternatives, preferenceProfile)
        else:
            return tieBreak(tieBreakOption, winningAlternatives)


def veto(preferenceProfile, tieBreakOption) -> int:
    '''
    Input: the preference profile as per the provided valuations
            and, the prefered tie breaking strategy
    Output: 
            In case of more than one alternative appears most time, the tie brake rule applies
    
    For every agent:
    0 point is assigned to the least prefered alternative
    1 point to every other alternative

    '''

    noOfAlternatives = len(list(preferenceProfile.values())[0])
    alternativePoints = [1 for _ in range(noOfAlternatives)]
    alternativePoints[-1] = 0
    return scoringRule(preferenceProfile, alternativePoints, tieBreakOption)


def borda(preferenceProfile, tieBreakOption) -> int:
    '''
    Input: the preference profile as per the provided valuations
            and, the prefered tie breaking strategy
    Output: 
            In case of more than one alternative appears most time, the tie brake rule applies
    
    For every agent:
        Assign a score of:
            0 to the least preferred alternative
            1 to the second least preferred alternative
            2 to the third least preferred
            and so on.....

    '''

    preferences = list(preferenceProfile.values())

    noOfAlternatives = len(preferences[0])
    alternativePoints = [i for i in range(noOfAlternatives)]

    return scoringRule(preferenceProfile, alternativePoints, tieBreakOption)


def harmonic(preferenceProfile, tieBreakOption) -> int:
    '''
    Input: the preference profile as per the provided valuations
            and, the prefered tie breaking strategy
    Output: 
            In case of more than one alternative appears most time, the tie brake rule applies
    
    For every agent:
        Assign a score of:
            1/m to the least preferred alternative
            1/(m-1) to the second least preferred alternative
            and so on.....
            such that:
            1/j is assigned to the jth favourable alternative
            
            So,
            1 should be assigned to the favourite alternative
            
    '''
    preferences = list(preferenceProfile.values())

    noOfAlternatives = len(preferences[0])
    alternativeScores = [1/(i+1) for i in range(noOfAlternatives)]
    
    return scoringRule(preferenceProfile, alternativeScores, tieBreakOption)


def STV(preferenceProfile, tieBreakOption) -> int:
    '''
    Input: the preference profile as per the provided valuations
            and, the prefered tie breaking strategy
    Output: 
            In case of more than one alternative appears most time, the tie brake rule applies
    
    The winner is selected in rounds:
        In each round, the alternative appearing least frequently in first position of agents is removed
        A frequency list shall be made for first position of agents' alternatives
        The alternative corresponding to least frequency is identified 
        There may be more than one alternative for the least frequency
        All such alternative shall be identified and shall be deleted from the list
        
        These deleted alternatives shall be stored in another list
        The alternatives deleted last shall participate in tie breaker, which shall resolve a winner alternative
        If only one alternative is deleted last, it is the winner alternative

    '''
    voteCountPreferenceProfile = copy.deepcopy(preferenceProfile)

    while voteCountPreferenceProfile[1]:        
        
        alternativeFrequency = {k:0 for k in list(voteCountPreferenceProfile.values())[0]}

        for i in list(voteCountPreferenceProfile.keys()):
            preferredAlternative = voteCountPreferenceProfile[i][0]
            alternativeFrequency[preferredAlternative] += 1
        
        #print(f"voteCountPreferenceProfile {voteCountPreferenceProfile}")
        #print(f"preferenceProfile {preferenceProfile}")
        #print(f"alternativeFrequency {alternativeFrequency}")

        minimumFrequency = min(alternativeFrequency.values())
        leastFrequentAlternatives = []

        for i in alternativeFrequency.keys():
            if alternativeFrequency[i] == minimumFrequency:
                leastFrequentAlternatives.append(i)

        for leastFrequent in leastFrequentAlternatives:
            for i in list(voteCountPreferenceProfile.keys()):
#                print(f"the preferences for agent {i} is {voteCountPreferenceProfile[i]}, least frequent alternatives are {leastFrequentAlternatives}, current alternative is {leastFrequent}")
                voteCountPreferenceProfile[i].remove(leastFrequent)


    if len(leastFrequentAlternatives) == 1:
        return leastFrequentAlternatives[0] #index [0] is added to ensure that an integer type is returned and NOT a list type
    else:
        if isinstance(tieBreakOption, int):
            #print(f"Tie Break option is indeed an int and competition is in betweeen {leastFrequentAlternatives} and the preference profile is {preferenceProfile}")
            return tieBreak(tieBreakOption, leastFrequentAlternatives, preferenceProfile)
        else:
            return tieBreak(tieBreakOption, leastFrequentAlternatives)


def rangeVoting(valuationSheet, tieBreakOption) -> int:
    '''
    Input: worksheet containing valuation
            and, the prefered tie breaking strategy
    Output: 
            In case of more than one alternative appears most time, the tie brake rule applies
    
    The winner is the alternative with maximum sum of valuations
        
    '''
    preferenceProfile = generatePreferences(valuationSheet)

    alternativeValueSum = []

    for valuation in valuationSheet.iter_cols(min_col=1, values_only=True):
        alternativeValueSum.append(sum(valuation))

    maxPoints = max(alternativeValueSum)
    #find all the alternatives
    winningAlternatives = []

    for i in range(len(alternativeValueSum)):
        
        if maxPoints == alternativeValueSum[i]:           
            winningAlternatives.append(i + 1)
    

    if len(winningAlternatives) == 1:
        return winningAlternatives[0] #index [0] is added to ensure that an integer type is returned and NOT a list type
    else:
        if isinstance(tieBreakOption, int):
            return tieBreak(tieBreakOption, winningAlternatives, preferenceProfile)
        else:
            return tieBreak(tieBreakOption, winningAlternatives)
