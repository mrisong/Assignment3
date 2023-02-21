# Assignment3 - Voting
A part of coursework for 'Programming Fundamentals':

* 'voting.py' is the main file with all the functions corresponding to different voting rules
* 'run_voting.py' shall be used to run 'voting.py', the file loads an excel file and shall be used to call functions corresponding to the voting rule required by the user.
* 'votingTest.xlsx' contains numerical data which corresponds to the evaluation assigned to various alternatives by different agents

The purpose of the program is to choose a winner among multiple alternatives, each voted by several agents.
The user may use any of the available voting method to choose the winner
The voting methods are implemented as different functions in 'voting.py'.
In case of tie between different alternatives, a tie break function is called to declare the winner.
The tie break either decide by alternative number- choose the alternative with highest or lowest index number as per user preference, OR it can also decide based upon a particular agent's preference, i.e. the winner will be the alternative most preffered by the selected agent.

An excel file is needed as input.
The file shall provide the qualitative evaluation assigned to each alternative by each agent.
The 'agents' shall be represented by rows and 'alternatives' by columns.
