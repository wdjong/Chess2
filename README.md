# Chess2
This is currently in VB.NET (VS2019) and tested on Windows 10.

Past
This project started in 2000 I had a separate project that created a DLL for the basic chess piece functionality. There were a bunch of chess piece classes deriving from a generic piece class. This was in order to develop my skills in .Net, object oriented programming and a bit of chess practice. My son was playing a bit of chess and I thought it would be fun to think about an automated opponent. 

Basic play
From the File menu, select new. You are red (=white);. 
Pieces are moved in accordance with standard chess rules. (It won't correct illegal moves)
To move, drag a piece onto its destination square. To castle, right-click on a rook.
If a piece is taken, it sits on the bar and must be brought back into play before other pieces can be moved.
A pawn reaching the 8th rank is automatically promoted to a queen.
Computer opponent
The computer opponent uses a basic method of scoring each possible move and then evaluating the resulting board situation weighting various criteria by a number of factor. These factors relate to strategy elements or qualities that more or less equate to temperaments. E.g. the tendency to want to threaten pieces is called 'aggression'. I haven't done much yet in terms of optimising these. I don't do any thinking ahead (i.e. if I do this he'll probably do that) (0 ply) 
Threatening 1 'aggressiveness - increase our threatening 
AvoidingThreat 1 'wimpishness - reduce their threatening 
Supporting 1 'supporting our own pieces - increase our support
Undermine 1 'reduce their support 
Control 4 'aim to control more of the board 
Frustrate 1 'aim to reduce their mobility and control 
Keep 2 'aim to keep as many pieces as possible 
Erode 1 'aim to reduce their pieces 
Consolidate 5 'aim to avoid leaving piece unsupported 
Isolate 1 'aim to isolate their pieces leaving them unsupported 
In order to allow the computer to 'fine tune' its play to suite playing against you, there is a capacity to weight various characteristics, and prefer to use more successful strategies in response to success. The results of this can be seen in the Tools Settings screen. 
Commands
File menu
New
Start a new game
Open
Exit
Close the program.
Edit menu
Undo
Undo as many moves as you like.
Tools
Move
Let the computer make the next move
Top Human
toggle between blue (black) being a human or computer
Bottom Human
toggle between red (white) being a human or computer
Help
Manual
Display user manual.
Other

Future
It doesn't play very smartly at the moment. I'll probbly start by optimising weightings, implementing learning as per BGC. Implement > 0 ply. I'm interested in Reinforcement learning (Strategy) and learning more about deep Learning (Neural networks) a
