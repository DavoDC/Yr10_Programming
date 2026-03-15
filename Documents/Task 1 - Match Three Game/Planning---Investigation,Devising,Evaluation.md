# Planning - Investigation,Devising,Evaluation

*Converted from Planning - Investigation,Devising,Evaluation.docx*

---

Programming Year 10

Task 1 – Match Three


David Charkey






Contents:
Page 1 : Title Page
Page 2: Investigation
Page 3: Devising
Page 4: Evaluation + User Instructions + Assignment Location









Investigation:
 Define the Problem:
I must create a program in Visual Basic that will simulate a game that matches three graphics (pictures) and allocates a score. It should consist of three frames, each capable of displaying a graphic out of the same set of the ten different graphics. The graphics should be randomly selected.

I don’t like this kind of UI, it is not compact enough and uses too much form room. It also doesn’t give extra information about turns left and points earned per turn.I don’t like this kind of UI, it is not compact enough and uses too much form room. It also doesn’t give extra information about turns left and points earned per turn. Compare user interfaces that you like and don’t like and explain why:
I don’t like this kind of UI, it is not compact enough and uses too much form room. It also doesn’t give extra information about turns left and points earned per turn.
I don’t like this kind of UI, it is not compact enough and uses too much form room. It also doesn’t give extra information about turns left and points earned per turn.
This kind of UI is exceptional. It makes very good use of Visual Basic features and form space. It also is very informative, giving details regarding score, high score, turns and points earned in a turn. The message boxes save a great amount of form space, making the program much neater.This kind of UI is exceptional. It makes very good use of Visual Basic features and form space. It also is very informative, giving details regarding score, high score, turns and points earned in a turn. The message boxes save a great amount of form space, making the program much neater.
This kind of UI is exceptional. It makes very good use of Visual Basic features and form space. It also is very informative, giving details regarding score, high score, turns and points earned in a turn. The message boxes save a great amount of form space, making the program much neater.
This kind of UI is exceptional. It makes very good use of Visual Basic features and form space. It also is very informative, giving details regarding score, high score, turns and points earned in a turn. The message boxes save a great amount of form space, making the program much neater.

IPO Diagram
Input
Process
Output
Press Command Button
Subtract 1 from Turns
If all 3 images are the same = Add 100 to score.
If 2 of the images are the same = Add 20 to score.
If a wildcard appears =
Add 5 to score.
If 2 wildcards appear =
Add 35 to score.
If 3 wildcards appear =
Add 270 to score.
Display Score
Display High Score
Display Turns
Displays Points earned that Turn



Devising





 Show the properties you will change for each object.
For this information , please check the folder called ‘Other Documents ‘ and you will a Word Document named ‘Objects’

For Font , Form and Object colours refer to diagram above



For Pseudocode, find a folder in the same folder in which this document is located  called ‘Other Documents  ‘ and inside it you will find a Word document named ‘Pseudocode’

Algorithm =
Get Random Seed











Evaluation
 Justify changes of final product compared to original design
The original design did not include the Message Box notifications that appear when certain events happen in the game. (For example, when you get 3 images that are the same, you receive a message stating: ‘Well done! You matched three!) . I added this feature while creating the program.
The original design did not include the information regarding ‘Points earned this Turn’. I thought of this idea while testing out the game and I added it soon after.
The original design did not include a wild card system. I added the wild card feature as a way for the player to get a higher score easier.

 Discuss any difficulties and how you solved the problem and what improvements would you make in your next version.
I had some difficulty getting the pictures to load into the image boxes. After checking the code I realised that I had entered the wrong location into the code. I substituted the correct location into the code afterwards, fixing the problem.
I had some trouble with the turns system. It was difficult to stop the turns value from going into negative numbers. I solved this problem by making the first line of code of the command button to set the value of the turns to 0. I also added code to the scoring system which pushed the value of the turns to 20 after it reached 0.
I am not sure what improvements I would make in my next version of this program as all the improvements that I could think of were added to my program.

 How well would you evaluate your program?
I am very confident in my program/game and I believe I created it to the best of my ability.
All aspects of it work very well, for example the scoring and message box features.
I think my program is user friendly and visually appealing.


User Instructions
For Internal Documentation, find a folder in the same folder in which this document is located  called ‘Other Documents  ‘ and inside it you will find a Word document named ‘Internal Documentation’

For External Documentation, find a folder in the same folder in which this document is located  called ‘Other Documents  ‘ and inside it you will find a Word document named ‘External Documentation’

If you would like to know how to play the game, check the in-game instructions.

Assignment Location
Z:\Programming Year 10\Assignments\Task 1 - Match Three Game\Match Three\

By David Charkey