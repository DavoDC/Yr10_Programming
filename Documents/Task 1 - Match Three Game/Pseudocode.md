# Pseudocode

*Converted from Pseudocode.docx*

---

Pseudocode for Command Button

Subtract 1 from Current Value of Turns
Set Points Earned to 0

Randomize

Image Box 1
Choose random integer from 1 to 11
Select one of these cases randomly

Case 1
Image Box 1= Load Picture 1

Case 2
Image Box 1= Load Picture 2

Case 3
Image Box 1= Load Picture 3

Case 4
Image Box 1= Load Picture 4

Case 5
Image Box 1= Load Picture 5

Case 6
Image Box 1= Load Picture 6

Case 7
Image Box 1= Load Picture 7

Case 8
Image Box 1= Load Picture 8

Case 9
Image Box 1= Load Picture 9

Case 10
Image Box 1= Load Picture 10

Case 11
Image Box 1= Load Picture 11



Image Box 2
Choose random integer from 1 to 11
Select one of these cases randomly

Case 1
Image Box 2= Load Picture 1

Case 2
Image Box 2= Load Picture 2

Case 3
Image Box 2= Load Picture 3

Case 4
Image Box 2= Load Picture 4

Case 5
Image Box 2= Load Picture 5

Case 6
Image Box 2= Load Picture 6

Case 7
Image Box 2= Load Picture 7

Case 8
Image Box 2= Load Picture 8

Case 9
Image Box 2= Load Picture 9

Case 10
Image Box 2= Load Picture 10

Case 11
Image Box 2= Load Picture 11


Image Box 3
Choose random integer from 1 to 11
Select one of these cases randomly

Case 1
Image Box 3= Load Picture 1

Case 2
Image Box 3= Load Picture 2

Case 3
Image Box 3= Load Picture 3

Case 4
Image Box 3= Load Picture 4

Case 5
Image Box 3= Load Picture 5

Case 6
Image Box 3= Load Picture 6

Case 7
Image Box 3= Load Picture 7

Case 8
Image Box 3= Load Picture 8

Case 9
Image Box 3= Load Picture 9

Case 10
Image Box 3= Load Picture 10

Case 11
Image Box 3= Load Picture 11

If Turns = 0 Then
Display Message = you have used up all 20 turns!   If you have achieved a High Score for this round it will appear in the space below after closing this message

If ImageBox1 = ImageBox2 and ImageBox 2 = ImageBox 3 Then
Display Message" Well Done! You matched three!"
Add 100 to Score
Add 100 to Points Earned This Turn

If (ImageBox1 = ImageBox2) or (ImageBox 2 = ImageBox3) or (ImageBox1 = ImageBox3) Then
Add 20 to Score
Add 20 to Points Earned This Turn

If (ImageBox1 has Picture 11) or (ImageBox2 has Picture 11) or (ImageBox3 has Picture 11) Then
Add 5 to Score
Add 5 to Points earned this Turn

If (ImageBox1 = 11) and (ImageBox2 = 11) and (ImageBox3 = 11) Then
Display Message “Well Done! You matched three!"
Add 115 to Score
Add 115 to Points earned this Turn

If (ImageBox1 = 11) and (ImageBox2 = 11) Then
Add 10 to Score
Add 10 to Points earned this Turn

If (ImageBox2 = 11) and (ImageBox3 = 11) Then
Add 10 to Score
Add 10 to Points earned this Turn

If (ImageBox1 = 11) and (ImageBox3 = 11) Then
Add 10 to Score
Add 10 to Points earned this Turn

If (Turns = 0) and (High Score < Score) Then
Display Message "Well Done! You beat your High Score!"
High Score= Score
Add 20 to Turns
Set Score to 0
Set Points earned this Turn to 0

If (Turns= 0) and (Highscore > Score) Then
Highscore= Highscore
Add 20 to Turns
Set Score to 0
Set Points earned this Turn to 0

If (Score= 0) and (Highscore = Score) Then
Highscore= Highscore
Add 20 to Turns
Set Score to 0
Set Points earned this Turn to 0






Pseudocode for Restart Button (In Menu)
Set Score to 0
Set Highscore to 0
Set Points earned this Turn to 0
Set Turns to 20
ImageBox 1 = Load Picture ‘Blank’
ImageBox 2 = Load Picture ‘Blank’
ImageBox 3 = Load Picture ‘Blank’


Pseudocode for Quit Button (In Menu)
Display Message: "Thankyou for playing Match 3"
End Program/Close Window





Pseudocode for Instructions Button (In Menu)
Display Message: "The objective of the game is to score as many points as possible in 20 turns. You win 20 points when any two frames show the same graphic and 120 points when all three frames show the same graphic ( The Wild Card is an exception to these rules)."


Pseudocode for Wild Card Explanation Button (In Menu)
Display Message: "You win 5 points when any frame displays the wild card graphic.   If two frames show the wild card graphic, you will score 20 points for matching two graphics and 10 points for getting 2 wild card graphics. You will also get 5 bonus points for this feat (Total = 20 + (5x2) + 5 = 35).     If all three frames show the wild card graphic, you will score 100 points for matching three graphics and 15 points for getting 3 wild card graphics. You will also gain 155 bonus points by making such a huge achievement (Total = 100 + (5x3) + 155 = 270)."

Pseudocode for About Button (In Menu)
Display Message: “Match 3 Game ©Copyright - All Rights Reserved    -   Created by David Charkey (2015)”