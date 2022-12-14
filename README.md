# VBA-challenge

#<h1>Grace's VBA Challenge Read Me </h1>

##<h2> Explanation of Code</h2>

###<h3> Main Solution </h3>
The first step is to declare all the variables needed. 

###<h3> Looping through each worksheet </h3>
In order to make sure the entire code iterated for each worksheet, I wrapped the entire code with a "For Each...Next" statement. 

In order for this to work, each reference to a particular cell, column, or range needs to be appended with the variable used in the "for each" statement in the beginning. I used "ws" as the variable and made sure every Cell() and Range() had a ws. in front of it. 
