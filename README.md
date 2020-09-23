<div align="center">

## A Fully \(i hope\) string manipulation example / tutorial\.\.\. InStr / Replace / Split / Join / strComp


</div>

### Description

Utilises a lot of the string manipulation commands to do just about anything with a line of text. Some of the more advanced functions i learnt as i went along, so they arent gonna be particually perfect! I hope this should be more or less bug free tho... I designed this to be for beginners / intermediatte needing a hand geting going in this field. Im sure other people are gonna have string manipulation tutorials already on PSC, but what the hell! have another one... :)

Beginners: if you have any problems getting the code to work, email me at the address at end.

tHE_cLeanER
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon Barker](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-barker.md)
**Level**          |Beginner
**User Rating**    |4.9 (147 globes from 30 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-barker-a-fully-i-hope-string-manipulation-example-tutorial-instr-replace-split-join-st__1-29350/archive/master.zip)





### Source Code

```
The following functions and procedures can be used to manipulate general strings, and more or less do whatever you like with them!
If you get stuck, look at the bottom of this tutorial for contact information. (Soz about the spelling.. im doin this in notepad!)
-------------------------------------
Right, first starting with the basic stuff:
1. Getting the length of a string or varible
msgbox len(text1.text)
Messages the number of characters in text1 text box. This will be in the form of a numerical value.
strText = "How long is this text?"
r = len(strText)
msgbox r
This produces a messagebox saying "22" because 'strText' is 22 charaters long.
-------------------------------------
2. The following code is used to get a part of a string. Useful for cutting off bits that arent needed throughtout the rest of the code.
msgbox Left("How are you today?", 3)
This pulls the 3 left characters of the specified text, and therefore produces a messagebox saying 'How'
text1.text = "Today is it Tuesday"
r = Left(text1.text, 5)
text1.text = r
The above code will first make the text1 textbox say 'Today is it Tuesday' then cut off everything except the day. After the code has been exucuted, the text1 textbox will read 'Today'; which is the left 5 characters, as specified in the code.
As well as the Left(string, number of places) code, there is Right(string, number of places). The Right function works the same way as left, but starts from the other side.
msgbox Right("6 Toffee Bakewells", 9)
This would produce a message reading 'Bakewells'.
Now you have the left and right parts of a string, you may wish to get a centre, or middle part. To do so, you use the Mid(string, Start point, length) function.
msgbox Mid("You now owe me 32 pounds", 16)
This would produce a message saying '32 Pounds' considering I set the code to start at 16 places into the string. You may have noticed that I left the end part of the code, the lengh off. This part is optional, and so if you dont specify the length, then it will go right to the end of the string. To get just the amount of money in the above code, you add the lengh to the end as so:
msgbox Mid("You now owe me 32 pounds", 16, 2)
This means read 16 places into 'You now owe me 32 pounds', and get the next 2 characters, which will be 32. Therefore you will get a message saying '32'.
---------------------------------------
3. If you wish to search the text for a particular word, then you will use the InStr(Start Place, Search String, Find word) function. This function is very customisable to your needs, and so has a lot of optional extras that can be added, but in the interests of simplicity, I'll leave these off the turorial. The InStr command returns its value as an integer (number) as a place where it found the string in the search text.
msgbox InStr(1, "The weather today is reasonably warm and sunny", "warm")
The above code starts at the beginning of 'The weather today is reasonably warm and sunny', as specifed by the 1 at start, and searches for the word 'warm' in it. If it does not find the word warm in the string, then it will return the value as 0, and you get a message saying '0'. However, if it finds the word, then it returns a number saying where it found the start of the word. In this case, you would see a messagebox saying '33' because the 'w' of warm is 33 characters into the string.
If you wish to make a simple search program, to find searchword text2.text in the string text1.text, then this is how you would go about doing it:
text1.text = "Welcome to the grand parade"
text2.text = "grand"
r = InStr(1, text1.text, text2.text)
if r > 0 then
	msgbox "Found word, " & r & " characters into the search string."
else
	msgbox "Sorry, could not find the search text"
end if
As well as that, there is the InStrRev command, which does exactly the same thing, but starts from the end of the search string. It will return an integer just as InStr does, which defines the placement of the word, but starting from the end. This is called as InStrRev(searchstring, findtext)
--------------------------------------
4. Next, is the Replace(search in string, search for text, replace with text). It is used to search through a string, and replace certain words or characters with other ones. If you want an active example of this one, look in my 'other uploads by thi person' on PSC, to find 'the lamerizer'.
msgbox replace("Only a fool goes outside in the cold without a coat on", "fool", "brave bloke")
This code would produce a message replaceing 'fool' with 'brave bloke', and therefore will look like this: 'Only a brave bloke goes outside in the cold without a coat on'.
Another example of this use, is to remove a swearword from a sentance etc, as follows:
text1.text = replace(text1.text, "oh my god", "oh my goodness")
This code searches through text1.text textbox, and replaces any instances of 'oh my god', with 'oh my goodness', then returns the text back into text1.text, without the cursing.
----------------------------------------
5. Setting uppercase or lowercase a user has typed.
This is useful for making sure that if a user types something in uppercase (capitals) then it will still comply with something in your code that is lowercase. For example, if you are making a text adventure, and the user is given a choice of left or right, and they type LEFT, as VB is case sensitive, your program would'nt accept their answer, adnd tell them it was invalid!
To combat this, you use the LCase(string) or UCase(string) commands
To make a sentance lowercase, you use the following:
text1.text = LCase(text1.text)
Or to convert to uppercase, use the following:
text1.text = UCase(text1.text)
-----------------------------------------
6. Reversing the order of characters in a string.
If you wish to flip around the front and back end of a string, then the StrReverse(string) is for you. It is used in the following way:
msgbox StrReverse("PSC is a rather large database")
This would pop up a message saying 'esabatad egral rehtar a si CSP'. Im not quite sure why you'd want to use this function, but may be usefull to know!
------------------------------------------
7. Comparing strings in terms of ASCII values / Case.
The StrComp function seems reaonably usefull in this feild. It is used in context StrComp(string1, string2).
This function returns its value as an integer, specifying what it found.
text1.text = strComp("tHe_cLeanER", "THE_CLEANER")
if text1.text = -1 then msgbox "String 1 is less than string 2"
if text1.text = 0 then msgbox "String 2 is equal to string 1"
if text1.text = 1 then msgbox "String 1 is greater than string 2"
if text1.text = Null then msgbox "String 1 and / or string two is null"
In this case, you would get text1.text texbox giving you the value 1, because tHe_cLeanER is greater in ascii value than THE_CLEANER.
------------------------------------------
8. Creating arrays with the Split(string1, split character) function.
This function allows you to create a one-dimensional array, by splitting a string by reconising a certain character, then putting any text after the character on a new line in the array.
Basic use of this function could be for getting a list of names from a multiline text box as follows:
r = Split(Text1.Text, Chr(13))
For i = 0 To UBound(r)
  MsgBox r(i)
Next i
This will pull all lines of the text box, and use them to create an array, which is stored in r. You extract these values from the array by selecting where in the array you wish to look. The look-in-line is defined after the r, inb brackets. Example: Msgbox r(3) would pull the FORTH line of the array that is being held in r. Msgbox r(5) would pull the 6th line being held in the array.
------------------------------------------
9. Joining an array back into one string. Uses the Join(array string, split character) function.
If you have an array, and wish to compile it back into one string, then the Join function (Which is the opposite of the Split function) is the one to use.
Note: this will only work if r contains an array. See previous to create an array.
z = Join(r, Chr(13))
MsgBox z
This code will put back together an array into a string, seperating different lines in the array with the specified character. In this case, i used the carrige return char, which is the equivilent of pressing Enter. The above code will compile an array created from a multiline text box. It will work fine with the previous procedure.
------------------------------------------
Theres a few more string manipulation comands that i havent gone into, possibly becuase im board of typing!
neway.. im sorta hoping most of this code will work, if you have any probs, mail me on jBistoGOOD@Hotmail.com, and i'll see what can be done...
keep coding!
tHe_cLeanER
```

