# hangman-game
This is a Hangman game project, implemented using VBA (Visual Basic for Applications) in Microsoft Excel. This game is a classic interpretation of the popular word game, with additional features and a user-friendly interface.

Functionality
- Random and Custom Words: Players can start the game with a randomly generated word or enter their own word.
- Interactive Interface: The game offers simple and intuitive user interactions through buttons and messages.
- Graphical Representation of the Hangman: The game progress is visually represented by a hangman drawing, which changes with the number of mistakes.
- Score Management: The game tracks the number of errors made and displays the results.
  
Installation
Excel File: To play the game, open the Excel file containing the project.
Macros: Ensure that macros are enabled in Excel to allow the game to function.

How to Play
Starting the Game: Press the 'Start' button to begin a new game. You can choose between entering your own word or a randomly generated word.
Guessing Letters: Click the 'Submit' button to input a letter. If the letter is in the word, it will be displayed.
Tracking Progress: The number of mistakes and progress in guessing the word are tracked and visualized on the board.
Ending the Game: The game ends when all letters are guessed (win) or when the hangman drawing is completed (loss).

Code Components
IsInArray: Checks if an element is in an array.
GetRandomWord: Retrieves a random word from an Excel sheet.
CzyWygrana (IsWin): Checks if the user has guessed the whole word.
IsLetter: Verifies if the input character is a letter.
Button_Podaj_Click (Button_Submit_Click): Logic for the letter submission button.
Button_reset_Click: Resets the game to its initial state.
Button_Start_Click: Starts a new game.

Support and Collaboration
I invite collaboration on the development of this project. If you have suggestions, questions, or want to report a bug, please contact me through GitHub.

Note: The code includes elements specific to the Polish language, including Polish diacritical characters. Please adjust the code as needed for your language.
