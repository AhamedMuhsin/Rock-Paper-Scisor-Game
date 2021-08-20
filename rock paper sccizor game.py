import random

def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

token = ["r","p","s"]

print("\t\t\t\t\tWelcome to Rock Paper and Scissor Game\n")
speak("\t\t\t\t\tWelcome to Rock Paper and Scissor Game\n")

print("Instruction of this Game\n1. In this game you have to select any of this token given here\n"
      "2. And then computer will randomly chose any token\n3. You have to type only the first letter of the token\n"
      "4. In this Game we have point system so if your point is less you will Lose\n")
speak("Instruction of this Game\n1. In this game you have to select any of this token given here\n"
      "2. And then computer will randomly chose any token\n3. You have to type only the first letter of the token\n"
      "4. In this Game we have point system so if your point is less you will Lose\n")

speak("Enter your name : ")
name = input("Enter your name : ")

print("Choose r for rock\n""Choose p for paper\n""Choose s for scissor")
speak("Choose r for rock\n""Choose p for paper\n""Choose s for scissor")

chance = 5
no_of_chance = 0
Human_point = 0
Muhsin_point = 0

while no_of_chance < chance:
    _player = input(" Rock , Paper , Scissor : ")
    _random = random.choice(token)

    if _player == _random:
        print(f"Tie both {name} and Muhsin have no point")
        speak(f"Tie both {name} and Muhsin have no point")

    elif _player == "r" and _random == "p":
        Muhsin_point = Muhsin_point + 1
        print(f"{name} guess {_player} and Muhsin guess {_random}\n")
        speak(f"{name} guess {_player} and Muhsin guess {_random}\n")
        print("Muhsin wins 1 point\n")
        speak("Muhsin wins 1 point\n")
        print(f"Muhsin point is {Muhsin_point} and {name} point is {Human_point}\n")
        speak(f"Muhsin point is {Muhsin_point} and {name} point is {Human_point}\n")

    elif _player == "r" and  _random == "s":
        Human_point = Human_point + 1
        print(f"{name} guess {_player} and Muhsin guess {_random}\n")
        speak(f"{name} guess {_player} and Muhsin guess {_random}\n")
        print(f"{name} wins 1 point\n")
        speak(f"{name} wins 1 point\n")
        print(f"Muhsin point is {Muhsin_point} and {name} point is {Human_point}\n")
        speak(f"Muhsin point is {Muhsin_point} and {name} point is {Human_point}\n")

    elif _player == "p" and _random == "s":
        Muhsin_point = Muhsin_point + 1
        print(f"{name} guess {_player} and Muhsin guess {_random}\n")
        speak(f"{name} guess {_player} and Muhsin guess {_random}\n")
        print("Muhsin wins 1 point\n")
        speak("Muhsin wins 1 point\n")
        print(f"Muhsin_point is {Muhsin_point} and {name} point is {Human_point}\n")
        speak(f"Muhsin point is {Muhsin_point} and {name} point is {Human_point}\n")

    elif _player == "p" and  _random == "r":
        Human_point = Human_point + 1
        print(f"{name} guess {_player} and Muhsin guess {_random}\n")
        speak(f"{name} guess {_player} and Muhsin guess {_random}\n")
        print(f"{name} wins 1 point\n")
        speak(f"{name} wins 1 point\n")
        print(f"Muhsin_point is {Muhsin_point} and {name} point is {Human_point}\n")
        speak(f"Muhsin point is {Muhsin_point} and {name} point is {Human_point}\n")

    elif _player == "s" and _random == "r":
        Muhsin_point = Muhsin_point + 1
        print(f"{name} guess {_player} and Muhsin guess {_random}\n")
        speak(f"{name} guess {_player} and Muhsin guess {_random}\n")
        print("Muhsin wins 1 point\n")
        speak("Muhsin wins 1 point\n")
        print(f"Muhsin_point is {Muhsin_point} and {name} point is {Human_point}\n")
        speak(f"Muhsin point is {Muhsin_point} and {name} point is {Human_point}\n")

    elif _player == "s" and  _random == "p":
        Human_point = Human_point + 1
        print(f"{name} guess {_player} and Muhsin guess {_random}\n")
        speak(f"{name} guess {_player} and Muhsin guess {_random}\n")
        print(f"{name} wins 1 point\n")
        speak(f"{name} wins 1 point\n")
        print(f"Muhsin point is {Muhsin_point} and {name} point is {Human_point}\n")
        speak(f"Muhsin point is {Muhsin_point} and {name} point is {Human_point}\n")


    else:
        print(f"{name} your input is wrong please input right token letter")
        speak(f"{name} your input is wrong please input right token letter")
    no_of_chance = no_of_chance + 1
    print(f"{chance - no_of_chance} chance is left out of {chance}\n")
    speak(f"{chance - no_of_chance} chance is left out of {chance}\n")

if Muhsin_point == Human_point:
    print(f"Tie both {name} and Muhsin have same point")
    speak(f"Tie both {name} and Muhsin have same point")

elif Muhsin_point > Human_point:
    print(f"{name} Lose and Muhsin Wins")
    speak(f"{name} Lose and Muhsin Wins")

else:
    print(f"{name} Wins and Muhsin Lose")
    speak(f"{name} Wins and Muhsin Lose")

print(f"{name} point is {Human_point} and Muhsin point is {Muhsin_point}\n")
speak(f"{name} point is {Human_point} and Muhsin point is {Muhsin_point}\n")

print("\t\t\t\t\t Thanks for playing this Game\n")
speak("\t\t\t\t\t Thanks for playing this Game\n")









