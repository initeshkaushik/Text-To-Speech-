# imports the win32com.client module and assigns it an alias of wincl
import win32com.client as wincl

# checks if the current script is being executed as the main module
if __name__ == '__main__':
    print("Welcome to Text to Speech Speaker Created by Nitesh")

    # starts infinite while loop to enter the text and exit when click on q
    while True:
        x = input("Enter what you want to be pronounced: ")
        if x == "q":  # if user inputs 'q', break out of the loop
            speaker = wincl.Dispatch("SAPI.SpVoice")
            speaker.Speak("thanks for using")  # use text-to-speech to say "thanks for using"
            break

        # this code converts the text to voice
        speaker = wincl.Dispatch("SAPI.SpVoice")
        speaker.Speak(x)  # use text-to-speech to speak the entered text