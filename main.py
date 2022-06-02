import requests
import json
from win32com.client import Dispatch
def speak(str):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)

def topic_num(no):
    if(no==1):
        return "business"
    elif(no==2):
        return "entertainment"
    elif(no==3):
        return "health"    
    elif(no==4):
        return "science"
    elif(no==5):
        return "sports"
    elif(no==6):
        return "technology"
    else:
        return None


print("Choose the topic of the news.\n 1. Business \n 2. Entertainment \n 3. Health \n 4. Science \n 5. Sports \n 6. Technology \n 00. For all")
choice = int(input("Enter your choice: "))

topiccccc = "business" if choice == 1 else "entertainment" if choice == 2 else "health" if choice == 3 else "science" if choice == 4 else "sports" if choice == 5 else "technology" if choice == 6 else "none" 


if topiccccc == "none":
    url = 'https://newsapi.org/v2/top-headlines?country=in&apiKey={API_KEY}'
else:
    url = f'https://newsapi.org/v2/top-headlines?country=in&category={topiccccc}&apiKey={API_KEY}'
    print(f"Reading best {topiccccc} news for you...")


response = requests.get(url).text

text = json.loads(response)
for num in range(0, 5):
    print(text['articles'][num]['title'])
    speak(text['articles'][num]['title'])
    if text['articles'][num]['description'] == None or text['articles'][num]['description'] == False:
        pass
    else:
        print(text['articles'][num]['description'])
        speak(text['articles'][num]['description'])
    print("\n")
    if num <= 3:
        speak("Next news is...")
    else:
        speak("That was the last news.")
    

speak("Thank you for using this tool.")
