import requests
import json

def speak(str):
    from win32com.client import Dispatch
    speak = Dispatch("Sapi.SpVoice")
    speak.Speak(str)

if __name__ == '__main__':
    speak("Some of todays news are")
    url = 'https://newsapi.org/v2/top-headlines?sources=the-times-of-india&apiKey=d093053d72bc40248998159804e0e67d'
    news = requests.get(url).text
    news_dict = json.loads(news)
    # print(news_dict["articles"])
    arts = news_dict['articles']
    for articles in arts:
        speak(articles['title'])
        speak("Moving to next news... ")
    speak("Thank you!!! That's all for todays news")