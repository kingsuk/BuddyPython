import speech_recognition as sr

r = sr.Recognizer()
s = sr.Microphone()

def getSpeechToText():
    
    with s as source:
        #print("Please wait. Calibrating microphone...")
        # listen for 1 second and create the ambient noise energy level
        #r.adjust_for_ambient_noise(s, duration=1)
        print("Say something!")
        audio = r.listen(source,phrase_time_limit=10)
        
 
    try:
        # for testing purposes, we're just using the default API key
        # to use another API key, use `r.recognize_google(audio, key="GOOGLE_SPEECH_RECOGNITION_API_KEY")`
        # instead of `r.recognize_google(audio)`
        result = r.recognize_google(audio)
        print(result)
        return result.lower()
        
    except sr.UnknownValueError:
        print("Google Speech Recognition could not understand audio")
        return getSpeechToText()
    except sr.RequestError as e:
        print("Could not request results from Google Speech Recognition service; {0}".format(e))
        return getSpeechToText()
