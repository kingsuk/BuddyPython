import requests

def AnalyseIntent(userInput):
    headers = {
        # Request headers
        'Ocp-Apim-Subscription-Key': '08e66a5c2e024e28963d0b23e1702b15',
    }
    params ={
        # Query parameter
        'q': userInput,
        # Optional request parameters, set to default values
        'timezoneOffset': '0',
        'verbose': 'false',
        'spellCheck': 'false',
        'staging': 'false',
    }

    try:
        r = requests.get('https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/7c51f1a4-4c38-4c6a-8929-337c6e5c7e6f',headers=headers, params=params)
        return r.json()

    except Exception as e:
        return "[Errno {0}] {1}".format(e.errno, e.strerror)
        