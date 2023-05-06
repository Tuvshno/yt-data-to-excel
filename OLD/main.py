from googleapiclient.discovery import build
import pandas as pd
import seaborn as sns

api_key = 'AIzaSyCyYE5tgXV7HJ_QCtQevzY66QJCIACBbRY'
channel_id = 'UCWIfnDrWU_Cvc1a8qZribhA'

youtube = build('youtube', 'v3', developerKey=api_key)

def get_channel_stats(youtube, channel_id):
    request = youtube.channels().list(part='snippet,contentDetails,statistics', id=channel_id)
    response = request.execute()

    data = response['item'][0]

    return data

stats = get_channel_stats(youtube, channel_id)
print(stats)


