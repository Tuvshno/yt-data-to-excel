import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import google_auth_oauthlib.flow
import googleapiclient.discovery
import googleapiclient.errors
import datetime
import xlsxwriter


def get_channel_videos(channel_id, youtube, stop_date=None):
    if stop_date:
        stop_date = datetime.datetime.strptime(stop_date, "%Y-%m-%d")

    videos = []
    next_page_token = None
    stop_date_reached = False

    while not stop_date_reached:
        try:
            request = youtube.search().list(
                part="snippet",
                channelId=channel_id,
                maxResults=50,
                pageToken=next_page_token,
                type="video",
                order="date"  # Order by upload date (most recent first)
            )
            response = request.execute()

            if not response["items"]:
                break

            for item in response["items"]:
                video_date = datetime.datetime.fromisoformat(item["snippet"]["publishedAt"][:-1])

                if stop_date and video_date < stop_date:
                    stop_date_reached = True
                    break
                
                # Add the video link to the video dictionary
                video = item.copy()
                video['videoLink'] = f"https://www.youtube.com/watch?v={item['id']['videoId']}"

                videos.append(video)

            next_page_token = response.get("nextPageToken")

        except googleapiclient.errors.HttpError as e:
            if "quotaExceeded" in str(e):
                print("API quota exceeded. Saving the data collected so far.")
                break
            else:
                raise e

        if not next_page_token or stop_date_reached:
            break

    return videos

def get_video_details(video_ids, youtube):
    request = youtube.videos().list(
        part="snippet,statistics",
        id=','.join(video_ids)
    )
    response = request.execute()
    return response["items"]

def get_video_statistics(video_id, youtube):
    request = youtube.videos().list(
        part="statistics",
        id=video_id
    )
    response = request.execute()
    return response["items"][0]["statistics"]

def write_to_excel(videos_data, output_file):
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    headers = ['Title', 'Views', 'Upload Date', 'Video Link']
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    for row_num, video_data in enumerate(videos_data, start=1):
        worksheet.write(row_num, 0, video_data['title'])
        worksheet.write(row_num, 1, video_data['views'])
        worksheet.write(row_num, 2, video_data['upload_date'])
        worksheet.write(row_num, 3, video_data['video_link'])

    workbook.close()

class YoutubeDataInput(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("YouTube Data Input")

        # Variables
        self.client_secret_file_path = tk.StringVar()
        self.youtube_channel_id = tk.StringVar()
        self.stop_date = tk.StringVar()

        # Labels
        tk.Label(self, text="Client Secret File:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        tk.Label(self, text="YouTube Channel ID:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        tk.Label(self, text="Stop Date (YYYY-MM-DD):").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)

        # Entries
        self.client_secret_entry = tk.Entry(self, textvariable=self.client_secret_file_path, width=40)
        self.client_secret_entry.grid(row=0, column=1, padx=5, pady=5)

        self.youtube_channel_entry = tk.Entry(self, textvariable=self.youtube_channel_id, width=40)
        self.youtube_channel_entry.grid(row=1, column=1, padx=5, pady=5)

        self.stop_date_entry = tk.Entry(self, textvariable=self.stop_date, width=40)
        self.stop_date_entry.grid(row=2, column=1, padx=5, pady=5)

        # Buttons
        self.browse_button = tk.Button(self, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=5, pady=5)

        self.submit_button = tk.Button(self, text="Submit", command=self.submit)
        self.submit_button.grid(row=3, columnspan=2, pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(title="Select Client Secret File", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if file_path:
            self.client_secret_file_path.set(file_path)

    def submit(self):
        client_secret_path = self.client_secret_file_path.get()
        youtube_id = self.youtube_channel_id.get()
        stop_date_str = self.stop_date.get()

        try:
            stop_date_obj = datetime.datetime.strptime(stop_date_str, '%Y-%m-%d')
            
            api_version = "v3"
            api_service_name = "youtube"

            # Get credentials and create an API client
            scopes = ["https://www.googleapis.com/auth/youtube.force-ssl"]
            flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(client_secret_path, scopes)
            credentials = flow.run_local_server(port=0)
            youtube = googleapiclient.discovery.build(api_service_name, api_version, credentials=credentials)

            videos = get_channel_videos(youtube_id, youtube, stop_date=stop_date_str)

            videos_data = []
            for video in videos:
                video_id = video['id']['videoId']
                video_title = video['snippet']['title']
                video_upload_date = video['snippet']['publishedAt'][:10]
                video_link = video['videoLink']

                video_statistics = get_video_statistics(video_id, youtube)
                video_views = int(video_statistics['viewCount'])

                videos_data.append({
                    'title': video_title,
                    'views': video_views,
                    'upload_date': video_upload_date,
                    'video_link': video_link
                })

            output_file = 'channel_videos.xlsx'
            write_to_excel(videos_data, output_file)
            print("Data saved to", output_file)
            self.destroy() 

        except ValueError:
            print("Invalid date format. Please use YYYY-MM-DD.")

if __name__ == "__main__":
    app = YoutubeDataInput()
    app.mainloop()