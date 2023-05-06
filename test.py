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

def main():


    try:
        api_version = "v3"
        api_service_name = "youtube"

        client_secrets_file = "D:\Dev\yt-excel\client_secret_137710264875-mngrtgrrojtm1opqi1hl36mga30o8m3b.apps.googleusercontent.com.json"  # Replace with the path to your client_secret.json file

        # Get credentials and create an API client
        scopes = ["https://www.googleapis.com/auth/youtube.force-ssl"]
        flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(client_secrets_file, scopes)
        credentials = flow.run_local_server(port=0)
        youtube = googleapiclient.discovery.build(api_service_name, api_version, credentials=credentials)

        channel_id = "UCbz_aqiqO2HcAx01iJM5WGA"  # Replace with your channel ID
        stop_date = "2022-02-24"  # Set the stop date (YYYY-MM-DD)
        videos = get_channel_videos(channel_id, youtube, stop_date=stop_date)

        videos_data = []
        for video in videos:
            video_id = video['id']['videoId']
            video_title = video['snippet']['title']
            video_upload_date = video['snippet']['publishedAt'][:10]
            video_link = video['videoLink']  # Extract the video link from the item

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
        
    except googleapiclient.errors.HttpError as e:
        if "quotaExceeded" in str(e):
            print("API quota exceeded. Saving the data collected so far.")
            output_file = 'channel_videos_partial.xlsx'
            write_to_excel(videos_data, output_file)
        else:
            raise e
        
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

if __name__ == "__main__":
    main()
