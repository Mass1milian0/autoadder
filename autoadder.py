import os
import re
import yt_dlp as youtube_dl
import win32pipe, win32file, pywintypes
import pyperclip

# Function to sanitize file names
def sanitize_filename(name):
    # Remove emojis and non-ASCII characters
    name = re.sub(r'[^\x00-\x7F]+', '', name)
    # Remove or replace special characters
    name = re.sub(r'[\\/*?:"<>|]', '', name)  # Removes characters that are invalid in Windows file names
    name = re.sub(r'[\s]+', '_', name)  # Replace spaces and similar characters with underscores
    return name

# Function to download audio from YouTube and convert to MP3
def download_youtube_audio(youtube_url):
    output_directory = os.getcwd()  # Get the program's current working directory
    with youtube_dl.YoutubeDL() as ydl:
        info_dict = ydl.extract_info(youtube_url, download=False)
        sanitized_title = sanitize_filename(info_dict['title'])
        full_output_path = os.path.abspath(os.path.join(output_directory, sanitized_title))

        # Calculate the relative path from the current working directory to the output directory
        rel_output_path = os.path.relpath(full_output_path, os.getcwd())

        ydl_opts = {
            'format': 'bestaudio/best',
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': '192',
            }],
            'outtmpl': rel_output_path + '.%(ext)s',  # Use the relative output path
        }

    with youtube_dl.YoutubeDL(ydl_opts) as ydl:
        ydl.download([youtube_url])
        return full_output_path + '.mp3'  # Add the '.mp3' extension to the returned path



# Function to connect to Soundpad's named pipe
def connect_to_soundpad():
    pipe_name = r'\\.\pipe\sp_remote_control'
    try:
        handle = win32file.CreateFile(
            pipe_name,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0,
            None,
            win32file.OPEN_EXISTING,
            0,
            None
        )
        print("Successfully connected to Soundpad's named pipe.")
        return handle
    except pywintypes.error as e:
        print(f"Error connecting to pipe: {e}")
        return None

# Function to send a request to Soundpad and read the response
def send_request(handle, request):
    data = str.encode(request)
    win32file.WriteFile(handle, data)
    resp = win32file.ReadFile(handle, 4096)
    return resp[1].decode().strip('\x00')

def add_sound_to_soundpad(file_path):
    handle = connect_to_soundpad()
    if handle:
        # Use a raw string for the file path
        raw_file_path = fr"{file_path}"
        quoted_file_path = f'"{raw_file_path}"'
        request = f"DoAddSound({quoted_file_path})"
        print(f"Sending request: {request}")  # Debug print
        response = send_request(handle, request)
        print(f"Response from Soundpad: {response}")
        win32file.CloseHandle(handle)

# Function to get URL from clipboard
def get_url_from_clipboard():
    clipboard_content = pyperclip.paste()
    if "youtube.com" in clipboard_content or "youtu.be" in clipboard_content:
        return clipboard_content
    else:
        print("No YouTube URL found in clipboard.")
        return None

# Main function
def main():
    youtube_url = get_url_from_clipboard()
    if youtube_url:
        print(f"Downloading audio from: {youtube_url}")
        file_path = download_youtube_audio(youtube_url)
        absolute_file_path = os.path.abspath(file_path)
        print(f"Downloaded audio file path: {absolute_file_path}")
        add_sound_to_soundpad(absolute_file_path)

if __name__ == "__main__":
    main()