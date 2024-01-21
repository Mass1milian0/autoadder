import os
import re
import yt_dlp as youtube_dl
import win32pipe, win32file, pywintypes
import pyperclip

def sanitize_filename(name):
    name = re.sub(r'[^\x00-\x7F]+', '', name)
    name = re.sub(r'[\\/*?:"<>|]', '', name)
    name = re.sub(r'[\s]+', '_', name)
    return name

def download_youtube_audio(youtube_url):
    output_directory = os.getcwd()
    with youtube_dl.YoutubeDL() as ydl:
        info_dict = ydl.extract_info(youtube_url, download=False)
        sanitized_title = sanitize_filename(info_dict['title'])
        full_output_path = os.path.abspath(os.path.join(output_directory, sanitized_title + '.mp3'))

        ydl_opts = {
            'format': 'bestaudio/best',
            'postprocessors': [{
                'key': 'FFmpegExtractAudio',
                'preferredcodec': 'mp3',
                'preferredquality': '192',
            }],
            'outtmpl': full_output_path,
        }

        with youtube_dl.YoutubeDL(ydl_opts) as ydl:
            ydl.download([youtube_url])

    return full_output_path

def connect_to_soundpad():
    pipe_name = r'\\.\pipe\sp_remote_control'
    try:
        return win32file.CreateFile(
            pipe_name,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0,
            None,
            win32file.OPEN_EXISTING,
            0,
            None
        )
    except pywintypes.error:
        return None

def send_request(handle, request):
    win32file.WriteFile(handle, request.encode())
    return win32file.ReadFile(handle, 4096)[1].decode().strip('\x00')

def add_sound_to_soundpad(file_path):
    handle = connect_to_soundpad()
    if handle:
        raw_file_path = fr"{file_path}"
        request = f"DoAddSound(\"{raw_file_path}\")"
        send_request(handle, request)
        win32file.CloseHandle(handle)

def get_url_from_clipboard():
    clipboard_content = pyperclip.paste()
    if "youtube.com" in clipboard_content or "youtu.be" in clipboard_content:
        return clipboard_content
    return None

def main():
    youtube_url = get_url_from_clipboard()
    if youtube_url:
        file_path = download_youtube_audio(youtube_url)
        add_sound_to_soundpad(file_path)

if __name__ == "__main__":
    main()
