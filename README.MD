# AutoAdder

AutoAdder is a Python script that automates the process of downloading audio from YouTube and adding it to Soundpad.

## Features

- Downloads audio from a YouTube URL
- Sanitizes the filename to ensure it's valid
- Connects to Soundpad's named pipe
- Sends a request to Soundpad and reads the response
- Adds the downloaded sound file to Soundpad
  
## Installation

Before running the script, you need to install the required Python libraries and `ffmpeg`. You can do this by running the following command in your terminal:

```bash
pip install yt-dlp pyperclip pywin32

## Usage

1. Copy a YouTube URL to your clipboard.
2. Run the `autoadder.py` script.

The script will automatically download the audio from the YouTube URL in your clipboard, sanitize the filename, and add it to Soundpad.

## Dependencies

This script uses the following Python libraries:

- os
- re
- yt_dlp
- win32pipe, win32file, pywintypes
- pyperclip

## Functions

- `sanitize_filename(name)`: Sanitizes the filename by removing emojis, non-ASCII characters, and special characters.
- `download_youtube_audio(youtube_url)`: Downloads audio from a YouTube URL and converts it to MP3.
- `connect_to_soundpad()`: Connects to Soundpad's named pipe.
- `send_request(handle, request)`: Sends a request to Soundpad and reads the response.
- `add_sound_to_soundpad(file_path)`: Adds a sound file to Soundpad without specifying a category.
- `get_url_from_clipboard()`: Gets a URL from the clipboard.
- `main()`: The main function that runs the script.

## License

[MIT](https://choosealicense.com/licenses/mit/)

## Acknowledgements

The implementation of the connection to Soundpad's named pipe in this project was inspired by the [Soundpad Connector](https://github.com/medokin/soundpad-connector) repository. We appreciate the authors for their valuable contribution.