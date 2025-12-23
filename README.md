# About
This tool extracts clips from multiple video files, resizes them to the same resolution, and combines them into a single "recap" video file, with crossfade and audio normalisation, as well as custom subtitles.

The user can either specify their desired clips manually or let `recap-video-generator` automatically select the best clip for each video by detecting choruses within the audio.

# Requirements
This tool uses the third party Python modules OpenPyXL, MoviePy, and Pychorus. Furthermore, Pychorus requires FFmpeg. Please install these before attempting to run the Python scripts included here.

If you are on Linux, install via the following commands:

```
pip install openpyxl moviepy pychorus
sudo apt install ffmpeg
```

If you are on Windows:

```
pip install openpyxl moviepy pychorus
winget install ffmpeg
```

# Instructions
1. Put the video files that are to be used for the recap in the `Videos` folder provided.
2. Enter the following data into `video_data.xlsx`:
    - The filenames of the video files in the `Videos` folder, including extensions;
    - (Optional) The start and end times of the video clips you wish to include in the recap, in the form `hh:mm:ss`;
    - (Optional) Three lines of subtitles for each clip.
3. (Optional) If you wish to use a custom image overlay on your first clip, place your image in the same directory as `video_data.xlsx` and name it `intro.png`.
4. (Optional) If you wish to use a custom font for subtitles, place your font file (in `.ttf` format) in the `Fonts` folder provided.
5. Run `recap_generator.py` and follow the on-screen instructions.
    - Note: If you specify that you are using your first clip as an intro clip, any subtitles for that clip will automatically be enlarged and placed in the centre of the screen. This may be used as an alternative to a custom image overlay.

Note, if the user does not specify a start and end time for a given clip, `recap-video-generator` will attempt to detect a chorus within the audio and select a clip automatically.