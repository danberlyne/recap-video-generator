# About
This tool extracts clips from multiple video files, resizes them to the same resolution, and combines them into a single "recap" video file, with crossfade and audio normalisation, as well as custom subtitles.

The user can either specify their desired clips manually or let `recap-video-generator` automatically select the best clip for each video by detecting choruses within the audio.

# Installation
## Python script
This script `recap_generator.py` uses the third party Python modules `OpenPyXL`, `MoviePy`, `Pychorus`, `jsonschema` and `ruamel.yaml`. Furthermore, `Pychorus` requires the non-Python package `FFmpeg`. Please install these before attempting to run the Python script.

If you are on Linux, install the required packages via the following commands:

```sh
pip install openpyxl moviepy pychorus jsonschema ruamel.yaml
sudo apt install ffmpeg
```

If you are on Windows:

```sh
pip install openpyxl moviepy pychorus jsonschema ruamel.yaml
winget install ffmpeg
```

If you are on Mac:

```sh
pip install openpyxl moviepy pychorus jsonschema ruamel.yaml
brew install ffmpeg
```

## Windows executable
An executable for Windows users is available at [https://berlyne.net/code/recap-generator](https://berlyne.net/code/recap-generator). If using the executable, the only requirement is installation of FFmpeg:

```sh
winget install ffmpeg
```

# Usage
1. Put the video files that are to be used for the recap in the `Videos` folder provided.
2. Enter the following data into `video_data.xlsx`:
    - The filenames of the video files in the `Videos` folder, including extensions;
    - (Optional) The start and end times of the video clips you wish to include in the recap, in the form `hh:mm:ss`;
    - (Optional) Three lines of subtitles for each clip.
3. Specify any options for recap generation in the `options.yaml` file.
    - This includes input/output file locations, clip selection options, subtitle properties, and intro clip properties.
    - Note: If you specify that you are using your first clip as an intro clip, any subtitles for that clip will automatically be placed in the centre of the screen. This may be used as an alternative to a custom image overlay.
4. Run `recap_generator.py` (or `recap_generator.exe` if using the Windows executable version).

Note, if the user does not specify a start and end time for a given clip, `recap-video-generator` will attempt to detect a chorus within the audio and select a clip automatically.