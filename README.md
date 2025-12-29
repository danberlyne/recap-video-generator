# About
This tool extracts clips from multiple video files, resizes them to the same resolution, and combines them into a single "recap" video file, with crossfade and audio normalisation, as well as custom subtitles.

The user can either specify their desired clips manually or let `recap-video-generator` automatically select the best clip for each video by detecting choruses within the audio.

# Installation
There are three options: either use the Python script directly, run the program inside a Docker container, or run as an executable (Windows only).

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

## Docker
First, put the video files to be used in the recap inside the `Videos` folder provided, specify any options for recap generation in `options.yaml`, and enter the clip data in `video_data.xlsx` (video filenames and optionally clip start/end times and subtitles). 

Then navigate to the `recap-video-generator` directory in the terminal and run the following commands:

```sh
docker build -t recap-generator:latest .
docker run --rm -v $PWD:/src recap-generator:latest
```

Note that the Docker container will only be able to see files within the `recap-video-generator` directory, so make sure any custom file locations specified in `options.yaml` (e.g. output file location, fonts, intro image) are within this directory and are specified as relative paths.

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
4. Run the recap generator (either by running `recap_generator.py` or by running the Docker container or by running the Windows executable `recap_generator.exe`)

Note, if the user does not specify a start and end time for a given clip, `recap-video-generator` will attempt to detect a chorus within the audio and select a clip automatically.