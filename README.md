# About
This tool selects clips from video files combines them into a single "recap" video file, with crossfade and audio normalisation, as well as custom subtitles.

# Requirements
This tool uses the third party Python modules OpenPyXL and MoviePy. Furthermore, MoviePy requires ImageMagick. Please install these before attempting to run the Python scripts included here.

You may also need to follow the instructions found [here](https://github.com/Zulko/moviepy/issues/401) to edit ImageMagick's access permissions in order for MoviePy to work properly. If you do not have write permissions for ImageMagick's `policy.xml` file, you can use [this alternative method](https://stackoverflow.com/questions/52998331/imagemagick-security-policy-pdf-blocking-conversion) to edit the file.

# Instructions
1. Put the video files that are to be used for the recap in the `Videos` folder.
2. Enter the following data into `video_data.xlsx`:
- the filenames of the video files in the `Videos` folder, including extensions;
- the start and end times of the video clips you wish to include in the recap, in the form `hh:mm:ss`;
- three lines of subtitles for each clip.
3. Run `recap_generator.py`.