# About
This tool extracts clips from multiple video files, resizes them to the same resolution, and combines them into a single "recap" video file, with crossfade and audio normalisation, as well as custom subtitles. It is compatible with Linux, Mac, and Windows.

# Requirements
This tool uses the third party Python modules OpenPyXL and MoviePy. Furthermore, MoviePy requires ImageMagick. Please install these before attempting to run the Python scripts included here. Note that MoviePy uses features of Pillow 9.5.0 that are no longer supported in the most recent version of Pillow, so you may need to roll back your version of Pillow. Do this by typing `pip install Pillow==9.5.0` in the terminal.

If you are on Linux/Mac, you will need to edit ImageMagick's access permissions in `policy.xml` in order for MoviePy to work properly:

1. Type `identify -list policy` in the terminal. The first line should tell you where your `policy.xml` file is located.
2. Type `sudo nano /etc/ImageMagick-6/policy.xml` in the terminal, replacing `/etc/ImageMagick-6/policy.xml` with the address of your `policy.xml` file.
3. Directly edit the file in the terminal, removing the line that reads `<policy domain="path" rights="none" pattern="@*" />`. This line should be near the end of the file.
4. Press Ctrl+O to save the changes, then Ctrl+X to exit the editor.

If you are on Windows, you will need to specify the path to ImageMagick in MoviePy's configuration settings:

1. Find `magick.exe` on your computer. Right-click the file and click 'Copy as path'.
2. Find the `config_defaults.py` file in MoviePy's folder and open it in Notepad or the editor of your choice.
3. Find the line that says `IMAGEMAGICK_BINARY = os.getenv('IMAGEMAGICK_BINARY', 'auto-detect')` and change it to `IMAGEMAGICK_BINARY = r"C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"`, where the part in quotation marks should be the path you just copied.
4. Save the changes.

# Instructions
1. Put the video files that are to be used for the recap in the `Videos` folder provided.
2. Enter the following data into `video_data.xlsx`:
- the filenames of the video files in the `Videos` folder, including extensions;
- the start and end times of the video clips you wish to include in the recap, in the form `hh:mm:ss`;
- three lines of subtitles for each clip.
3. Run `recap_generator.py`.