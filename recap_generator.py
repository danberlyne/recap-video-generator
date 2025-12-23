#!/usr/bin/env python3
# recap_generator.py - Extracts clips from video files and combines them into a single recap video.

import openpyxl
from pathlib import Path
from moviepy import VideoFileClip, CompositeVideoClip, ImageClip, TextClip
import moviepy.audio.fx as afx
import moviepy.video.fx as vfx
from moviepy.video.tools.subtitles import SubtitlesClip
from pychorus import find_and_output_chorus
from typing import Literal, Union
import warnings

warnings.filterwarnings('ignore')

# `sub_alignment` is the alignment of subtitles in the video (left/center/right).
# `clip_selection_method` determines whether to use `pychorus` to detect clips (auto) or whether to use user-defined clips (manual).
def generate_recap(clip_selection_method: Literal['auto', 'manual'], 
                   clip_length: int, 
                   sub_alignment: Literal['left', 'center', 'right'], 
                   include_intro: bool, 
                   overlay_intro_image: bool, 
                   fullscreen_intro_image: bool, 
                   intro_image_duration: int,
                   font_file: Union[str, None] = None
                   ) -> None:
    wb = openpyxl.load_workbook('video_data.xlsx')
    ws = wb.active

    print('Importing data from spreadsheet...')
    ids, id_cells = get_video_filenames(ws)

    print('Extracting clips...')
    video_clips = extract_clips(ws, ids, id_cells, clip_selection_method, clip_length)

    if overlay_intro_image:
        start_clip_idx = 1
    else:
        start_clip_idx = 0
    
    print('Resizing clips...')
    resized_clips = resize_clips(video_clips, start_clip_idx, intro_image_duration, fullscreen_intro_image)

    print('Adding crossfade to clips...')
    custom_padding = 1
    faded_clips = add_crossfade(resized_clips, custom_padding)
    
    print('Concatenating clips...')
    video = CompositeVideoClip(faded_clips)

    print('Generating subtitles...')
    subtitles, start_sub_idx = generate_subtitles(ws, id_cells, video_clips, custom_padding, include_intro, font_file)

    # Align subtitles according to user input.
    print('Adding subtitles...')
    recap = CompositeVideoClip([video] + 
                               [sub.with_position(('center','center')) for sub in subtitles[:start_sub_idx]] +    # Align intro text in centre
                               [sub.with_position((f'{sub_alignment}','bottom')) for sub in subtitles[start_sub_idx:]])

    # Save recap video file.
    print('Saving recap...')
    recap.write_videofile('recap.mp4')

def get_video_filenames(ws):
    # Get video filenames from spreadsheet.
    id_cells = [cell for cell in ws['A'] if cell.value != 'FILENAME' and cell.value != None and cell.value != '']
    ids = [str(cell.value) for cell in ws['A'] if cell.value != 'FILENAME' and cell.value != None and cell.value != '']
    return ids, id_cells

def extract_clips(ws, ids, id_cells, clip_selection_method, clip_length):
    if clip_selection_method == 'manual':
        # Pick out the specified clips from the files and normalise their audio.
        video_clips = []
        for i in range(len(ids)):
            video_clips.append(VideoFileClip(str(Path('Videos', f'{ids[i]}'))).subclipped(str(ws[f'B{id_cells[i].row}'].value), str(ws[f'C{id_cells[i].row}'].value)).with_effects([afx.AudioNormalize()]))
    elif clip_selection_method == 'auto':
        chorus_error = False
        missing_choruses = []
        video_clips = []
        for i in range(len(ids)):
            print(f'Extracting clip {i+1} of {len(ids)}')
            # Use manual clip if it is specified in spreadsheet.
            if ws[f'B{id_cells[i].row}'].value and ws[f'C{id_cells[i].row}'].value:
                video_clips.append(VideoFileClip(str(Path('Videos', f'{ids[i]}'))).subclipped(str(ws[f'B{id_cells[i].row}'].value), str(ws[f'C{id_cells[i].row}'].value)).with_effects([afx.AudioNormalize()]))
            # Otherwise select clip automatically via chorus detection.
            else:
                chorus_start = find_and_output_chorus(str(Path('Videos', f'{ids[i]}')), None)
                try:
                    video_clips.append(VideoFileClip(str(Path('Videos', f'{ids[i]}'))).subclipped(chorus_start, chorus_start + clip_length).with_effects([afx.AudioNormalize()]))
                except TypeError:
                    chorus_error = True
                    print(f'No chorus found for video {ids[i]}. Please choose a clip manually.')
                    missing_choruses.append(ids[i])
        if chorus_error:
            raise TypeError(f'Auto-generation failed for some clips. Please choose clips manually for the videos specified below then try again.\n{missing_choruses}')
    
    return video_clips

def resize_clips(video_clips, start_clip_idx, intro_image_duration, fullscreen_intro_image):
    # Resize clips while maintaining aspect ratio and then add black borders if necessary to reach 1920x1080p.
    # Also add intro image overlay.
    resized_clips = []

    for clip in video_clips[:start_clip_idx]:
        img_clip = ImageClip('intro.png', duration=intro_image_duration)
        if fullscreen_intro_image:
            img_size_ratio = 1.0
        else:
            img_size_ratio = 0.5
        
        if clip.w / clip.h <= 1920 / 1080 and img_clip.w / img_clip.h <= 1920 / 1080:
            resized_clips.append(CompositeVideoClip([ImageClip('1920x1080-black.jpg', duration=clip.duration).resized((1920, 1080)), 
                                                clip.resized(height=1080).with_position('center', 'center'),
                                                img_clip.resized(height=round(1080*img_size_ratio)).with_position('center', 'center')]))
        elif clip.w / clip.h > 1920 / 1080 and img_clip.w / img_clip.h <= 1920 / 1080:
            resized_clips.append(CompositeVideoClip([ImageClip('1920x1080-black.jpg', duration=clip.duration).resized((1920, 1080)), 
                                                clip.resized(width=1920).with_position('center', 'center'),
                                                img_clip.resized(height=round(1080*img_size_ratio)).with_position('center', 'center')])) 
        elif clip.w / clip.h <= 1920 / 1080 and img_clip.w / img_clip.h > 1920 / 1080:
            resized_clips.append(CompositeVideoClip([ImageClip('1920x1080-black.jpg', duration=clip.duration).resized((1920, 1080)), 
                                                clip.resized(height=1080).with_position('center', 'center'),
                                                img_clip.resized(width=round(1920*img_size_ratio)).with_position('center', 'center')])) 
        else: 
            resized_clips.append(CompositeVideoClip([ImageClip('1920x1080-black.jpg', duration=clip.duration), 
                                              clip.resized(width=1920).with_position('center', 'center'),
                                              img_clip.resized(width=round(1920*img_size_ratio)).with_position('center', 'center')]))

    resized_clips += [CompositeVideoClip([ImageClip('1920x1080-black.jpg', duration=clip.duration).resized((1920, 1080)), 
                                         clip.resized(height=1080).with_position('center', 'center')]) 
                      if clip.w / clip.h <= 1920 / 1080 
                      else CompositeVideoClip([ImageClip('1920x1080-black.jpg', duration=clip.duration), 
                                               clip.resized(width=1920).with_position('center', 'center')]) 
                      for clip in video_clips[start_clip_idx:]
                     ]
    
    return resized_clips

def add_crossfade(resized_clips, custom_padding):
    # Add crossfade to clips.
    faded_clips = [resized_clips[0].with_effects([afx.AudioFadeIn(custom_padding), afx.AudioFadeOut(custom_padding), vfx.FadeIn(custom_padding)])]
    idx = resized_clips[0].duration - custom_padding
    for i in range(len(resized_clips[1:])):
        clip = resized_clips[i+1].with_effects([afx.AudioFadeIn(custom_padding), afx.AudioFadeOut(custom_padding)])
        faded_clips.append(clip.with_start(idx).with_effects([vfx.CrossFadeIn(custom_padding)]))
        idx += clip.duration - custom_padding
    # Fade the final clip out to black.
    faded_clips[-1] = faded_clips[-1].with_effects([vfx.FadeOut(custom_padding)])

    return faded_clips

def generate_subtitles(ws, id_cells, video_clips, custom_padding, include_intro, font_file):
    subtitles = []
    start_clip_idx = 0
    start_sub_idx = 0

    if include_intro:
        start_clip_idx = 1
       
        intro_txt = ''
        if ws[f'D{id_cells[0].row}'].value:
            intro_txt += str(ws[f'D{id_cells[0].row}'].value)
        if ws[f'E{id_cells[0].row}'].value:
            intro_txt += '\n' + str(ws[f'E{id_cells[0].row}'].value)
        if ws[f'F{id_cells[0].row}'].value:
            intro_txt += '\n' + str(ws[f'F{id_cells[0].row}'].value)
        
        generator = lambda txt: TextClip(text=txt, font=font_file, font_size=150, color='white', stroke_color='black', stroke_width=3)
        intro_sub = [
                    ((custom_padding, # Start time of subtitle
                    int(video_clips[0].duration - custom_padding)), # End time of subtitle
                    intro_txt) # Text of subtitle
                    ]
        subtitles.append(SubtitlesClip(intro_sub, make_textclip=generator, encoding='utf-8'))

        start_sub_idx = len(subtitles)

    generator = lambda txt: TextClip(text=txt, font=font_file, font_size=50, color='white', stroke_color='black', stroke_width=3, margin=(10,20))
    subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
            int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
            str(ws[f'D{id_cells[video_clips.index(clip)].row}'].value) + '\n' + 
            str(ws[f'E{id_cells[video_clips.index(clip)].row}'].value) + '\n' + 
            str(ws[f'F{id_cells[video_clips.index(clip)].row}'].value)) # Text of subtitle
            for clip in video_clips[start_clip_idx:]]
    if subs:
        subtitles.append(SubtitlesClip(subs, make_textclip=generator, encoding='utf-8'))

    return subtitles, start_sub_idx


if __name__ == '__main__':
    clip_selection_method = ''
    while clip_selection_method not in ('auto', 'manual'):
        print('Please enter your desired method of clip selection (auto/manual)')
        clip_selection_method = input().lower()
        if clip_selection_method == 'automatic':
            clip_selection_method = 'auto'

    clip_length = '0'
    if clip_selection_method == 'auto':
        while not clip_length.isnumeric() or int(float(clip_length)) <= 0:
            print('Please enter your desired clip length in seconds')
            clip_length = input()
    clip_length = int(float(clip_length))

    sub_alignment = ''
    while sub_alignment not in ('left', 'center', 'right'):
        print('Please enter your desired subtitle alignment (left/center/right)')
        sub_alignment = input().lower()
        if sub_alignment == 'centre':
            sub_alignment = 'center'

    include_intro = ''
    while include_intro not in (True, False):
        print('Are you including an intro clip? (y/n)')
        include_intro = input().lower()
        if include_intro == 'y':
            include_intro = True
        elif include_intro == 'n':
            include_intro = False

    overlay_intro_image = ''
    if include_intro:
        while overlay_intro_image not in (True, False):
            print('Do you want to overlay a custom image on the intro clip? (y/n)')
            overlay_intro_image = input().lower()
            if overlay_intro_image == 'y':
                overlay_intro_image = True
            elif overlay_intro_image == 'n':
                overlay_intro_image = False

    fullscreen_intro_image = ''
    if overlay_intro_image:
        while fullscreen_intro_image not in (True, False):
            print('Make custom image fullscreen? (y/n)')
            fullscreen_intro_image = input().lower()
            if fullscreen_intro_image == 'y':
                fullscreen_intro_image = True
            elif fullscreen_intro_image == 'n':
                fullscreen_intro_image = False

    intro_image_duration = '0'
    if overlay_intro_image:
        while not intro_image_duration.isnumeric() or int(float(intro_image_duration)) <= 0 or int(float(intro_image_duration)) > clip_length:
            print('Please enter your desired duration of the custom image in seconds.')
            intro_image_duration = input()
    intro_image_duration = int(float(intro_image_duration))
    
    use_custom_font = False
    font_file = '🎬'
    while use_custom_font not in ('y', 'n'):
        print('Would you like to use a custom font? (y/n)')
        use_custom_font = input().lower()
        if use_custom_font == 'y':
            while not Path(font_file).exists():
                print('Please enter the path to your font file.')
                font_file = input()
        elif use_custom_font == 'n':
            font_file = 'Fonts/LiberationSans-Regular.ttf'

    generate_recap(clip_selection_method, 
                clip_length, 
                sub_alignment, 
                include_intro, 
                overlay_intro_image, 
                fullscreen_intro_image, 
                intro_image_duration,
                font_file=font_file)