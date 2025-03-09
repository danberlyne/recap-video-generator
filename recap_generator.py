#!/usr/bin/env python3
# recap_generator.py - Extracts clips from video files and combines them into a single recap video.

import openpyxl
from pathlib import Path
from moviepy.editor import VideoFileClip, CompositeVideoClip, ImageClip, TextClip
import moviepy.audio.fx.all as afx
import moviepy.video.fx.all as vfx
from moviepy.video.tools.subtitles import SubtitlesClip
from pychorus import find_and_output_chorus

# `alignment` is the alignment of subtitles in the video (left/center/right).
# `clip_selection` determines whether to use `pychorus` to detect clips (auto) or whether to use user-defined clips (manual).
def main(clip_selection, clip_length, alignment):
    # Get filenames from spreadsheet.
    print('Importing data from spreadsheet...')
    wb = openpyxl.load_workbook('video_data.xlsx')
    ws = wb.active
    id_cells = [cell for cell in ws['A'] if cell.value != 'FILENAME' and cell.value != None and cell.value != '']
    ids = [str(cell.value) for cell in ws['A'] if cell.value != 'FILENAME' and cell.value != None and cell.value != '']

    print('Extracting clips...')
    if clip_selection == 'manual':
        # Pick out the specified clips from the files and normalise their audio.
        video_clips = []
        for i in range(len(ids)):
            video_clips.append(VideoFileClip(str(Path('Videos', f'{ids[i]}'))).subclip(str(ws[f'B{id_cells[i].row}'].value), str(ws[f'C{id_cells[i].row}'].value)).fx(afx.audio_normalize))
    elif clip_selection == 'auto':
        chorus_error = False
        missing_choruses = []
        video_clips = []
        for i in range(len(ids)):
            print(f'Extracting clip {i+1} of {len(ids)}')
            # Use manual clip if it is specified in spreadsheet.
            if ws[f'B{id_cells[i].row}'].value and ws[f'C{id_cells[i].row}'].value:
                video_clips.append(VideoFileClip(str(Path('Videos', f'{ids[i]}'))).subclip(str(ws[f'B{id_cells[i].row}'].value), str(ws[f'C{id_cells[i].row}'].value)).fx(afx.audio_normalize))
            # Otherwise select clip automatically via chorus detection.
            else:
                chorus_start = find_and_output_chorus(str(Path('Videos', f'{ids[i]}')), None)
                try:
                    video_clips.append(VideoFileClip(str(Path('Videos', f'{ids[i]}'))).subclip(chorus_start, chorus_start + clip_length).fx(afx.audio_normalize))
                except TypeError:
                    chorus_error = True
                    print(f'No chorus found for video {ids[i]}. Please choose a clip manually.')
                    missing_choruses.append(ids[i])
        if chorus_error:
            raise TypeError(f'Auto-generation failed for some clips. Please choose clips manually for the videos specified below then try again.\n{missing_choruses}')

    print('Resizing clips...')
    # Resize clips while maintaining aspect ratio and then add black borders if necessary to reach 1920x1080p.
    resized_clips = [CompositeVideoClip([ImageClip('1920x1080-black.jpg', duration=clip.duration).resize((1920, 1080)), clip.resize(height=1080).set_position('center', 'center')]) if clip.w / clip.h <= 1920 / 1080 else CompositeVideoClip([ImageClip('1920x1080-black.jpg', duration=clip.duration), clip.resize(width=1920).set_position('center', 'center')]) for clip in video_clips]

    print('Adding crossfade to clips...')
    # Concatenate clips with crossfade.
    custom_padding = 1
    faded_clips = [vfx.fadein(afx.audio_fadeout(afx.audio_fadein(resized_clips[0], custom_padding), custom_padding), custom_padding)]
    idx = resized_clips[0].duration - custom_padding
    for i in range(len(resized_clips[1:])):
        clip = resized_clips[i+1]
        clip = afx.audio_fadein(clip, custom_padding)
        clip = afx.audio_fadeout(clip, custom_padding)
        faded_clips.append(clip.set_start(idx).crossfadein(custom_padding))
        idx += clip.duration - custom_padding
    # Fade the final clip out to black.
    faded_clips[-1] = vfx.fadeout(faded_clips[-1], custom_padding)
    print('Concatenating clips...')
    video = CompositeVideoClip(faded_clips)

    # Add subtitles.
    print('Generating subtitles...')
    subtitles = []
    generator = lambda txt: TextClip(txt, font='Arial', fontsize=50, color='white')
    if alignment == 'center':
        subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
                int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
                str(ws[f'D{id_cells[video_clips.index(clip)].row}'].value) + '\n' + ' ' + '\n' + ' ') # Text of subtitle
                for clip in video_clips]
        subtitles.append(SubtitlesClip(subs, generator))

        generator = lambda txt: TextClip(txt, font='Arial', fontsize=50, color='white')
        subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
                int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
                ' ' + '\n' + str(ws[f'E{id_cells[video_clips.index(clip)].row}'].value) + '\n' + ' ') # Text of subtitle
                for clip in video_clips]
        subtitles.append(SubtitlesClip(subs, generator))

        generator = lambda txt: TextClip(txt, font='Arial', fontsize=50, color='white')
        subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
                int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
                ' ' + '\n' + ' ' + '\n' + str(ws[f'F{id_cells[video_clips.index(clip)].row}'].value)) # Text of subtitle
                for clip in video_clips]
        subtitles.append(SubtitlesClip(subs, generator))
    elif alignment == 'left':
        subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
                int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
                ' ' + str(ws[f'D{id_cells[video_clips.index(clip)].row}'].value) + '\n' + ' ' + '\n' + ' ') # Text of subtitle
                for clip in video_clips]
        subtitles.append(SubtitlesClip(subs, generator))

        generator = lambda txt: TextClip(txt, font='Arial', fontsize=50, color='white')
        subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
                int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
                ' ' + '\n' + ' ' + str(ws[f'E{id_cells[video_clips.index(clip)].row}'].value) + '\n' + ' ') # Text of subtitle
                for clip in video_clips]
        subtitles.append(SubtitlesClip(subs, generator))

        generator = lambda txt: TextClip(txt, font='Arial', fontsize=50, color='white')
        subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
                int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
                ' ' + '\n' + ' ' + '\n' + ' ' + str(ws[f'F{id_cells[video_clips.index(clip)].row}'].value)) # Text of subtitle
                for clip in video_clips]
        subtitles.append(SubtitlesClip(subs, generator))
    elif alignment == 'right':
        subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
                int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
                str(ws[f'D{id_cells[video_clips.index(clip)].row}'].value) + '  ' + '\n' + ' ' + '\n' + ' ') # Text of subtitle
                for clip in video_clips]
        subtitles.append(SubtitlesClip(subs, generator))

        generator = lambda txt: TextClip(txt, font='Arial', fontsize=50, color='white')
        subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
                int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
                ' ' + '\n' + str(ws[f'E{id_cells[video_clips.index(clip)].row}'].value) + ' ' + '\n' + ' ') # Text of subtitle
                for clip in video_clips]
        subtitles.append(SubtitlesClip(subs, generator))

        generator = lambda txt: TextClip(txt, font='Arial', fontsize=50, color='white')
        subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
                int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
                ' ' + '\n' + ' ' + '\n' + str(ws[f'F{id_cells[video_clips.index(clip)].row}'].value) + ' ') # Text of subtitle
                for clip in video_clips]
        subtitles.append(SubtitlesClip(subs, generator))

    # Align subtitles according to user input.
    print('Adding subtitles...')
    recap = CompositeVideoClip([video] + [sub.set_pos((f'{alignment}','bottom')) for sub in subtitles])

    # Save recap video file.
    print('Saving recap...')
    recap.write_videofile('recap.mp4')

if __name__ == '__main__':
    clip_selection = ''
    while clip_selection not in ('auto', 'manual'):
        print('Please enter your desired method of clip selection (auto/manual)')
        clip_selection = input().lower()
        if clip_selection == 'automatic':
            clip_selection = 'auto'

    clip_length = '0'
    if clip_selection == 'auto':
        while not clip_length.isnumeric() or int(float(clip_length)) <= 0:
            print('Please enter your desired clip length in seconds')
            clip_length = input()
    clip_length = int(float(clip_length))

    alignment = ''
    while alignment not in ('left', 'center', 'right'):
        print('Please enter your desired subtitle alignment (left/center/right)')
        alignment = input().lower()
        if alignment == 'centre':
            alignment = 'center'
    
    main(clip_selection, clip_length, alignment)