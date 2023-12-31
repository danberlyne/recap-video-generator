#!/usr/bin/env python3
# recap_generator.py - Extracts clips from video files and combines them into a single recap video.

import openpyxl
from pathlib import Path
from moviepy.editor import *
import moviepy.audio.fx.all as afx
import moviepy.video.fx.all as vfx
from moviepy.video.tools.subtitles import SubtitlesClip

# `alignment` is the alignment of subtitles in the video (left/center/right).
def main(alignment):
    # Get filenames from spreadsheet.
    wb = openpyxl.load_workbook('video_data.xlsx')
    ws = wb.active
    id_cells = [cell for cell in ws['A'] if cell.value != 'FILENAME' and cell.value != None and cell.value != '']
    ids = [str(cell.value) for cell in ws['A'] if cell.value != 'FILENAME' and cell.value != None and cell.value != '']

    # Pick out the specified clips from the files and normalise their audio.
    video_clips = [VideoFileClip(str(Path('Videos', f'{id}'))).subclip(str(ws[f'B{id_cells[ids.index(id)].row}'].value), str(ws[f'C{id_cells[ids.index(id)].row}'].value)).fx(afx.audio_normalize) for id in ids]

    # Resize clips while maintaining aspect ratio and then add black borders if necessary to reach 1920x1080p.
    resized_clips = [CompositeVideoClip([ImageClip('1920x1080-black.jpg', duration=clip.duration).resize((1920, 1080)), clip.resize(height=1080).set_position('center', 'center')]) if clip.w / clip.h <= 1920 / 1080 else CompositeVideoClip([ImageClip('1920x1080-black.jpg', duration=clip.duration), clip.resize(width=1920).set_position('center', 'center')]) for clip in video_clips]

    # Concatenate clips with crossfade.
    custom_padding = 1
    faded_clips = [vfx.fadein(afx.audio_fadeout(afx.audio_fadein(resized_clips[0], custom_padding), custom_padding), custom_padding)]
    idx = resized_clips[0].duration - custom_padding
    for clip in resized_clips[1:]:
        clip = afx.audio_fadein(clip, custom_padding)
        clip = afx.audio_fadeout(clip, custom_padding)
        faded_clips.append(clip.set_start(idx).crossfadein(custom_padding))
        idx += clip.duration - custom_padding
    # Fade the final clip out to black.
    faded_clips[-1] = vfx.fadeout(faded_clips[-1], custom_padding)
    video = CompositeVideoClip(faded_clips)

    # Add subtitles.
    subtitles = []
    generator = lambda txt: TextClip(txt, font='Arial', fontsize=50, color='white')
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

    # Align subtitles according to user input.
    recap = CompositeVideoClip([video] + [sub.set_pos((f'{alignment}','bottom')) for sub in subtitles])

    # Save recap video file.
    recap.write_videofile('recap.mp4')

if __name__ == '__main__':
    alignment = ''
    while alignment not in ('left', 'center', 'right'):
        print('Please enter your desired subtitle alignment (left/center/right)')
        alignment = input().lower()
        if alignment == 'centre':
            alignment = 'center'
    main(alignment)