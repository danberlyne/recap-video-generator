#!/usr/bin/env python3
# recap_generator.py - Extracts clips from video files and combines them into a single recap video.

import openpyxl
from pathlib import Path
from moviepy.editor import *
import moviepy.audio.fx.all as afx
from moviepy.video.tools.subtitles import SubtitlesClip

def main():
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
    faded_clips = [resized_clips[0]]
    idx = resized_clips[0].duration - custom_padding
    for clip in resized_clips[1:]:
        faded_clips.append(clip.set_start(idx).crossfadein(custom_padding))
        idx += clip.duration - custom_padding
    video = CompositeVideoClip(faded_clips)

    # Add subtitles.
    generator = lambda txt: TextClip(txt, font='Arial', fontsize=50, color='white')
    subs = [((int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)])) + custom_padding, # Start time of subtitle
            int(sum(previous_clip.duration - custom_padding for previous_clip in video_clips[:video_clips.index(clip)+1]))), # End time of subtitle
            str(ws[f'D{id_cells[video_clips.index(clip)].row}'].value) + '\n' + str(ws[f'E{id_cells[video_clips.index(clip)].row}'].value) + '\n' + str(ws[f'F{id_cells[video_clips.index(clip)].row}'].value)) # Text of subtitle
            for clip in video_clips]
    subtitles = SubtitlesClip(subs, generator)
    recap = CompositeVideoClip([video, subtitles.set_pos(('center','bottom'))])

    # Save recap video file.
    recap.write_videofile('recap.mp4')

if __name__ == '__main__':
    main()