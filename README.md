# Video to XLSX

This is a short tutorial of converting a video to an Excel workbook, as an effort to reproduce the results of [this viral video](https://www.facebook.com/watch/?v=1963929290331098) from a while ago.

Script tested on Apple M1, macOS Sonoma 14.0, Office 365 whatever version as of Oct 2023. 
Output XLSX file works on Office 365 for Mac, as well as Office 2010 on a Windows 7 VM (and it plays much faster than realtime!!)

Sample workbook of the first 256 frames of [Mirai wa Kaze no You ni](https://www.youtube.com/watch?v=l6t2PFGRgbY) is available in the Releases section.

## What you need
- How to use the commandline
- ffmpeg
- Python 3.10+
- This repository
- Microsoft Excel. LibreOffice is not compatible with this script, since internally it seems to convert OOXML to ODT upon opening the file and the process is very very very slow (it's single threaded by default). You could try it but I tried and gave up. You can try porting the code to write directly to ODT instead of XLSX

This tutorial expects that you are already a technical user. Read the instruction carefully. If you find the instruction not clear, read the code. It's quite short ;)

All commands in this section are written with Unix in mind. If you use Windows, convert the commands yourself.

## Backstory

<details>
  <summary>Expand</summary>
Back in 2018, I came across a video of someone using Excel to play the anime opening "Only My Railgun". I thought it was such a cool, outside-the-box idea! Being a total tech geek, I tried to recreate it myself with C#.NET and the Microsoft.Office.Interop API to fill each cell with the color of the matching pixel. But it was way too slow to actually work. Parallel programming? As if! I was just a high school student and didn't know nothing about that advanced stuff. Actually, the approach in this repo still does not use concurrency because the openpyxl API is not thread-safe per workbook. It could work regardless, but I have not tried.  


The major issue was that if you fill in the colors as-is, Excel will error out and say the file is corrupted. This is because of a [known limitation](https://learn.microsoft.com/en-us/office/troubleshoot/excel/too-many-different-cell-formats-in-excel) where Excel workbooks can only have 64,000 unique cell formatting combinations. With each frame being 160x90 pixels, that's 14,400 possible color combos per sheet. So you can only fit like 4.44 frames before hitting the limit. That lined up with what I saw - I could only fill 4 or 5 sheets before corrupting the workbook.  


Several days ago, I revisited this idea after a few years. I have come to know a lot more about how stuff work under the hood, as well as the dirty techniques to preprocess various types of data, including images. The hidden weapon is color quantization. This knocks down the number of color combinations to just 128 (as in this tutorial).
</details>

## Installation

1. Make sure you fullfill the requirements described above
2. Create a Python virtualenv
   ```bash
   python -m virtualenv /path/to/venv
   source /path/to/venv/bin/activate
   ```
3. Install the dependencies
   ```bash
    python -m pip install -r requirements.txt
   ```

## Steps

The config suggested in this section is to "play" the video in Excel _in real time_, without the need of screen-recording it and speed it up later. If this is not your case, you can opt for a higher framerate and resolution.

1. Download the video you need to convert and put it somewhere
2. Extract the frames to some directory
   ```bash
    ffmpeg -i <path-to-input-video> \
        -vf "scale=<output-resolution>" \
        -r <fps> \
        -q:v 2 \
        "<output-directory>/%04d.jpg"
   ```
   - Where `<output-resolution>` is in the format like `1920x1080`. The resolution should be small, very small, like `190x60` or even lower. I tested with `200x112` and it "played" fine on my Air M1
   - `<fps>` is the fps of the output images. It should be a number such that `1 / <fps>` is rational (you will know why later). Something like 5, 10, 12 is fine. Higher fps leads to more images and more processing time, and can harm your "playback" performance
   - `q:v 2` is the parameter to control the resulting jpeg quality. You can leave it as-is
   - Output filename should be in the format of `%04d.jpg` (four digits padded by zeros like `0001, 0002, 0003, ..., 9999`), and must starts at `0001`. Nothing else. This is hard-coded in the script to make sure the frames follow the right order
3. Run the script
   ```bash
   source /path/to/venv/bin/activate
   python img2xlsx.py \
       --frames-dir /path/to/frames/directory \
       --output /path/to/output/file.xlsx
   ```
   You can run `python img2xlsx.py --help` to see more customizable params or use `--head N` to only generate the first N frames. It's quite limited. If you need more customization just modify the code instead.
   This script does the following tasks:
   - Load each image and quantize the colors to the [X11 256-color palette](https://www.ditig.com/256-colors-cheat-sheet)
   - Write the image pixel-by-pixel into each sheet using Pillow and openpyxl

## Example end-to-end workflow

```bash
# Install
python -m venv /Users/me/envs/img2xlsx
source /Users/me/envs/img2xlsx/activate
pip install -r requirements.txt

# Convert video to frames
mkdir -p /Users/me/Pictures/the-frames
ffmpeg -i /Users/me/Downloads/some-video.mp4 \
   -vf "scale=200x112" \
   -r 10 \
   -q:v 2 \
   "/Users/me/Pictures/the-frames/%04d.jpg"

# Convert frames to excel
source /Users/me/envs/
python img2xlsx.py \
      --frames-dir /Users/me/Pictures/the-frames \
      --output output.xlsx \
      --head 20 # only generate first 20 frames
```

The frame to excel conversion process is quite slow since it's O(n^2) without any parallelism. With the above config, it runs at around 1.4 it/s on my M1 Air.

## "Playing" the "video"

Congrats, you now have a "video" in Excel ready to be played. You should notice that the color is _very_ washed out, and not accurate at all. That is due to the image quantization process to overcome Excel's formatting limitation discussed in the [Backstory](#backstory) section. Any MR to overcome this using another method is welcomed (such as using adaptive quantization on the whole dataset).

To "play" it, you have two (actually three) options:
- Use VBA macro
- Use OfficeScript
- Manually press Ctrl+PgDown to switch to the next sheet ;). On non-fullsize Mac keyboards it's Control+Fn+DownArrow

I highly recommend OfficeScript since it's pretty easy to write (since it's literally TypeScript) and I only tested that option along with the last one. No VBA `Sleep` approach worked with my Mac setup so far.

The script that I used:

```ts
const INV_FPS = 200;
const sleep = (delay: number) => new Promise((resolve) => setTimeout(resolve, delay));

async function main(workbook: ExcelScript.Workbook) {
    let sheets = workbook.getWorksheets();

    for (let i = 0; i < sheets.length; i++) {
        sheets[i].activate();
        await sleep(INV_FPS);
    }
}
```

Set `INV_FPS` to `1/<fps>*1000`, where `<fps>` is the fps you used during the frame extraction process. If your fps is 5, `INV_FPS` should be 200. This is the reason why I suggested choosing an fps such that its inversion is rational. If not you might experience some difficulties when aligning the music in the editing process later (if you wish).

## Slow playback
Your spreadsheets playing too slow? Try:
- Decreasing the video resolution
- Decreate the fps
- Use a simpler color palette that has less than 256 colors (if your video is simple enough). This requires modifying the code, however

## Other approaches
Some of you might have come up with another approach to do this after reading the tutorial. One of them might be writing the frames vertically in one sheet, instead of creating multiple sheets. I tried and it worked, but it was very slow to open and the performance of scrolling down was much worse than switching sheets.
