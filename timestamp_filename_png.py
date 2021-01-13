# This code is written for usage in subject TPR2251 Pattern Recognition. 
# You may modify the parameters to suit your needs
# This code is a modified version of timestamp_png.py, to extract the creation time from the filename

# See https://stackoverflow.com/a/39501288 for explanation on getting the creation time. 
# See https://pillow.readthedocs.io/en/stable/ for documentation of the PIL library

import os
import time
import re
from PIL import Image, ImageDraw, ImageFont

def timestamp_photo(input_photo, output_photo, text, pos):
    """
    The timestamp_photo function take four arguments, and will return <input_photo> with <text> at <pos> and saved it as <output_photo>. 
    
    You may modify the font family, font size and the font color according to your need. 
    """
    photo = Image.open(input_photo)

    # make the image editable
    drawing = ImageDraw.Draw(photo)

    font = ImageFont.truetype(r"C:\Windows\Fonts\Arial.ttf", 60)
    drawing.text(pos, text, fill = "black", font = font) # change "fill" to white for images with dark background
    photo.save(output_photo)

def photo_created_time(filename):
    """
    The created_time function is used to return the creation time of the files based on the filename. 
    
    My filename format: "IMG_YYYYMMDD_HHMMSS.jpg"
    """
    basename = os.path.basename(filename)
    formatted_basename = basename.split('_')
    datetime = formatted_basename[1] + " " + formatted_basename[2][:-4] # [:-4] to exclude the file extension
    return datetime


count = 1
for photo in os.listdir("Train Set"): # change "folder" to the path to your folder
    path = "Train Set/" + photo
    created_time = photo_created_time(path)
    output_name = "Stamped/" + str(count) + ".png"
    
    timestamp_photo(path, output_name, str(created_time), pos=(5, 5))
    count += 1
    
print("All timestamp added. ")