# This code is written for usage in subject TPR2251 Pattern Recognition. 
# You may modify the parameters to suit your needs
# This is not a professional code by any means, it is just a random code to help fulfilling my assignment more easily. 

# See https://stackoverflow.com/a/39501288 for explanation on getting the creation time. 
# See https://pillow.readthedocs.io/en/stable/ for documentation of the PIL library

import os
import platform
import time
from PIL import Image, ImageDraw, ImageFont

def timestamp_photo(input_photo, output_photo, text, pos):
    """
    The timestamp_photo function take four arguments, and will return <input_photo> with <text> at <pos> and saved it as <output_photo>. 
    
    You may modify the font family, font size and the font color according to your need. 
    """
    photo = Image.open(input_photo)

    # make the image editable
    drawing = ImageDraw.Draw(photo)

    font = ImageFont.truetype(r"C:\Windows\Fonts\Arial.ttf", 60) # You will need to change this path if you are not using Windows
    drawing.text(pos, text, fill = "black", font = font)
    photo.save(output_photo)

def photo_created_time(filename):
    """
    The created_time function is used to return the creation time of the files for different OSes. 
    
    For Linux, unfortunately it seems like there is no easy way to extract the creation time, the last modified time is used instead. 
    """
    if platform.system() == 'Windows': # for Windows User
        return os.path.getmtime(filename) # You may change this to getctime(filename) depending on the attributes of your images
    else:
        filestat = os.stat(filename) # getting all the stats 
        try: # for Mac User
            return filestat.st_birthtime
        except AttributeError: # for Linux
            return filestat.st_mtime


count = 1
for photo in os.listdir("folder"): # change "folder" to the path to your folder
    path = "source/" + photo
    created_time = photo_created_time(path)
    formatted_created_time = time.ctime(created_time)
    output_name = "destination/" + str(count) + ".png"
    
    timestamp_photo(path, output_name, str(formatted_created_time), pos=(5, 5))
    count += 1
    
print("All timestamp added. ")
