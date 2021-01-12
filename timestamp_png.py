# This code is written for usage in subject TPR2251 Pattern Recognition only. 
# You may modify the parameters to suit your needs

# See https://stackoverflow.com/a/39501288 for explanation on getting the creation time. 
# See https://pillow.readthedocs.io/en/stable/ for documentation of the PIL library

import os
import time
import pillow

def timestamp_photo(input_photo, output_photo, text, pos):
	"""
	The timestamp_photo function take four arguments, and will return <input_photo> with <text> at <pos> and saved it as <output_photo>. 
	
	You may modify the font family, font size and the font color according to your need. 
	"""
    photo = Image.open(input_photo)

    # make the image editable
    drawing = ImageDraw.Draw(photo)

    font = ImageFont.truetype("Pillow/Tests/fonts/FreeMono.ttf", 12)
    drawing.text(pos, text, fill = "black", font = font)
    photo.save(output_photo)
    
def created_time(filename):
	"""
	The created_time function is used to return the creation time of the files for different OSes. 
	
	For Linux users, unfortunately it seems like there is no easy way to extract the creation time, instead, the last modified time is used. 
	"""
	if platform.system() == 'Windows': # for Windows User
		return os.path.getctime(filename)
	else:
		filestat = os.stat(filename) # getting all the stats 
		try: # for Mac User
			return filestat.st_birthtime
		except AttributeError: # for Linux User
			return filestat.st_mtime
    
for photo_name in os.listdir("folder"): # change "folder" to the path to your folder
	created_time = photo_created_time(photo_name)
	output_name = created_time + ".png"
	
    timestamp_photo(photo_name, output_name, created_time, pos=(5, 5))
    
print("All timestamp added. ")
