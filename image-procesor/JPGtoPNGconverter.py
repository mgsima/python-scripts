import sys
import os
from PIL import Image

# Grab the first and second argument
try:
    pokedex_path = sys.argv[1]
    new_path = sys.argv[2]
except: print("Not arguments given")

# check if \new exist
if os.path.exists(new_path):
    pass
else: 
    os.makedirs(new_path)

# loop throw the pokedex
for file in os.listdir(pokedex_path):
    img_path = os.path.join(pokedex_path, file)
    name_file, ext = os.path.splitext(file)
    new_png = name_file + '.png'

    # convert images
    if os.path.isfile(img_path):
        img = Image.open(img_path)
        # save it to the new folder
        png_route = os.path.join(new_path, new_png)
        img.save(png_route, "PNG")



