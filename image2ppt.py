#!/usr/local/bin/python
import sys, os
from pptx import Presentation
from pptx.util import Inches

INPUT_PATH = os.getcwd()
OUTPUT_PATH = os.getcwd()
PPT_NAME = "image2ppt"
SLIDE_WIDTH = 1920
SLIDE_HEIGHT = 1080
RESOLUTION = 72
PPT_NAME = "image2ppt"
SUPPORTED_FORMATS = ['.jpg', '.JPG', '.jpeg', '.JPEG', '.png', '.PNG' ]

def pixelToEmu(pixel, resolution):
    return int(int(pixel) * 914400 / int(resolution))

img_folder = raw_input("Enter path of image folder [Current directory]: ").strip() or INPUT_PATH
output_name = raw_input("Enter name of presentation [" + PPT_NAME + "]: ").strip() or PPT_NAME
output_path = raw_input("Enter output path [Current directory]: ").strip() or OUTPUT_PATH
width_pixel = raw_input("Image width in pixels [" + str(SLIDE_WIDTH) + "]: ") or SLIDE_WIDTH
height_pixel = raw_input("Image height in pixels [" + str(SLIDE_HEIGHT) + "]: ") or SLIDE_HEIGHT
img_resolution = raw_input("Image resolution in ppi [" + str(RESOLUTION) + "]:") or RESOLUTION

img_width = pixelToEmu(width_pixel, img_resolution)
img_height = pixelToEmu(height_pixel, img_resolution)

files = os.listdir(img_folder)
images = []
for f in files:
    path = os.path.join(img_folder, f)
    if os.path.isfile(path):
        filename, extension = os.path.splitext(f)
        if extension in SUPPORTED_FORMATS:
            images.append(path)

print(str(len(images)) + " images found")

prs = Presentation()
prs.slide_height = img_height
prs.slide_width = img_width

for image in images:
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    pic = slide.shapes.add_picture(image, Inches(0) , Inches(0), img_width,  img_height)
    print( "\"" + os.path.basename(image) + "\" is added")


prs.save(os.path.join(output_path, output_name + '.pptx'))
print(output_name + '.pptx is saved')
