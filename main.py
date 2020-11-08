from extract import PowerPoint
import argparse
import os
#from pyfiglet import Figlet

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

class Hello(argparse.Action):
    def __call__(self, parser, namespace, values, option_string=None):
        print(parser)


path_output = 'temp/'
#f = Figlet(font='starwars')
#print(f.renderText('MaDyEl'))
title = f"{bcolors.WARNING}Extract media file of Power Point -  v0.3{bcolors.ENDC}"
print(title)
parser = argparse.ArgumentParser(description=title, add_help=False)
parser.add_argument('-h', '--help', '-help', action="help", help="Path file PowerPoint")
requiredNamed = parser.add_argument_group('required named arguments')
requiredNamed.add_argument('-p', '--pptx', help='Path file ppt (PowerPoint', required=True)
requiredNamed.add_argument('-o', '--output', help='Output media files', default='temp')
parser.add_argument('-a', '--all', help='Extract all media files', default=False)
parser.add_argument('-i', '--image', help='Extract all images files', default=False)
parser.add_argument('-m', '--video', help='Extract all movies files', default=False)
args = parser.parse_args()



pptx = args.pptx
output = args.output
all = args.all
image = args.image
video = args.video
type='video'
if image:
    type = 'image'

if not os.path.isfile(pptx):
    print("File not exist")
    exit()

if __name__ == '__main__':
    ppt = PowerPoint(pptx, type, output)
    if all:
        ppt.extractAllMedia()
    else:
        ppt.extract()