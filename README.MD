# Extract Media PowerPoint

![](https://travis-ci.org/madyel/extract_media_ppt.svg?branch=main) [![GitHub license](https://img.shields.io/github/license/madyel/extract_media_ppt)](https://github.com/madyel/extract_media_ppt/blob/main/LICENSE.txt) ![GitHub last commit](https://img.shields.io/github/last-commit/madyel/extract_media_ppt)

Example:

```python
from extract import PowerPoint
'''
DEFAULT:

default_permission_video = ['mp4', 'avi', 'mpg', 'mpeg', 'wmv']
default_permission_image = ['png', 'jpeg', 'jpg', 'bmp', 'svg']
output_directory = 'temp'
'''

if __name__ == '__main__':
    pptx = 'filename.pptx'
    ppt = PowerPoint(pptx, 'video', 'temp, 'wmv','mp4','avi')
#or
    #ppt = PowerPoint(pptx, 'video')
#or
    #ppt = PowerPoint(pptx, 'image', 'temp, 'svg','png','jpg')
#or
    #ppt = PowerPoint(pptx, 'image')
#extract all media audio,video and image
    ppt.extractAllMedia() 
    ppt.extract()
```

### Install Requires

```shell script
pip install python-pptx
```

```python
              .:.               
             .::::.             
..         ..::::::''::         
::::..  .::''''':::    ''.      
':::::::'         '.  ..  '.    
 ::::::'            : '::   :   
  :::::     .        : ':'   :  
  :::::    :::       :.     .'. 
 .::::::    ':'     .' '.:::: : 
 ::::::::.         .    ::::: : 
:::::    '':.... ''      '''' : 
':::: .:'              ...'' :  
 ..::.   '.........:::::'   :   
  '':::.   '::'':'':::'   .'    
        '..  ''.....'  ..'      
           ''........''

```